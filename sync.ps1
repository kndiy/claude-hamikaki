# sync.ps1 - mirror Claude-related text files from the projects tree into this
# repo, commit them on a NEW branch sync/<timestamp>, and push that branch for
# review. main only changes when you merge the pull request on GitHub.
# What gets mirrored is set in mirror.conf next to this script; the script itself
# is identical in every mirror repo (_claude-rnd4u, _claude-hamikaki).
# Windows PowerShell 5.1, ASCII only. Linux/macOS twin: sync.sh (keep both in sync).
#
# Usage (from anywhere):
#   powershell -ExecutionPolicy Bypass -File sync.ps1                 # sync + commit + push branch
#   powershell -ExecutionPolicy Bypass -File sync.ps1 -Remote <url>   # first time: set origin
#   ... -DryRun          list what would be mirrored, change nothing
#   ... -NoPush          sync + commit on the new branch, no push
#   ... -AllowSecrets    commit even if the secret scan finds something
#   ... -Message "text"  custom commit message
param(
    [string]$Remote = "",
    [string]$Message = "",
    [switch]$DryRun,
    [switch]$NoPush,
    [switch]$AllowSecrets
)
$ErrorActionPreference = "Stop"

# ---------- fixed rules (same in every mirror; keep in step with sync.sh) ----------
# A file is Claude-related when its name is one of these ...
$ClaudeNames = @("CLAUDE.md", "CLAUDE.local.md", ".claude-project")
# ... or it sits inside one of these folders.
$ClaudeDirs = @(".claude", "zClaudeDocs")
# Never mirror these names.
$ExcludeNames = @("settings.local.json")
# Text-only whitelist (lower case). Everything else (docx, xlsx, csv, images,
# pdf, zip, crt, extensionless data dumps) is skipped.
$TextExt = @(".md", ".txt", ".json", ".xml", ".py", ".groovy", ".java", ".sh",
             ".ps1", ".patch", ".diff", ".yml", ".yaml", ".sql", ".properties",
             ".cmd", ".command", ".bat", ".js")
# Files that must stay executable (git mode 755) for Linux/macOS.
$ExecExt = @(".sh", ".command")
$MaxBytes = 1MB
$SkipDirs = @(".git", "node_modules", "target", "build", "dist", "__pycache__", ".venv", "venv")
# High-confidence secret patterns (.NET regex).
$SecretPatterns = @(
    '-----BEGIN [A-Z ]*PRIVATE KEY-----',
    'gh[pousr]_[A-Za-z0-9]{36}',
    'github_pat_[A-Za-z0-9_]{20,}',
    'AKIA[0-9A-Z]{16}',
    'sk-ant-[A-Za-z0-9_-]{20,}',
    'xox[abprs]-[A-Za-z0-9-]{10,}',
    '"(password|passwd|pwd|secret|clientSecret|apiKey|api_key|accessToken)"\s*:\s*"[^"<>${}*]{6,}"'
)
# -----------------------------------------------------------------------------------

$Mirror = $PSScriptRoot
$Root = Split-Path -Parent $Mirror
# Files that belong to the mirror repo itself; everything else is regenerated each run.
$MirrorOwn = @("README.md", "mirror.conf", "sync.ps1", "sync.sh", "restore.ps1", "restore.sh", ".gitignore", ".gitattributes")
$Utf8 = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = $Utf8
[Console]::OutputEncoding = $Utf8

# mirror.conf sections, one item per line ('#' starts a comment):
#   [claude]   folders: only their Claude files (+ helper scripts, see [scripts])
#   [whole]    folders: every text file
#   [files]    single files at the projects root
#   [exclude]  folders skipped entirely
#   [scripts]  extensions of helper scripts to include inside [claude] folders,
#              but only outside nested git repos (code in a repo is already in git)
$Conf = @{ claude = @(); whole = @(); files = @(); exclude = @(); scripts = @() }
$confPath = Join-Path $Mirror "mirror.conf"
if (-not (Test-Path -LiteralPath $confPath)) { throw "Missing $confPath" }
$section = ""
foreach ($raw in [IO.File]::ReadAllLines($confPath)) {
    $line = ($raw -replace '#.*$', '').Trim()
    if (-not $line) { continue }
    if ($line -match '^\[(\w+)\]$') { $section = $Matches[1].ToLower(); if (-not $Conf.ContainsKey($section)) { throw "Unknown section [$section] in mirror.conf" }; continue }
    if (-not $section) { throw "mirror.conf: item outside a section: $line" }
    $line = $line.Replace("\", "/").TrimEnd("/")
    if ($line.StartsWith("/") -or $line.Contains("..")) { throw "mirror.conf: bad path: $line" }
    if ($section -eq "scripts") { foreach ($x in ($line -split '\s+')) { $Conf.scripts += ("." + $x.TrimStart(".").ToLower()) } }
    else { $Conf[$section] += $line }
}
$AllScopes = @($Conf.claude) + @($Conf.whole)
$RootFiles = @($Conf.files)

function ConvertTo-LongPath([string]$p) { if ($p.StartsWith("\\?\")) { $p } else { "\\?\" + $p } }

function Test-Excluded([string]$rel) {
    foreach ($x in $Conf.exclude) { if ($rel -eq $x -or $rel.StartsWith($x + "/")) { return $true } }
    return $false
}

function Test-Wanted([string]$rel, [string]$name, [string]$mode, [bool]$inRepo) {
    if ($ExcludeNames -contains $name) { return $false }
    $ext = [IO.Path]::GetExtension($name).ToLower()
    if ($mode -eq "whole") { return ($TextExt -contains $ext) }
    $isClaude = $ClaudeNames -contains $name
    if (-not $isClaude) {
        foreach ($d in $ClaudeDirs) { if (("/" + $rel) -like ("*/" + $d + "/*")) { $isClaude = $true; break } }
    }
    if ($isClaude) {
        if ($name -eq ".claude-project") { return $true }
        return ($TextExt -contains $ext)
    }
    return ((-not $inRepo) -and ($Conf.scripts -contains $ext))
}

$found = New-Object System.Collections.Generic.List[object]
$skippedBig = New-Object System.Collections.Generic.List[string]

function Walk([string]$dir, [string]$mode, [bool]$inRepo) {
    $relDir = ""
    if ($dir.Length -gt $Root.Length) { $relDir = $dir.Substring($Root.Length + 1).Replace("\", "/") }
    if ($relDir -and (Test-Excluded $relDir)) { return }
    if (-not $inRepo) { $inRepo = [IO.Directory]::Exists((ConvertTo-LongPath (Join-Path $dir ".git"))) -or [IO.File]::Exists((ConvertTo-LongPath (Join-Path $dir ".git"))) }
    foreach ($f in [IO.Directory]::EnumerateFiles((ConvertTo-LongPath $dir))) {
        $name = [IO.Path]::GetFileName($f)
        $rel = $f.Substring(4).Substring($Root.Length + 1).Replace("\", "/")
        if (Test-Wanted $rel $name $mode $inRepo) {
            $len = (New-Object IO.FileInfo($f)).Length
            if ($len -gt $MaxBytes) { $skippedBig.Add($rel) } else { $found.Add(@{ Src = $f; Rel = $rel }) }
        }
    }
    foreach ($d in [IO.Directory]::EnumerateDirectories((ConvertTo-LongPath $dir))) {
        $n = [IO.Path]::GetFileName($d)
        if ($SkipDirs -contains $n) { continue }
        if ([IO.File]::GetAttributes($d) -band [IO.FileAttributes]::ReparsePoint) { continue }
        Walk $d.Substring(4) $mode $inRepo
    }
}

foreach ($s in $AllScopes) {
    $p = Join-Path $Root ($s.Replace("/", "\"))
    if (-not (Test-Path -LiteralPath $p)) { Write-Host "WARN scope not found: $s" -ForegroundColor Yellow; continue }
    $mode = "claude"; if ($Conf.whole -contains $s) { $mode = "whole" }
    Walk $p $mode $false
}
foreach ($r in $RootFiles) {
    $p = Join-Path $Root ($r.Replace("/", "\"))
    if (Test-Path -LiteralPath $p -PathType Leaf) { $found.Add(@{ Src = (ConvertTo-LongPath $p); Rel = $r }) }
    else { Write-Host "WARN root file not found: $r" -ForegroundColor Yellow }
}

$desc = ($AllScopes + $RootFiles) -join ", "
if ($Conf.exclude.Count -gt 0) { $desc += " (excluding " + ($Conf.exclude -join ", ") + ")" }
Write-Host ("Found {0} file(s) in: {1}" -f $found.Count, $desc)
foreach ($b in $skippedBig) { Write-Host "  skipped (> 1 MB): $b" -ForegroundColor Yellow }

if ($DryRun) {
    $found | ForEach-Object { Write-Host ("  " + $_.Rel) }
    Write-Host "Dry run: nothing changed."
    exit 0
}

function Invoke-Git { $ErrorActionPreference = "Continue"; & git @args; if ($LASTEXITCODE -ne 0) { throw ("git " + ($args -join " ") + " failed") } }
function Test-Git { $ErrorActionPreference = "Continue"; & git @args *> $null; return ($LASTEXITCODE -eq 0) }

# Git prep: every run goes to a NEW branch sync/<timestamp>, so main only changes
# when you merge the pull request on GitHub. The branch starts from origin/main, or
# from the newest sync branch while that one is still open (stacked), so edits that
# exist only on an unmerged branch (e.g. script changes) are never dropped.
Push-Location $Mirror
try {
    if (-not (Test-Path -LiteralPath (Join-Path $Mirror ".git"))) { Invoke-Git init -q -b main }
    Invoke-Git config core.longpaths true
    # Use Windows OpenSSH so git can use keys unlocked in the Windows ssh-agent
    # (Git for Windows' bundled ssh cannot reach that agent).
    $winSsh = "C:/Windows/System32/OpenSSH/ssh.exe"
    if ((Test-Path -LiteralPath $winSsh) -and -not (Test-Git config --local --get core.sshCommand)) {
        Invoke-Git config core.sshCommand $winSsh
    }
    if ($Remote) {
        if (Test-Git remote get-url origin) { Invoke-Git remote set-url origin $Remote } else { Invoke-Git remote add origin $Remote }
    }
    $hasRemote = Test-Git remote get-url origin
    if ($hasRemote -and -not $NoPush) { Invoke-Git fetch -q --prune origin }
    if (Test-Git rev-parse -q --verify refs/remotes/origin/main) { $base = "origin/main" }
    elseif (Test-Git rev-parse -q --verify refs/heads/main) { $base = "main" }
    else { throw "No main branch yet. Commit the scripts to main and push it once (see README)." }
    # The scope folders are regenerated below, so discard leftovers from a failed run.
    foreach ($s in ($AllScopes + $RootFiles)) {
        if (Test-Git ls-files --error-unmatch -- $s) { Invoke-Git checkout -q HEAD -- $s }
        Invoke-Git clean -fdq -- $s
    }
    # Newest existing sync branch: a run with nothing new since then makes no new branch.
    $Prev = git for-each-ref --sort=-refname --count=1 --format="%(refname:short)" refs/heads/sync/
    $Branch = "sync/" + (Get-Date -Format "yyyyMMdd-HHmmss")
    $start = $base
    if ($Prev) {
        $merged = Test-Git merge-base --is-ancestor $Prev $base
        $onRemote = Test-Git rev-parse -q --verify ("refs/remotes/origin/" + $Prev)
        $wasPushed = Test-Git config --get ("branch." + $Prev + ".remote")
        # Open = not merged, and still on GitHub (or never pushed). A branch closed and
        # deleted on GitHub without merging counts as rejected and is not stacked on.
        if (-not $merged -and ($onRemote -or -not $wasPushed)) {
            $start = $Prev
            Write-Host "Stacking on unmerged $Prev. Merge only the newest PR and close the older one."
        }
    }
    Invoke-Git checkout -q -b $Branch $start
} finally { Pop-Location }

# Rebuild the mirror from scratch: remove every tracked file except the mirror's own
# (so files that left the scope or were deleted at the source drop out), then copy.
Push-Location $Mirror
try { $tracked = @(git -c core.quotepath=off ls-files) } finally { Pop-Location }
foreach ($t in $tracked) {
    if ($MirrorOwn -contains $t) { continue }
    $tp = ConvertTo-LongPath (Join-Path $Mirror ($t.Replace("/", "\")))
    if ([IO.File]::Exists($tp)) { [IO.File]::Delete($tp) }
}
foreach ($s in ($AllScopes + $RootFiles)) {
    $target = [IO.Path]::GetFullPath((Join-Path $Mirror ($s.Replace("/", "\"))))
    if (-not $target.StartsWith($Mirror + "\")) { throw "Refusing to delete outside mirror: $target" }
    if ([IO.Directory]::Exists((ConvertTo-LongPath $target))) { [IO.Directory]::Delete((ConvertTo-LongPath $target), $true) }
    elseif ([IO.File]::Exists((ConvertTo-LongPath $target))) { [IO.File]::Delete((ConvertTo-LongPath $target)) }
}
foreach ($f in $found) {
    $dst = Join-Path $Mirror ($f.Rel.Replace("/", "\"))
    [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath (Split-Path -Parent $dst)))
    [IO.File]::Copy($f.Src, (ConvertTo-LongPath $dst), $true)
}

# Secret scan on the mirrored copies.
$hits = @()
foreach ($f in $found) {
    $text = [IO.File]::ReadAllText((ConvertTo-LongPath (Join-Path $Mirror ($f.Rel.Replace("/", "\")))))
    foreach ($pat in $SecretPatterns) {
        if ($text -match $pat) { $hits += ("{0}  ~ {1}" -f $f.Rel, $pat); break }
    }
}
if ($hits.Count -gt 0) {
    Write-Host "Possible secrets found:" -ForegroundColor Red
    $hits | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    if (-not $AllowSecrets) {
        Write-Host "Not committing. Remove them at the source, or rerun with -AllowSecrets if they are false positives." -ForegroundColor Red
        exit 2
    }
}

# Commit on the sync branch and push it.
Push-Location $Mirror
try {
    Invoke-Git add -A
    foreach ($f in $found) {
        if ($ExecExt -contains [IO.Path]::GetExtension($f.Rel).ToLower()) { Invoke-Git update-index --chmod=+x -- $f.Rel }
    }
    if (Test-Git diff --cached --quiet $base) {
        Write-Host "Nothing new compared with $base. Removing branch $Branch."
        Invoke-Git checkout -q --detach $base
        Invoke-Git branch -q -D $Branch
        exit 0
    }
    if ($Prev -and (Test-Git diff --cached --quiet $Prev)) {
        Write-Host "Nothing new since $Prev (not merged yet). Review that branch instead; no new branch made."
        Invoke-Git checkout -q $Prev
        Invoke-Git branch -q -D $Branch
        exit 0
    }
    if (-not $Message) { $Message = "Sync Claude files " + (Get-Date -Format "yyyy-MM-dd HH:mm") }
    Invoke-Git commit -q -m $Message
    git --no-pager log -1 --shortstat --format="Committed %h on $Branch"
    # Each sync branch holds everything not yet on main, so older local ones are redundant.
    git for-each-ref --format="%(refname:short)" refs/heads/sync/ | Where-Object { $_ -ne $Branch } | ForEach-Object { Invoke-Git branch -q -D $_ }
    if ($NoPush) { Write-Host "Push skipped (-NoPush). Branch $Branch is local only."; exit 0 }
    if (-not $hasRemote) {
        Write-Host "No remote yet. Run once with: -Remote git@github.com:<owner>/<repo>.git" -ForegroundColor Yellow
        exit 0
    }
    Invoke-Git push -q -u origin $Branch
    $url = (git remote get-url origin) -replace '^git@github\.com:', 'https://github.com/' -replace '\.git$', ''
    Write-Host ""
    Write-Host "Pushed $Branch. Review and merge it here:" -ForegroundColor Green
    Write-Host ("  {0}/compare/main...{1}?expand=1" -f $url, $Branch) -ForegroundColor Green
} finally {
    Pop-Location
}
