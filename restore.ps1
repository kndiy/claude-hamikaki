# restore.ps1 - fetch the reviewed mirror (origin/main) and copy its files back
# into the real projects tree. Windows PowerShell 5.1, ASCII only.
# Linux/macOS twin: restore.sh (keep both in sync).
#
# Per file:
#   missing locally                         -> created
#   identical                               -> left alone
#   differs, local = an older mirrored copy -> updated (local was just stale)
#   differs, local never mirrored           -> CONFLICT: skipped (unsynced local edits);
#                                              -Force backs it up, then overwrites
# Never deletes local files. Skips files whose folder does not exist yet (clone the
# project first, else git clone refuses the non-empty folder); -CreateMissing overrides.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File restore.ps1 -DryRun   # show what would happen
#   powershell -ExecutionPolicy Bypass -File restore.ps1           # apply
#   ... -Force           overwrite conflicts (backups in .restore-backup\<timestamp>\)
#   ... -CreateMissing   also create files in folders that do not exist yet
#   ... -Ref <ref>       restore from another ref (default origin/main)
param(
    [switch]$DryRun,
    [switch]$Force,
    [switch]$CreateMissing,
    [string]$Ref = "origin/main"
)
$ErrorActionPreference = "Stop"

# Files that belong to the mirror repo itself, never restored.
$MirrorOwn = @("README.md", "mirror.conf", "sync.ps1", "sync.sh", "restore.ps1", "restore.sh", ".gitignore", ".gitattributes")

$Mirror = $PSScriptRoot
$Root = Split-Path -Parent $Mirror
$Stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$BackupDir = Join-Path $Mirror (".restore-backup\" + $Stamp)
$Utf8 = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = $Utf8
[Console]::OutputEncoding = $Utf8

# Scope from mirror.conf (same format as sync.ps1), used to keep the
# "removed in mirror" report to this mirror's own paths.
$Conf = @{ claude = @(); whole = @(); files = @(); exclude = @(); scripts = @() }
$section = ""
foreach ($raw in [IO.File]::ReadAllLines((Join-Path $Mirror "mirror.conf"))) {
    $line = ($raw -replace '#.*$', '').Trim()
    if (-not $line) { continue }
    if ($line -match '^\[(\w+)\]$') { $section = $Matches[1].ToLower(); continue }
    if ($section -and $Conf.ContainsKey($section) -and $section -ne "scripts") { $Conf[$section] += $line.Replace("\", "/").TrimEnd("/") }
}
function Test-InScope([string]$rel) {
    if ($Conf.files -contains $rel) { return $true }
    foreach ($x in $Conf.exclude) { if ($rel.StartsWith($x + "/")) { return $false } }
    foreach ($s in (@($Conf.claude) + @($Conf.whole))) { if ($rel.StartsWith($s + "/")) { return $true } }
    return $false
}

function ConvertTo-LongPath([string]$p) { if ($p.StartsWith("\\?\")) { $p } else { "\\?\" + $p } }
function Invoke-Git { $ErrorActionPreference = "Continue"; $o = & git -c core.quotepath=off -c core.longpaths=true @args; if ($LASTEXITCODE -ne 0) { throw ("git " + ($args -join " ") + " failed") }; $o }
function Test-Git { $ErrorActionPreference = "Continue"; & git @args *> $null; return ($LASTEXITCODE -eq 0) }

Push-Location $Mirror
$tmp = $null
try {
    $winSsh = "C:/Windows/System32/OpenSSH/ssh.exe"
    if ((Test-Path -LiteralPath $winSsh) -and -not (Test-Git config --local --get core.sshCommand)) {
        Invoke-Git config core.sshCommand $winSsh | Out-Null
    }
    if ($Ref -like "origin/*") { Invoke-Git fetch -q origin | Out-Null }
    if (-not (Test-Git rev-parse -q --verify ($Ref + "^{commit}"))) { throw "Ref not found: $Ref" }
    Write-Host ("Restoring from {0} ({1})" -f $Ref, ((Invoke-Git log -1 --format="%h %s" $Ref) -join ""))

    # Every version of every path in the reviewed history of $Ref: "path<TAB>blob".
    # Only $Ref, not unmerged sync branches: a local file matching a pending
    # (unmerged) version must count as a conflict, not be reverted.
    $known = New-Object 'System.Collections.Generic.HashSet[string]'
    $everPaths = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($line in (Invoke-Git log --raw --no-abbrev --no-renames --format= $Ref)) {
        if (-not $line.StartsWith(":")) { continue }
        $parts = $line.Split("`t")
        $meta = $parts[0].Split(" ")
        $path = $parts[1]
        [void]$everPaths.Add($path)
        foreach ($b in @($meta[2], $meta[3])) { if ($b -notmatch '^0+$') { [void]$known.Add($path + "`t" + $b) } }
    }

    # Files at the ref.
    $entries = @()
    foreach ($line in (Invoke-Git ls-tree -r $Ref)) {
        $parts = $line.Split("`t")
        $meta = $parts[0].Split(" ")
        if ($MirrorOwn -contains $parts[1]) { continue }
        $entries += @{ Mode = $meta[0]; Blob = $meta[2]; Rel = $parts[1] }
    }

    # Check out the ref into a temporary worktree to read the file contents.
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("claude-restore-" + $Stamp)
    Invoke-Git worktree add -q --detach $tmp $Ref | Out-Null

    $counts = @{ created = 0; updated = 0; same = 0; conflict = 0; forced = 0; nofolder = 0 }
    foreach ($e in $entries) {
        $dst = Join-Path $Root ($e.Rel.Replace("/", "\"))
        $src = Join-Path $tmp ($e.Rel.Replace("/", "\"))
        $dstL = ConvertTo-LongPath $dst
        $parent = Split-Path -Parent $dst
        $action = ""
        if (-not [IO.File]::Exists($dstL)) {
            if (-not [IO.Directory]::Exists((ConvertTo-LongPath $parent)) -and -not $CreateMissing) { $action = "nofolder" }
            else { $action = "created" }
        } else {
            $h = (Invoke-Git hash-object --no-filters -- $dst) -join ""
            if ($h -eq $e.Blob) { $action = "same" }
            elseif ($known.Contains($e.Rel + "`t" + $h)) { $action = "updated" }
            elseif ($Force) { $action = "forced" }
            else { $action = "conflict" }
        }
        $counts[$action]++
        # Plain if/continue: "continue" inside a switch only leaves the switch, not the loop.
        if ($action -eq "same") { continue }
        if ($action -eq "nofolder") { Write-Host ("  skip (folder missing)  {0}" -f $e.Rel) -ForegroundColor DarkGray; continue }
        if ($action -eq "conflict") { Write-Host ("  CONFLICT (local edits) {0}" -f $e.Rel) -ForegroundColor Red; continue }
        Write-Host ("  {0,-22} {1}" -f $action, $e.Rel)
        if ($DryRun) { continue }
        if ($action -eq "forced") {
            $bak = Join-Path $BackupDir ($e.Rel.Replace("/", "\"))
            [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath (Split-Path -Parent $bak)))
            [IO.File]::Copy($dstL, (ConvertTo-LongPath $bak), $true)
        }
        [void][IO.Directory]::CreateDirectory((ConvertTo-LongPath $parent))
        [IO.File]::Copy((ConvertTo-LongPath $src), $dstL, $true)
    }

    # Paths the mirror once held but no longer does: report only.
    $atRef = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($e in $entries) { [void]$atRef.Add($e.Rel) }
    foreach ($p in $everPaths) {
        if ($atRef.Contains($p) -or ($MirrorOwn -contains $p) -or -not (Test-InScope $p)) { continue }
        if ([IO.File]::Exists((ConvertTo-LongPath (Join-Path $Root ($p.Replace("/", "\")))))) {
            Write-Host ("  removed in mirror, kept locally: {0}" -f $p) -ForegroundColor Yellow
        }
    }

    Write-Host ("Created {0}, updated {1}, unchanged {2}, conflicts {3}, forced {4}, skipped (no folder) {5}." -f `
        $counts.created, $counts.updated, $counts.same, $counts.conflict, $counts.forced, $counts.nofolder)
    if ($DryRun) { Write-Host "Dry run: nothing changed." }
    elseif ($counts.forced -gt 0) { Write-Host "Backups of overwritten files: $BackupDir" }
    if ($counts.conflict -gt 0) {
        Write-Host "Conflicts were skipped. Run sync first to send those local edits for review, or rerun with -Force." -ForegroundColor Red
        exit 3
    }
} finally {
    if ($tmp -and (Test-Path -LiteralPath $tmp)) { & git worktree remove --force $tmp *> $null }
    & git worktree prune *> $null
    Pop-Location
}
