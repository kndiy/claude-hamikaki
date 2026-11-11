<#
.SYNOPSIS
    Check tools, pick a project under the projects root, and open Claude Code in it.
.DESCRIPTION
    Tree: root > subRoot > [group > ...] > project. A project is a folder that
    contains a .claude-project marker; every other folder under a subRoot is a group.

    Tools come from the Windows tools table in any CLAUDE.md on the path. The root
    table is checked at start; the subRoot, group and project tables are checked
    once a project is chosen. Detected versions are written back to each table.
    Missing required/recommended tools go to Install-Tools.ps1.
.PARAMETER Root
    Use this root for this run only. The saved root lives in config.json (default
    "..", the folder holding _launcher); change it with the C option in the menu.
.PARAMETER SkipToolCheck
    Skip all tool checks (faster start).
#>
param(
    [string]$Root,
    [switch]$SkipToolCheck
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Launcher.Common.ps1')

$config = Get-LauncherConfig $PSScriptRoot
$Root = Resolve-LauncherRoot $PSScriptRoot $(if ($Root) { $Root } else { $config.root })

# ---------- tool check ----------

function Invoke-ToolCheck([string[]]$ClaudeMds) {
    if ($SkipToolCheck) { return }
    $missing   = @(Test-Tools $ClaudeMds)
    $toInstall = @($missing | Where-Object { $_.Level -in 'required', 'recommended' })
    if ($toInstall) {
        Write-Host "`nMissing tools will be installed with:" -ForegroundColor Yellow
        foreach ($t in $toInstall) { Write-Host "  $($t.Name) (in $($t.Dir)): $($t.Install)" }
        if ((Read-Line 'Install now? [Y/n]') -notmatch '^[Nn]') {
            foreach ($g in ($toInstall | Group-Object Dir)) {
                & (Join-Path $PSScriptRoot 'Install-Tools.ps1') -ClaudeMd (Join-Path $g.Name 'CLAUDE.md') -Names ($g.Group | ForEach-Object Name)
            }
            $missing = @(Test-Tools $ClaudeMds)
        }
    }
    $blocking = @($missing | Where-Object { $_.Level -eq 'required' })
    if ($blocking) {
        Write-Host "`nRequired tools are missing. Install them, open a NEW terminal, and run the launcher again:" -ForegroundColor Red
        foreach ($t in $blocking) { Write-Host "  $($t.Name) (in $($t.Dir)): $($t.Install)" }
        exit 1
    }
}

# ---------- tree navigation ----------

function Get-Children([string]$Dir) {
    @(Get-ChildItem -LiteralPath $Dir -Directory -Force |
        Where-Object { $_.Name -notmatch '^[._]' -and -not ($_.Attributes -band [IO.FileAttributes]::Hidden) } |
        Sort-Object Name)
}

function Test-Project([string]$Dir) { Test-Path -LiteralPath (Join-Path $Dir $ProjectMarker) }

function Read-FolderName([string]$Kind) {
    $name = (Read-Line "New $Kind name (blank to cancel)").Trim()
    if (-not $name) { return $null }
    # Same rules on Windows and Linux so the tree can move between them.
    if ($name -match '[\\/:*?"<>|]' -or $name -match '^[._]' -or $name -match '[. ]$') {
        Write-Host 'Invalid name: no \ / : * ? " < > |, must not start with . or _, or end with . or space.' -ForegroundColor Red
        return $null
    }
    $name
}

# Runs Claude Code in a tree folder, then comes back to the menu. Started as a plain
# child process on this console: '& claude' inside a function whose output is captured
# (Select-Project's is) would redirect Claude's stdout into that output and break its UI.
# Arguments must not contain double quotes; trailing backslashes are doubled so the
# closing quote survives Windows command-line parsing.
function Invoke-ClaudeIn([string]$Dir, [string]$Label, [string[]]$ClaudeArgs) {
    $title = $Host.UI.RawUI.WindowTitle
    try {
        $Host.UI.RawUI.WindowTitle = "Claude $Label - $(Split-Path $Dir -Leaf)"
        Write-Host "`nOpening Claude Code in $Dir ($Label)" -ForegroundColor Green
        $psi = New-Object Diagnostics.ProcessStartInfo (Get-Command claude -CommandType Application | Select-Object -First 1).Source
        $psi.Arguments        = ($ClaudeArgs | ForEach-Object { '"' + ($_ -replace '(\\+)$', '$1$1') + '"' }) -join ' '
        $psi.WorkingDirectory = $Dir
        $psi.UseShellExecute  = $false
        [Diagnostics.Process]::Start($psi).WaitForExit()
    } finally {
        $Host.UI.RawUI.WindowTitle = $title
    }
    # The session may have changed the subRoots list; script edits apply on the next start.
    $script:config.subRoots = (Get-LauncherConfig $PSScriptRoot).subRoots
    Write-Host 'Back in the launcher. Changes to launcher scripts apply the next time it starts.' -ForegroundColor DarkGray
}

# Opens Claude Code in a root, subRoot or group folder to edit that level's CLAUDE.md
# and/or the launcher. The launcher folder is added with --add-dir; its CLAUDE.md
# does not load from there, hence the appended prompt.
function Start-SettingsSession([string]$Dir, [string]$Kind) {
    $prompt = "Settings session started from the Claude Projects launcher at the $Kind folder $Dir. " +
        "Scope: this folder's CLAUDE.md (rules and tools tables for everything below it) and, when the user asks, " +
        "the launcher in $PSScriptRoot. Read $(Join-Path $PSScriptRoot 'CLAUDE.md') before changing the launcher, " +
        "keep the Windows (.ps1, ASCII only), Linux (linux/*.sh) and macOS (mac/*.sh) versions in sync, and update that CLAUDE.md. " +
        "Do not change projects below this folder unless asked. When the user exits, the launcher menu comes back."
    Invoke-ClaudeIn $Dir 'settings' @('--add-dir', $PSScriptRoot, '--append-system-prompt', $prompt)
}

# Returns the folder holding projects.json when $Dir is a claude.ai data export
# (at its top or in its only subfolder), otherwise $null.
function Find-ClaudeExport([string]$Dir) {
    if (Test-Path -LiteralPath (Join-Path $Dir 'projects.json')) { return $Dir }
    $subs = @(Get-ChildItem -LiteralPath $Dir -Directory -Force)
    if ($subs.Count -eq 1 -and (Test-Path -LiteralPath (Join-Path $subs[0].FullName 'projects.json'))) { return $subs[0].FullName }
    $null
}

# Asks for a folder (or a .zip, e.g. a claude.ai data export) elsewhere and opens
# Claude Code here to analyse it, ask the user, and copy it in as new project(s).
# A .zip is extracted to a temp folder outside the tree and deleted afterwards.
# Returns the new project's path when exactly one appeared and the user wants to
# open it, otherwise $null.
function Start-MigrateSession([string]$Dir, [string]$Kind) {
    $src = (Read-Line 'Project folder, or claude.ai export .zip, to migrate (blank to cancel)').Trim().Trim('"').TrimEnd('\')
    if (-not $src) { return $null }
    $temp = $null
    if ((Test-Path -LiteralPath $src -PathType Leaf) -and $src -match '(?i)\.zip$') {
        $zip  = (Get-Item -LiteralPath $src).FullName
        $temp = Join-Path $env:TEMP ('claude-migrate-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
        Write-Host "Extracting to $temp (deleted after the session) ..." -ForegroundColor DarkGray
        try { Expand-Archive -LiteralPath $zip -DestinationPath $temp }
        catch { Write-Host "Could not extract: $($_.Exception.Message)" -ForegroundColor Red; return $null }
        $src = $temp
    } elseif (-not (Test-Path -LiteralPath $src -PathType Container)) {
        Write-Host "Folder or .zip not found: $src" -ForegroundColor Red
        return $null
    } else {
        $zip = $null
        $src = (Get-Item -LiteralPath $src).FullName.TrimEnd('\')
        if (([IO.Path]::GetFullPath($Dir).TrimEnd('\') + '\').StartsWith($src + '\', [StringComparison]::OrdinalIgnoreCase)) {
            Write-Host 'That folder contains this one; pick a folder outside it.' -ForegroundColor Red
            return $null
        }
    }
    try {
        $before = @(Get-Children $Dir | ForEach-Object Name)
        $export = Find-ClaudeExport $src
        $origin = if ($zip) { "$src (extracted from $zip; the launcher deletes this copy after the session)" } else { $src }
        $tools  = "Tools section whose Windows table sits between the lines <!-- launcher:tools:windows:start --> and " +
            "<!-- launcher:tools:windows:end --> (columns Tool, Check, Level, Install, Status; Status -)"
        if ($export) { $rules = Get-ExportMigrateRules $origin $export $Kind $Dir $tools }
        else         { $rules = Get-FolderMigrateRules $origin $Kind $Dir $tools }
        $first = if ($export) { "List the projects in the claude.ai export at $export and ask which to migrate here." }
                 else         { "Analyse $src and plan its migration into a new project here." }
        Invoke-ClaudeIn $Dir 'migrate' @('--add-dir', $src, '--append-system-prompt', $rules, $first)
    } finally {
        if ($temp) { Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue }
    }
    $new = @(Get-Children $Dir | Where-Object { $_.Name -notin $before -and (Test-Project $_.FullName) })
    if ($new.Count -eq 1 -and (Read-Line "Open the new project '$($new[0].Name)' now? [Y/n]") -notmatch '^[Nn]') {
        return $new[0].FullName
    }
    $null
}

# Rules for migrating a claude.ai (Claude Desktop) chat-project export. The format is
# not documented by Anthropic, so Claude is told to inspect it rather than trust this.
function Get-ExportMigrateRules([string]$Origin, [string]$Export, [string]$Kind, [string]$Dir, [string]$Tools) {
    "Migration session started from the Claude Projects launcher. Source: $Origin, a claude.ai data export " +
    "(Claude Desktop chat Projects), found at $Export. Target: new project folder(s) inside the $Kind folder $Dir " +
    "(the current folder). Expected content, unverified: projects.json (each project with its custom instructions " +
    "and knowledge docs, the docs as extracted text only), conversations.json (every chat of the account, not linked " +
    "to projects), memories.json, users.json. Inspect the real structure first with a small script (files can be " +
    "large) and do not assume field names. Work in this order. " +
    "1) Without changing anything, list the projects (name, dates, number and size of knowledge docs, whether it has " +
    "instructions), then ask the user (AskUserQuestion) which to migrate, the folder name for each (propose names " +
    "that meet the root naming rules), and whether to bring chats: chats can only be matched by name or topic, so " +
    "show the candidates and let the user confirm them. Do not guess. " +
    "2) For each chosen project: create the folder with an empty .claude-project marker file; write each knowledge " +
    "doc as a UTF-8 file under knowledge/ with its original name; keep the custom instructions verbatim in " +
    "knowledge/_project_instructions.md; write confirmed chats as Markdown under chats/. " +
    "3) Write the project CLAUDE.md per root rule 3: purpose and rules taken from the custom instructions (say which " +
    "parts need the user's review), the knowledge and chats layout, a $Tools, and a change log entry naming the " +
    "export and date. " +
    "4) Verify doc counts against projects.json and report what could not be migrated: original PDFs, images and " +
    "other binaries (only their text is in the export; list them so the user can download them from claude.ai), " +
    "memory, and anything else missing. " +
    "5) Copy only the chosen projects and chats. The export holds the whole account, so never copy it wholesale, " +
    "never copy users.json, and never modify the export. " +
    "When the user exits, the launcher menu comes back and offers to open the new project."
}

function Get-FolderMigrateRules([string]$Origin, [string]$Kind, [string]$Dir, [string]$Tools) {
    "Migration session started from the Claude Projects launcher. Source: $Origin. " +
        "Target: a new project folder inside the $Kind folder $Dir (the current folder). Work in this order. " +
        "1) Analyse the source without changing anything: purpose, tech stack, size and file count, git state " +
        "(uncommitted changes, remotes), existing CLAUDE.md, virtual environments, dependency and build folders, caches, " +
        "secrets, and absolute paths or config tied to the old location (OneDrive, user folders, drive letters, the NAS). " +
        "2) Give a short summary, then ask the user (AskUserQuestion) everything you need before copying: the new folder " +
        "name (propose one that meets the root naming rules), copy or move, what to exclude and rebuild instead, which " +
        "hard-coded paths to change, which tools the project needs. Do not guess. " +
        "3) Migrate: create the folder with an empty .claude-project marker file, copy with robocopy (keep .git), rebuild " +
        "excluded environments if the user agreed, and change only the paths the user approved. " +
        "4) Create or update the project CLAUDE.md per root rule 3, including a $Tools, and a change log entry " +
        "naming the source path and date. " +
        "5) Verify: compare file counts and sizes, run or build the project if that is cheap, and report differences. " +
        "6) Never modify the source. Delete it only if the user chose move, after verification, and after asking again. " +
        "Never change or delete anything on L: (the company NAS). " +
        "When the user exits, the launcher menu comes back and offers to open the new project."
}

function New-TreeFolder([string]$Parent, [string]$Kind) {
    $name = Read-FolderName $Kind
    if (-not $name) { return $null }
    $dir = Join-Path $Parent $name
    if (Test-Path -LiteralPath $dir) { Write-Host "'$name' already exists." -ForegroundColor Red; return $null }
    New-Item -ItemType Directory -Path $dir | Out-Null
    New-FolderClaudeMd $dir $Kind
    if ($Kind -eq 'project') { New-Item -ItemType File -Path (Join-Path $dir $ProjectMarker) | Out-Null }
    Write-Host "Created $Kind '$name'" -ForegroundColor Green
    $dir
}

# Returns the folders from the subRoot down to the chosen project, or $null to go
# back to the subRoot menu. Navigation keeps a stack of visited folders instead of
# comparing path strings, which break on 8.3 short names and case differences.
function Select-Project([string]$SubRootDir) {
    $stack = New-Object System.Collections.Generic.List[string]
    $stack.Add($SubRootDir)
    while ($true) {
        $dir   = $stack[$stack.Count - 1]
        $atTop = $stack.Count -eq 1
        $items = Get-Children $dir
        $crumb = ($stack | ForEach-Object { Split-Path $_ -Leaf }) -join ' > '
        Write-Host "`n== $crumb ==" -ForegroundColor Cyan
        if (-not $items) { Write-Host '  (empty)' -ForegroundColor DarkGray }
        for ($i = 0; $i -lt $items.Count; $i++) {
            $isProj = Test-Project $items[$i].FullName
            $tag = if ($isProj) { '[project]' } else { '[group]  ' }
            Write-Host ("  {0,2}) {1} {2}" -f ($i + 1), $tag, $items[$i].Name) -ForegroundColor $(if ($isProj) { 'White' } else { 'Gray' })
        }
        Write-Host '   N) new project here   G) new group here' -ForegroundColor DarkCyan
        Write-Host '   M) migrate a project here: folder or claude.ai export .zip (Claude analyses, asks first)' -ForegroundColor DarkCyan
        Write-Host '   S) settings: open Claude Code here (this CLAUDE.md, the launcher)' -ForegroundColor DarkCyan
        if (-not $atTop) { Write-Host '   O) open THIS folder as a project (marks it)' -ForegroundColor DarkCyan }
        Write-Host '   B) back   Q) quit' -ForegroundColor DarkCyan

        $c = (Read-Line 'Choose')
        if ($c -match '^\d+$' -and [int]$c -ge 1 -and [int]$c -le $items.Count) {
            $pick = $items[[int]$c - 1].FullName
            $stack.Add($pick)
            if (Test-Project $pick) { return , $stack.ToArray() }
            continue
        }
        switch -regex ($c) {
            '^[Nn]$' { $p = New-TreeFolder $dir 'project'; if ($p) { $stack.Add($p); return , $stack.ToArray() } }
            '^[Gg]$' { $g = New-TreeFolder $dir 'group'; if ($g) { $stack.Add($g) } }
            '^[Ss]$' { Start-SettingsSession $dir $(if ($atTop) { 'subRoot' } else { 'group' }) }
            '^[Mm]$' {
                $m = Start-MigrateSession $dir $(if ($atTop) { 'subRoot' } else { 'group' })
                if ($m) { $stack.Add($m); return , $stack.ToArray() }
            }
            '^[Oo]$' {
                if ($atTop) { break }
                if (@(Get-Children $dir | Where-Object { Test-Project $_.FullName }).Count) {
                    Write-Host 'This folder contains projects, so it must stay a group.' -ForegroundColor Red
                    break
                }
                New-Item -ItemType File -Force -Path (Join-Path $dir $ProjectMarker) | Out-Null
                New-FolderClaudeMd $dir 'project'
                return , $stack.ToArray()
            }
            '^[Bb]$' {
                if ($atTop) { return $null }
                $stack.RemoveAt($stack.Count - 1)
            }
            '^[Qq]$' { exit 0 }
            default { Write-Host 'Invalid choice.' -ForegroundColor Red }
        }
    }
}

function Set-NewRoot {
    $path = (Read-Line 'New root path (blank to cancel)').Trim().Trim('"').TrimEnd('\')
    if (-not $path) { return }
    $err = Test-RootPath $path
    if ($err) { Write-Host $err -ForegroundColor Red; return }
    if (-not (Test-Path -LiteralPath $path)) {
        if ((Read-Line "'$path' does not exist. Create it? [y/N]") -notmatch '^[Yy]') { return }
    }
    # A new root starts from the current root's rules and tools table.
    Initialize-Root $path $config.subRoots (Join-Path $Root 'CLAUDE.md')
    $config.root = $path
    Save-LauncherConfig $PSScriptRoot $config
    $script:Root = $path
    Write-Host "Root is now $path (saved to config.json)." -ForegroundColor Green
}

# ---------- main ----------

$err = Test-RootPath $Root
if ($err) { Write-Host "Root '$Root': $err" -ForegroundColor Red; exit 1 }
Initialize-Root $Root $config.subRoots $null
Set-Location -LiteralPath $Root
Write-Host "Root: $Root" -ForegroundColor Cyan
Invoke-ToolCheck @(Join-Path $Root 'CLAUDE.md')

$chain = $null
while (-not $chain) {
    Write-Host "`n== $Root ==" -ForegroundColor Cyan
    for ($i = 0; $i -lt $config.subRoots.Count; $i++) { Write-Host ("  {0,2}) {1}" -f ($i + 1), $config.subRoots[$i]) }
    Write-Host '   C) change root   S) settings: open Claude Code at the root   Q) quit' -ForegroundColor DarkCyan
    $c = (Read-Line 'Choose subRoot')
    if ($c -match '^\d+$' -and [int]$c -ge 1 -and [int]$c -le $config.subRoots.Count) {
        $chain = Select-Project (Join-Path $Root $config.subRoots[[int]$c - 1])
    } elseif ($c -match '^[Ss]$') {
        Start-SettingsSession $Root 'root'
    } elseif ($c -match '^[Cc]$') {
        Set-NewRoot
        Invoke-ToolCheck @(Join-Path $Root 'CLAUDE.md')
        Set-Location -LiteralPath $Root
    } elseif ($c -match '^[Qq]$') {
        exit 0
    }
}

# Tools declared by the subRoot, groups and the project itself.
Invoke-ToolCheck @($chain | ForEach-Object { Join-Path $_ 'CLAUDE.md' })

$project = $chain[-1]
Set-Location -LiteralPath $project
$Host.UI.RawUI.WindowTitle = "Claude - $(Split-Path $project -Leaf)"
Write-Host "`nOpening Claude Code in $project" -ForegroundColor Green
& claude
