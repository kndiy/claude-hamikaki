# Shared helpers for Start-Claude.ps1, Install-Tools.ps1 and Launch.ps1. Dot-source only.
# Keep this file ASCII: Windows PowerShell 5.1 reads BOM-less scripts as ANSI.

$script:Utf8NoBom     = New-Object System.Text.UTF8Encoding($false)
$script:ToolsStart    = '<!-- launcher:tools:windows:start -->'
$script:ToolsEnd      = '<!-- launcher:tools:windows:end -->'
$script:ProjectMarker = '.claude-project'

# ---------- files ----------

function Read-Text([string]$Path) { [IO.File]::ReadAllText($Path) }

function Write-Text([string]$Path, [string]$Text) { [IO.File]::WriteAllText($Path, $Text, $script:Utf8NoBom) }

# Read-Host that treats end of input (piped stdin exhausted) as quit.
function Read-Line([string]$Prompt) {
    $r = Read-Host $Prompt
    if ($null -eq $r) { exit 0 }
    $r.Trim()
}

# ---------- config ----------

# root may be relative to the launcher folder ("..", the default) so the whole
# tree can be copied to another machine. A saved absolute root that does not
# exist on this machine falls back to "..".
function Get-LauncherConfig([string]$LauncherDir) {
    $path = Join-Path $LauncherDir 'config.json'
    $cfg  = if (Test-Path -LiteralPath $path) { Read-Text $path | ConvertFrom-Json } else { $null }
    $get  = { param($name, $default) if ($cfg -and $cfg.$name) { $cfg.$name } else { $default } }
    [pscustomobject]@{
        root      = [string](& $get 'root' '..')
        subRoots  = @(& $get 'subRoots' @('codes', 'noncodes'))
        wslDistro = [string](& $get 'wslDistro' 'Ubuntu')
        wslRoot   = [string](& $get 'wslRoot' '~/Projects')
    }
}

function Resolve-LauncherRoot([string]$LauncherDir, [string]$Root) {
    $default = [IO.Path]::GetFullPath((Join-Path $LauncherDir '..'))
    if (-not $Root) { return $default }
    $full = if ([IO.Path]::IsPathRooted($Root)) { $Root } else { [IO.Path]::GetFullPath((Join-Path $LauncherDir $Root)) }
    if (-not (Test-Path -LiteralPath $full)) {
        Write-Host "Saved root '$full' not found on this machine; using $default" -ForegroundColor Yellow
        return $default
    }
    $full.TrimEnd('\')
}

function Save-LauncherConfig([string]$LauncherDir, $Config) {
    Write-Text (Join-Path $LauncherDir 'config.json') ($Config | ConvertTo-Json)
}

# Returns an error message, or $null when the path is an acceptable root.
function Test-RootPath([string]$Path) {
    if (-not [IO.Path]::IsPathRooted($Path)) { return 'Use an absolute path.' }
    if ($Path -match '(?i)\\OneDrive') { return 'Roots must not be on OneDrive.' }
    if ($Path -match '(?i)^L:') { return 'L:\ is the company NAS (read/copy only).' }
    $null
}

# ---------- CLAUDE.md templates ----------

function Get-DefaultToolBlock {
@"
$($script:ToolsStart)
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
| Claude Code | ``claude --version`` | required | ``irm https://claude.ai/install.ps1 \| iex`` | - |
| Git for Windows | ``git --version`` | recommended | ``winget install --id Git.Git -e --source winget --accept-source-agreements --accept-package-agreements`` | - |
$($script:ToolsEnd)
"@
}

function New-FolderClaudeMd([string]$Dir, [string]$Kind) {
    $path = Join-Path $Dir 'CLAUDE.md'
    if (Test-Path -LiteralPath $path) { return }
    $name  = Split-Path $Dir -Leaf
    $today = Get-Date -Format 'yyyy-MM-dd'
    if ($Kind -eq 'project') {
        $body = @"
# $name

Created $today by the launcher. Fill in each section during the first session.

## Purpose and goals
TBD

## Tech stack and key dependencies
TBD

## Structure and key files
TBD

## Commands (build / run / test / lint)
TBD

## Conventions, constraints, decisions
TBD

## Open questions and next steps
TBD

## Tools
Add rows for tools this project needs; the launcher checks them when the project is opened and runs Install from this folder.

$($script:ToolsStart)
| Tool | Check | Level | Install | Status |
|---|---|---|---|---|
$($script:ToolsEnd)

## Change log
- ${today}: Project created.
"@
    } else {
        $body = @"
# $name ($Kind)

Rules here apply to every project below this folder. Put project-specific rules in that project's own CLAUDE.md.

## Scope
TBD

## Shared rules
(none yet)
"@
    }
    Write-Text $path ($body + "`r`n")
}

# Creates the root, its subRoots and their CLAUDE.md files when missing.
# $SeedClaudeMd seeds the root CLAUDE.md for a brand-new root.
function Initialize-Root([string]$Root, [string[]]$SubRoots, [string]$SeedClaudeMd) {
    New-Item -ItemType Directory -Force -Path $Root | Out-Null
    $rootMd = Join-Path $Root 'CLAUDE.md'
    if (-not (Test-Path -LiteralPath $rootMd)) {
        if ($SeedClaudeMd -and (Test-Path -LiteralPath $SeedClaudeMd)) {
            Copy-Item -LiteralPath $SeedClaudeMd -Destination $rootMd
        } else {
            Write-Text $rootMd "# Root rules`r`n`r`nApplies to every project under this root.`r`n"
        }
    }
    if ((Read-Text $rootMd) -notmatch [regex]::Escape($script:ToolsStart)) {
        Write-Text $rootMd ((Read-Text $rootMd).TrimEnd() + "`r`n`r`n### Windows tools`r`n`r`n" + (Get-DefaultToolBlock) + "`r`n")
    }
    foreach ($s in $SubRoots) {
        $dir = Join-Path $Root $s
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        New-FolderClaudeMd $dir 'subRoot'
    }
}

# ---------- tools tables ----------

function Split-Row([string]$Line) {
    $t = $Line.Trim()
    if ($t.StartsWith('|')) { $t = $t.Substring(1) }
    if ($t.EndsWith('|') -and -not $t.EndsWith('\|')) { $t = $t.Substring(0, $t.Length - 1) }
    , @([regex]::Split($t, '(?<!\\)\|') | ForEach-Object { $_.Trim() })
}

function ConvertFrom-Cell([string]$Cell) {
    $c = $Cell.Trim()
    if ($c.Length -ge 2 -and $c.StartsWith('`') -and $c.EndsWith('`')) { $c = $c.Substring(1, $c.Length - 2) }
    $c -replace '\\\|', '|'
}

# Line indexes of the data rows inside the marker block (header and separator skipped).
function Get-ToolRowIndexes([string[]]$Lines) {
    $start = -1; $end = -1
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i].Trim() -eq $script:ToolsStart) { $start = $i }
        elseif ($Lines[$i].Trim() -eq $script:ToolsEnd -and $start -ge 0) { $end = $i; break }
    }
    if ($start -lt 0 -or $end -le $start + 1) { return @() }
    $rows = @(($start + 1)..($end - 1) | Where-Object { $Lines[$_].Trim().StartsWith('|') })
    @($rows | Select-Object -Skip 2)
}

# Tools from one CLAUDE.md; empty when the file or its table is missing.
function Get-ToolTable([string]$ClaudeMd) {
    if (-not (Test-Path -LiteralPath $ClaudeMd)) { return }
    $lines = (Read-Text $ClaudeMd) -split "`r?`n"
    foreach ($i in Get-ToolRowIndexes $lines) {
        $c = Split-Row $lines[$i]
        if ($c.Count -lt 5) { continue }
        [pscustomobject]@{
            Name    = ConvertFrom-Cell $c[0]
            Check   = ConvertFrom-Cell $c[1]
            Level   = (ConvertFrom-Cell $c[2]).ToLower()
            Install = ConvertFrom-Cell $c[3]
            Status  = ConvertFrom-Cell $c[4]
            Dir     = Split-Path $ClaudeMd -Parent
        }
    }
}

# Writes Status cells back; only touches the file when something changed.
function Set-ToolStatus([string]$ClaudeMd, [hashtable]$StatusByName) {
    $text  = Read-Text $ClaudeMd
    $nl    = if ($text -match "`r`n") { "`r`n" } else { "`n" }
    $lines = $text -split "`r?`n"
    $changed = $false
    foreach ($i in Get-ToolRowIndexes $lines) {
        $c = Split-Row $lines[$i]
        $name = ConvertFrom-Cell $c[0]
        if ($c.Count -lt 5 -or -not $StatusByName.ContainsKey($name)) { continue }
        $c[4] = $StatusByName[$name]
        $new = '| ' + ($c -join ' | ') + ' |'
        if ($new -ne $lines[$i]) { $lines[$i] = $new; $changed = $true }
    }
    if ($changed) { Write-Text $ClaudeMd ($lines -join $nl) }
}

# Runs Check (a PowerShell command line) in $Dir. Returns the first version number
# printed, 'installed' when none is printed, or $null when the check fails: a
# non-zero exit code, a PowerShell error (e.g. Get-Item on a missing file), or an
# unknown command. Native stderr output alone is not a failure.
function Get-ToolVersion([string]$Check, [string]$Dir) {
    $ErrorActionPreference = 'Continue'
    $exe = ($Check.Trim() -split '\s+')[0]
    # The Microsoft Store "python" alias is a stub that opens the Store, not a real install.
    $cmd = Get-Command $exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($exe -like 'python*' -and $cmd -and $cmd.Source -like '*\WindowsApps\*') { return $null }
    if ($Dir) { Push-Location -LiteralPath $Dir }
    try {
        $global:LASTEXITCODE = 0
        $res = @(Invoke-Expression "$Check 2>&1")
        if ($LASTEXITCODE -ne 0) { return $null }
        $psErrors = $res | Where-Object { $_ -is [Management.Automation.ErrorRecord] -and $_.FullyQualifiedErrorId -notlike 'NativeCommandError*' }
        if ($psErrors) { return $null }
        $out = $res | Out-String
    } catch {
        return $null
    } finally {
        $global:LASTEXITCODE = 0
        if ($Dir) { Pop-Location }
    }
    if ($out -match '\d+(\.\d+)+') { return $Matches[0] }
    'installed'
}

# Checks every tool, prints a table, and writes Status back to each CLAUDE.md.
# Returns the tools that are missing.
function Test-Tools([string[]]$ClaudeMds) {
    $missing = @()
    foreach ($md in $ClaudeMds) {
        $tools = @(Get-ToolTable $md)
        if (-not $tools) { continue }
        Write-Host "`nTools in $md" -ForegroundColor Cyan
        $status = @{}
        foreach ($t in $tools) {
            $v = Get-ToolVersion $t.Check $t.Dir
            $status[$t.Name] = if ($v) { $v } else { 'MISSING' }
            if (-not $v) { $missing += $t }
            $color = if ($v) { 'Green' } elseif ($t.Level -eq 'optional') { 'DarkGray' } else { 'Yellow' }
            Write-Host ("  {0,-20} {1,-12} {2}" -f $t.Name, $t.Level, $status[$t.Name]) -ForegroundColor $color
        }
        Set-ToolStatus $md $status
    }
    $missing
}

# Picks up PATH changes made by installers without opening a new terminal.
function Update-SessionPath {
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                [Environment]::GetEnvironmentVariable('Path', 'User')
    $localBin = Join-Path $env:USERPROFILE '.local\bin'  # Claude Code native install location
    if ((Test-Path $localBin) -and ($env:Path -split ';') -notcontains $localBin) { $env:Path += ";$localBin" }
}
