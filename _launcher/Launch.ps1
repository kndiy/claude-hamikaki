<#
.SYNOPSIS
    Master runner: register this root on this machine, then start the Windows or
    the WSL Ubuntu launcher.
.DESCRIPTION
    The root is always the folder holding _launcher, so a copied tree works as-is.
    On first run (or after the tree moves) it offers to create a Start Menu shortcut
    "Claude Projects" pointing at <root>\Launch.cmd.

    WSL: the Linux root is separate (config.json: wslDistro, wslRoot). If the
    launcher isn't there yet, it's copied from this tree together with the root
    CLAUDE.md. Projects are not copied.
#>
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Launcher.Common.ps1')

$config = Get-LauncherConfig $PSScriptRoot
$root   = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd('\')
$bridge = Join-Path $PSScriptRoot 'linux\wsl-bridge.sh'

# ---------- register ----------

function Register-Root {
    $stateDir = Join-Path $env:LOCALAPPDATA 'claude-projects-launcher'
    $state    = Join-Path $stateDir 'registered-root.txt'
    $known    = if (Test-Path $state) { (Get-Content $state -Raw).Trim() } else { '' }
    if ($known -eq $root) { return }

    $lnk = Join-Path ([Environment]::GetFolderPath('Programs')) 'Claude Projects.lnk'
    Write-Host "Root on this machine: $root" -ForegroundColor Cyan
    if ((Read-Line "Create Start Menu shortcut 'Claude Projects' for this root? [Y/n]") -notmatch '^[Nn]') {
        $sc = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk)
        $sc.TargetPath       = Join-Path $root 'Launch.cmd'
        $sc.WorkingDirectory = $root
        $sc.Description      = 'Pick a project and open Claude Code'
        $sc.Save()
        Write-Host "Shortcut created: $lnk" -ForegroundColor Green
    }
    New-Item -ItemType Directory -Force -Path $stateDir | Out-Null
    Set-Content -Path $state -Value $root -Encoding UTF8
}

# ---------- WSL ----------

# wsl.exe helpers run with 'Continue': under 'Stop', Windows PowerShell 5.1 turns any
# native stderr output (e.g. "WSL is not installed") into a terminating error.
function Test-WslDistro([string]$Distro) {
    $ErrorActionPreference = 'Continue'
    if (-not (Get-Command wsl.exe -ErrorAction SilentlyContinue)) { return $false }
    $env:WSL_UTF8 = '1'   # newer wsl.exe prints UTF-16 otherwise
    $list = (& wsl.exe -l -q 2>$null | Out-String) -replace "`0", ''
    if ($LASTEXITCODE -ne 0) { return $false }
    @($list -split "`r?`n" | ForEach-Object { $_.Trim() }) -contains $Distro
}

function ConvertTo-WslPath([string]$WinPath) {
    $ErrorActionPreference = 'Continue'
    (& wsl.exe -d $config.wslDistro -- wslpath -a $WinPath | Out-String).Trim()
}

function Invoke-Bridge([string]$Mode) {
    $ErrorActionPreference = 'Continue'
    & wsl.exe -d $config.wslDistro --cd '~' -- bash (ConvertTo-WslPath $bridge) $Mode $config.wslRoot (ConvertTo-WslPath $PSScriptRoot)
}

function Start-Wsl {
    $d = $config.wslDistro
    if (-not (Test-WslDistro $d)) {
        Write-Host "WSL distro '$d' is not installed. In an ADMIN PowerShell run:" -ForegroundColor Yellow
        Write-Host "  wsl --install -d $d"
        Write-Host 'then restart Windows, open Ubuntu once to create your Linux user, and run this again.'
        return
    }
    Invoke-Bridge 'status'
    if ($LASTEXITCODE -ne 0) {
        if ((Read-Line "Set up the launcher in $d at $($config.wslRoot)? [Y/n]") -match '^[Nn]') { return }
        Invoke-Bridge 'deploy'
        if ($LASTEXITCODE -ne 0) { Write-Host 'WSL setup failed.' -ForegroundColor Red; return }
    }
    Invoke-Bridge 'run'
}

# ---------- menu ----------

Set-Location -LiteralPath $root
Register-Root
while ($true) {
    Write-Host "`n== Claude Projects ($root) ==" -ForegroundColor Cyan
    Write-Host '  1) Windows'
    Write-Host "  2) WSL $($config.wslDistro)  (root $($config.wslRoot))"
    Write-Host '  3) Update the WSL launcher from this copy' -ForegroundColor DarkGray
    Write-Host '  Q) quit' -ForegroundColor DarkCyan
    switch -regex ((Read-Line 'Run which version')) {
        '^1$' { & (Join-Path $PSScriptRoot 'Start-Claude.ps1'); return }
        '^2$' { Start-Wsl; return }
        '^3$' { if (Test-WslDistro $config.wslDistro) { Invoke-Bridge 'deploy' } else { Write-Host 'WSL distro not installed.' -ForegroundColor Yellow } }
        '^[Qq]$' { return }
        default { Write-Host 'Invalid choice.' -ForegroundColor Red }
    }
}
