<#
.SYNOPSIS
    Installs the missing tools listed in one CLAUDE.md's Windows tools table.
.DESCRIPTION
    Runs each tool's Install command from the folder that holds the CLAUDE.md, so
    project-relative commands like "python -m pip install -r requirements.txt" work.
    A tool that can't be installed automatically (manual: entries, winget missing,
    or the install failed) is listed at the end with the command to run yourself.
    Exit code 1 when anything is still missing.
.PARAMETER ClaudeMd
    The CLAUDE.md whose table to use.
.PARAMETER Names
    Only these tools. Default: every tool in the table that is missing.
.EXAMPLE
    .\Install-Tools.ps1 -ClaudeMd ..\CLAUDE.md -Names 'Git for Windows'
#>
param(
    [Parameter(Mandatory)][string]$ClaudeMd,
    [string[]]$Names
)
. (Join-Path $PSScriptRoot 'Launcher.Common.ps1')

$ClaudeMd = (Resolve-Path -LiteralPath $ClaudeMd).Path
$tools = @(Get-ToolTable $ClaudeMd | Where-Object { if ($Names) { $Names -contains $_.Name } else { -not (Get-ToolVersion $_.Check $_.Dir) } })
$failed = @()

foreach ($t in $tools) {
    Write-Host "`n>> $($t.Name): $($t.Install)" -ForegroundColor Cyan
    if ($t.Install -match '^manual:') {
        $failed += $t; continue
    }
    if ($t.Install -match '^winget\b' -and -not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host 'winget is not available.' -ForegroundColor Yellow
        $failed += $t; continue
    }
    Push-Location -LiteralPath $t.Dir
    try {
        Invoke-Expression $t.Install
    } catch {
        Write-Host "Install failed: $_" -ForegroundColor Red
    } finally {
        Pop-Location
    }
    Update-SessionPath
    if (Get-ToolVersion $t.Check $t.Dir) {
        Write-Host "$($t.Name) installed." -ForegroundColor Green
    } else {
        $failed += $t
    }
}

if ($failed) {
    Write-Host "`nThese tools could not be installed automatically. Run the commands yourself in PowerShell, then open a NEW terminal:" -ForegroundColor Yellow
    foreach ($t in $failed) {
        Write-Host ("  {0,-20} (in {1})" -f $t.Name, $t.Dir)
        Write-Host ("      {0}" -f ($t.Install -replace '^manual:\s*', ''))
    }
    exit 1
}
exit 0
