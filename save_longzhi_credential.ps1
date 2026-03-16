param(
    [Parameter(Mandatory = $true)][string]$Username,
    [Parameter(Mandatory = $true)][string]$Password
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$downloadRoot = Join-Path $repoRoot "components\paperdownload"
$scriptPath = Join-Path $downloadRoot "save-longzhi-credential.ps1"
$outputRoot = Join-Path $downloadRoot "longzhi_batch_output"

if (-not (Test-Path $scriptPath)) {
    throw "Script not found: $scriptPath. Run .\setup_windows.ps1 first."
}

PowerShell -ExecutionPolicy Bypass -File $scriptPath `
    -Username $Username `
    -Password $Password `
    -OutputRoot $outputRoot

if ($LASTEXITCODE -ne 0) {
    throw "Failed to save credential."
}
