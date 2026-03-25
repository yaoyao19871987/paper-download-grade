param(
  [Parameter(Mandatory = $true)][string]$Username,
  [Parameter(Mandatory = $true)][string]$Password,
  [string]$OutputRoot = ""
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$rootScript = Join-Path $repoRoot 'scripts\credentials\save_longzhi_credential.ps1'

if (-not (Test-Path $rootScript)) {
  throw "Root credential script not found: $rootScript"
}

PowerShell -ExecutionPolicy Bypass -File $rootScript `
  -Username $Username `
  -Password $Password

if ($LASTEXITCODE -ne 0) {
  throw "Failed to save Longzhi credential."
}
