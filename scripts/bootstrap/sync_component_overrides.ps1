param(
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "..\lib\project_paths.ps1")
$paths = Get-ProjectPaths -StartPath $MyInvocation.MyCommand.Path
$repoRoot = $paths.RepoRoot
$overridesRoot = Join-Path $repoRoot "component_overrides"
$componentsRoot = Join-Path $repoRoot "components"

if (-not (Test-Path -LiteralPath $overridesRoot)) {
    if (-not $Quiet) {
        Write-Host "No component overrides found."
    }
    return
}

if (-not (Test-Path -LiteralPath $componentsRoot)) {
    throw "Components directory not found: $componentsRoot"
}

$copiedCount = 0
Get-ChildItem -LiteralPath $overridesRoot -Recurse -File | ForEach-Object {
    $relativePath = $_.FullName.Substring($overridesRoot.Length).TrimStart('\')
    $targetPath = Join-Path $componentsRoot $relativePath
    $targetDir = Split-Path -Parent $targetPath
    if (-not (Test-Path -LiteralPath $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }
    Copy-Item -LiteralPath $_.FullName -Destination $targetPath -Force
    $copiedCount += 1
}

if (-not $Quiet) {
    Write-Host "Applied $copiedCount component override files."
}
