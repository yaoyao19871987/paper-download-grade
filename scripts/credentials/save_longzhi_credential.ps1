param(
    [Parameter(Mandatory = $true)][string]$Username,
    [Parameter(Mandatory = $true)][string]$Password
)

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "..\lib\project_paths.ps1")
$paths = Import-ProjectEnvironment -StartPath $MyInvocation.MyCommand.Path
$repoRoot = $paths.RepoRoot
$legacyCredentialPath = Join-Path $env:PAPERDOWNLOAD_OUTPUT_ROOT "state\longzhi_credential.json"
$credentialStoreScript = $paths.CredentialStoreScript

. $credentialStoreScript

$entryPath = Save-CredentialStoreEntry `
    -RepoRoot $repoRoot `
    -Service "longzhi" `
    -Fields @{
        username = $Username
        password = $Password
    } `
    -Metadata @{
        login_url = "http://longzhi.net.cn/"
        note = "Longzhi download account"
    }

if (Test-Path -LiteralPath $legacyCredentialPath) {
    Remove-Item -Force -LiteralPath $legacyCredentialPath
}

Write-Output "Credential saved to $entryPath"
