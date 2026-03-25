param(
    [Parameter(Mandatory = $true)][string]$Username,
    [Parameter(Mandatory = $true)][string]$Password
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$downloadRoot = Join-Path $repoRoot "components\paperdownload"
$legacyCredentialPath = Join-Path $downloadRoot "longzhi_batch_output\state\longzhi_credential.json"
$credentialStoreScript = Join-Path $repoRoot "credential_store.ps1"

if (-not (Test-Path $credentialStoreScript)) {
    throw "Credential store script not found: $credentialStoreScript"
}

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
