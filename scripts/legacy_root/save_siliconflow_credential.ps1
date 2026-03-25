param(
    [Parameter(Mandatory = $true)][string]$ApiKey,
    [string]$ApiBaseUrl = "https://api.siliconflow.cn/v1",
    [string]$PortalUrl = "https://cloud.siliconflow.cn",
    [string]$DefaultModel = "Pro/moonshotai/Kimi-K2.5"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$credentialStoreScript = Join-Path $repoRoot "credential_store.ps1"

if (-not (Test-Path $credentialStoreScript)) {
    throw "Credential store script not found: $credentialStoreScript"
}

. $credentialStoreScript

$entryPath = Save-CredentialStoreEntry `
    -RepoRoot $repoRoot `
    -Service "siliconflow" `
    -Fields @{
        api_key = $ApiKey
    } `
    -Metadata @{
        api_base_url = $ApiBaseUrl
        portal_url = $PortalUrl
        default_model = $DefaultModel
        provider = "siliconflow"
    }

Write-Output "Credential saved to $entryPath"
