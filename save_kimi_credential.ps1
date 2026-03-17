param(
    [Parameter(Mandatory = $true)][string]$ApiKey,
    [string]$ApiBaseUrl = "https://api.kimi.com/coding/v1",
    [string]$PortalUrl = "https://www.kimi.com/code",
    [string]$DefaultModel = "kimi-for-coding"
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
    -Service "moonshot_kimi" `
    -Fields @{
        api_key = $ApiKey
    } `
    -Metadata @{
        api_base_url = $ApiBaseUrl
        portal_url = $PortalUrl
        default_model = $DefaultModel
        provider = "kimi_code"
        usage_scope = "coding_agents_only"
    }

Write-Output "Credential saved to $entryPath"
