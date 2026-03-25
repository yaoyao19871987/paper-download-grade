param(
    [Parameter(Mandatory = $true)][string]$ApiKey,
    [string]$ApiBaseUrl = "https://api.kimi.com/coding/v1",
    [string]$PortalUrl = "https://www.kimi.com/code",
    [string]$DefaultModel = "kimi-for-coding"
)

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "..\lib\project_paths.ps1")
$paths = Import-ProjectEnvironment -StartPath $MyInvocation.MyCommand.Path
$repoRoot = $paths.RepoRoot
$credentialStoreScript = $paths.CredentialStoreScript

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
