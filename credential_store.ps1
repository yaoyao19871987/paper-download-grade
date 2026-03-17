Set-StrictMode -Version Latest

function Get-CredentialStoreRoot {
    param(
        [string]$RepoRoot = ""
    )

    if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
        $RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    return (Join-Path $RepoRoot ".credential_store")
}

function Protect-CredentialStorePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolvedPath = if (Test-Path -LiteralPath $Path) {
        (Resolve-Path -LiteralPath $Path).Path
    } else {
        $Path
    }

    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $item = Get-Item -Force -LiteralPath $resolvedPath
    $acl = Get-Acl -LiteralPath $resolvedPath
    $acl.SetAccessRuleProtection($true, $false)

    foreach ($rule in @($acl.Access)) {
        [void]$acl.RemoveAccessRuleAll($rule)
    }

    $inheritanceFlags = if ($item.PSIsContainer) {
        [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit"
    } else {
        [System.Security.AccessControl.InheritanceFlags]::None
    }
    $propagationFlags = [System.Security.AccessControl.PropagationFlags]::None
    $allowedIdentities = @(
        $currentUser,
        'BUILTIN\Administrators',
        'NT AUTHORITY\SYSTEM'
    )

    foreach ($identity in $allowedIdentities) {
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $identity,
            [System.Security.AccessControl.FileSystemRights]::FullControl,
            $inheritanceFlags,
            $propagationFlags,
            [System.Security.AccessControl.AccessControlType]::Allow
        )
        [void]$acl.AddAccessRule($rule)
    }

    Set-Acl -LiteralPath $resolvedPath -AclObject $acl
    $item.Attributes = ($item.Attributes -bor [System.IO.FileAttributes]::Hidden)
}

function Protect-SecretValue {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value)

    $secure = ConvertTo-SecureString $Value -AsPlainText -Force
    return ConvertFrom-SecureString $secure
}

function Unprotect-SecretValue {
    param([Parameter(Mandatory = $true)][string]$CipherText)

    $secure = ConvertTo-SecureString $CipherText
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    } finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Save-CredentialStoreEntry {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$Service,
        [Parameter(Mandatory = $true)][hashtable]$Fields,
        [hashtable]$Metadata = @{}
    )

    $storeRoot = Get-CredentialStoreRoot -RepoRoot $RepoRoot
    if (-not (Test-Path -LiteralPath $storeRoot)) {
        New-Item -ItemType Directory -Path $storeRoot | Out-Null
    }
    Protect-CredentialStorePath -Path $storeRoot

    $entryPath = Join-Path $storeRoot ("{0}.json" -f $Service)
    $encryptedFields = [ordered]@{}
    foreach ($key in $Fields.Keys) {
        $encryptedFields[$key] = Protect-SecretValue -Value ([string]$Fields[$key])
    }

    $payload = [ordered]@{
        version = 1
        service = $Service
        saved_at = (Get-Date).ToString('s')
        protection = 'DPAPI_CURRENT_USER'
        metadata = $Metadata
        fields = $encryptedFields
    }

    $payload | ConvertTo-Json -Depth 8 | Set-Content -Path $entryPath -Encoding UTF8
    Protect-CredentialStorePath -Path $entryPath
    return $entryPath
}

function Read-CredentialStoreEntry {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$Service
    )

    $storeRoot = Get-CredentialStoreRoot -RepoRoot $RepoRoot
    $entryPath = Join-Path $storeRoot ("{0}.json" -f $Service)
    if (-not (Test-Path -LiteralPath $entryPath)) {
        throw "Credential entry not found: $entryPath"
    }

    $payload = Get-Content -LiteralPath $entryPath -Encoding UTF8 -Raw | ConvertFrom-Json
    if ($payload.protection -and $payload.protection -ne 'DPAPI_CURRENT_USER') {
        throw "Unsupported credential protection mode: $($payload.protection)"
    }

    $decryptedFields = @{}
    foreach ($prop in $payload.fields.PSObject.Properties) {
        $decryptedFields[$prop.Name] = Unprotect-SecretValue -CipherText ([string]$prop.Value)
    }

    return [pscustomobject]@{
        service = $payload.service
        saved_at = $payload.saved_at
        metadata = $payload.metadata
        fields = [pscustomobject]$decryptedFields
        path = $entryPath
    }
}
