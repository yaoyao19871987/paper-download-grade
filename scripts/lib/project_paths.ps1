Set-StrictMode -Version Latest

function Get-ProjectRoot {
    param(
        [string]$StartPath = ""
    )

    $candidate = if ([string]::IsNullOrWhiteSpace($StartPath)) {
        $PSScriptRoot
    } else {
        $StartPath
    }

    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
        $candidate = Split-Path -Parent $candidate
    }

    $current = (Resolve-Path -LiteralPath $candidate).Path
    while ($null -ne $current) {
        if (Test-Path -LiteralPath (Join-Path $current ".git")) {
            return $current
        }
        $parent = Split-Path -Parent $current
        if ($parent -eq $current) {
            break
        }
        $current = $parent
    }

    throw "Unable to locate project root from: $StartPath"
}

function Get-ProjectPaths {
    param(
        [string]$StartPath = ""
    )

    $repoRoot = Get-ProjectRoot -StartPath $StartPath
    $runtimeRoot = Join-Path $repoRoot "runtime"

    return [ordered]@{
        RepoRoot              = $repoRoot
        RuntimeRoot           = $runtimeRoot
        ConfigEnvScript       = Join-Path $repoRoot "config\env\project.env.ps1"
        ConfigEnvLocalScript  = Join-Path $repoRoot "config\env\project.env.local.ps1"
        PipelineConfig        = Join-Path $repoRoot "config\pipeline\pipeline.config.json"
        CredentialStoreScript = Join-Path $repoRoot "scripts\credentials\credential_store.ps1"
        SyncOverridesScript   = Join-Path $repoRoot "scripts\bootstrap\sync_component_overrides.ps1"
        TeacherBatchScript    = Join-Path $repoRoot "scripts\run\run_teacher_batch.ps1"
        SetupScript           = Join-Path $repoRoot "scripts\bootstrap\setup_windows.ps1"
    }
}

function Import-ProjectEnvironment {
    param(
        [string]$StartPath = ""
    )

    $paths = Get-ProjectPaths -StartPath $StartPath

    if (-not (Test-Path -LiteralPath $paths.ConfigEnvScript)) {
        throw "Project environment script not found: $($paths.ConfigEnvScript)"
    }

    & $paths.ConfigEnvScript -RepoRoot $paths.RepoRoot

    if (Test-Path -LiteralPath $paths.ConfigEnvLocalScript) {
        & $paths.ConfigEnvLocalScript -RepoRoot $paths.RepoRoot
    }

    return $paths
}
