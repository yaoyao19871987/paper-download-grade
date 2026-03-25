[CmdletBinding()]
param(
    [string]$TeacherName,
    [string]$TargetPageUrl,
    [string]$StageLabel = "初稿",
    [ValidateSet("initial_draft", "final")]
    [string]$Stage = "initial_draft",
    [ValidateSet("auto", "openai", "moonshot", "siliconflow", "expert", "heuristic", "off")]
    [string]$VisualMode = "auto",
    [string]$VisualModel = "gpt-5.4",
    [ValidateSet("off", "auto", "expert", "siliconflow", "moonshot")]
    [string]$TextMode = "expert",
    [string]$TextPrimaryModel = "deepseek-ai/DeepSeek-V3.2",
    [string]$TextSecondaryModel = "kimi-for-coding",
    [int]$MaxStudents = 0,
    [int]$Limit = 0,
    [switch]$QueueGrade,
    [switch]$GradeEvenIfNoNew,
    [switch]$UseActiveSource,
    [switch]$SkipDoctor,
    [switch]$NoBundle,
    [switch]$OverwriteBundle,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "..\lib\project_paths.ps1")
$paths = Import-ProjectEnvironment -StartPath $MyInvocation.MyCommand.Path
$repoRoot = $paths.RepoRoot
$runner = Join-Path $repoRoot "pipeline\run_pipeline.ps1"
$sourceRegistryPath = Join-Path $env:PIPELINE_STATE_DIR "source_registry.json"
$caseExportsRoot = $env:PAPER_PIPELINE_CASE_EXPORTS_DIR

function Get-SafeSourceKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Teacher,
        [Parameter(Mandatory = $true)]
        [string]$StageName
    )

    $folderName = if ([string]::IsNullOrWhiteSpace($Teacher)) {
        $StageName
    } else {
        "$Teacher`_$StageName"
    }
    $safe = [Regex]::Replace($folderName, '[<>:"/\\|?*\x00-\x1F\s]+', "_").Trim("_")
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return "source"
    }
    return $safe
}

function Read-ActiveSource {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RegistryPath
    )

    if (-not (Test-Path -LiteralPath $RegistryPath)) {
        throw "source registry not found: $RegistryPath"
    }
    $raw = Get-Content -Path $RegistryPath -Raw -Encoding UTF8
    $json = $raw | ConvertFrom-Json
    $activeKey = [string]($json.active_source_key)
    if ([string]::IsNullOrWhiteSpace($activeKey)) {
        throw "active_source_key is empty. Run set-source first."
    }

    $source = $json.sources.$activeKey
    if ($null -eq $source) {
        throw "active_source_key '$activeKey' not found in source_registry.json."
    }
    return @{
        key = $activeKey
        folder_name = [string]($source.folder_name)
        teacher_name = [string]($source.teacher_name)
        stage_label = [string]($source.stage_label)
    }
}

function Invoke-PipelineStep {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunnerPath,
        [Parameter(Mandatory = $true)]
        [string]$StepName,
        [Parameter(Mandatory = $true)]
        [string[]]$Args,
        [switch]$NoExec
    )

    $displayArgs = ($Args | ForEach-Object {
            if ($_ -match "\s") { '"' + $_ + '"' } else { $_ }
        }) -join " "
    Write-Host ""
    Write-Host "[$StepName] $displayArgs" -ForegroundColor Cyan

    if ($NoExec) {
        return
    }

    & $RunnerPath @Args
    if ($LASTEXITCODE -ne 0) {
        throw "Step '$StepName' failed, exit code: $LASTEXITCODE"
    }
}

if (-not (Test-Path -LiteralPath $runner)) {
    throw "pipeline runner not found: $runner"
}

if (-not $UseActiveSource) {
    if ([string]::IsNullOrWhiteSpace($TeacherName)) {
        throw "TeacherName is required unless -UseActiveSource is specified."
    }
    if ([string]::IsNullOrWhiteSpace($TargetPageUrl)) {
        throw "TargetPageUrl is required unless -UseActiveSource is specified."
    }
}

$sourceKey = ""
$caseFolderName = ""

Push-Location $repoRoot
try {
    if (-not $SkipDoctor) {
        Invoke-PipelineStep -RunnerPath $runner -StepName "doctor" -Args @("doctor") -NoExec:$DryRun
    }

    if ($UseActiveSource) {
        $active = Read-ActiveSource -RegistryPath $sourceRegistryPath
        $sourceKey = [string]$active.key
        $caseFolderName = [string]$active.folder_name
        Write-Host "Using active source: $sourceKey ($($active.teacher_name) / $($active.stage_label))" -ForegroundColor Yellow
    } else {
        Invoke-PipelineStep -RunnerPath $runner -StepName "set-source" -Args @(
            "set-source",
            "--teacher-name", $TeacherName,
            "--target-page-url", $TargetPageUrl,
            "--stage-label", $StageLabel
        ) -NoExec:$DryRun

        $sourceKey = Get-SafeSourceKey -Teacher $TeacherName -StageName $StageLabel
        $caseFolderName = if ([string]::IsNullOrWhiteSpace($TeacherName)) { $StageLabel } else { "$TeacherName`_$StageLabel" }
    }

    $runAllArgs = @(
        "run-all",
        "--stage", $Stage,
        "--visual-mode", $VisualMode,
        "--visual-model", $VisualModel,
        "--text-mode", $TextMode,
        "--text-primary-model", $TextPrimaryModel,
        "--text-secondary-model", $TextSecondaryModel
    )
    if ($MaxStudents -gt 0) {
        $runAllArgs += @("--max-students", "$MaxStudents")
    }
    if ($Limit -gt 0) {
        $runAllArgs += @("--limit", "$Limit")
    }
    if ($QueueGrade) {
        $runAllArgs += "--queue-grade"
    }
    if ($GradeEvenIfNoNew) {
        $runAllArgs += "--grade-even-if-no-new"
    }

    Invoke-PipelineStep -RunnerPath $runner -StepName "run-all" -Args $runAllArgs -NoExec:$DryRun
    Invoke-PipelineStep -RunnerPath $runner -StepName "refresh-log" -Args @("refresh-log") -NoExec:$DryRun

    if (-not $NoBundle) {
        $bundleArgs = @("bundle-source", "--source-key", $sourceKey)
        if ($OverwriteBundle) {
            $bundleArgs += "--overwrite"
        }
        Invoke-PipelineStep -RunnerPath $runner -StepName "bundle-source" -Args $bundleArgs -NoExec:$DryRun
    }
}
finally {
    Pop-Location
}

Write-Host ""
if ($DryRun) {
    Write-Host "Dry run completed." -ForegroundColor Green
} else {
    Write-Host "Batch pipeline completed." -ForegroundColor Green
}
if (-not [string]::IsNullOrWhiteSpace($caseFolderName)) {
    $bundlePath = Join-Path $caseExportsRoot $caseFolderName
    Write-Host "Bundle folder: $bundlePath"
}
