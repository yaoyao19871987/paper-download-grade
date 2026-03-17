param(
    [Parameter(Mandatory = $true)]
    [string]$PaperPath,

    [string]$RunLabel,

    [ValidateSet("initial_draft", "final")]
    [string]$Stage = "initial_draft",

    [ValidateSet("auto", "openai", "moonshot", "siliconflow", "expert", "heuristic", "off")]
    [string]$VisualMode = "auto",

    [string]$VisualModel = "gpt-5.4",

    [string[]]$ReferenceDoc = @()
)

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
$pythonExe = if ($env:ESSAYGRADE_PYTHON) {
    $env:ESSAYGRADE_PYTHON
} elseif (Test-Path $venvPython) {
    $venvPython
} else {
    "python"
}
$resolvedPaperPath = (Resolve-Path $PaperPath).Path
$timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
$defaultLabel = [System.IO.Path]::GetFileNameWithoutExtension($resolvedPaperPath)
$labelSeed = if ([string]::IsNullOrWhiteSpace($RunLabel)) { $defaultLabel } else { $RunLabel }
$safeLabel = ($labelSeed -replace '[^0-9A-Za-z_-]', '_').Trim('_')
if ([string]::IsNullOrWhiteSpace($safeLabel)) {
    $safeLabel = "paper"
}
$runRoot = Join-Path $projectRoot ("grading_runs\" + $timestamp + "_" + $safeLabel)

$dirs = @(
    $runRoot,
    (Join-Path $runRoot "source"),
    (Join-Path $runRoot "reports"),
    (Join-Path $runRoot "json"),
    (Join-Path $runRoot "notes"),
    (Join-Path $runRoot "visual")
)

foreach ($dir in $dirs) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
    }
}

$paperName = [System.IO.Path]::GetFileName($resolvedPaperPath)
$paperCopyPath = Join-Path $runRoot "source\$paperName"
Copy-Item -Force $resolvedPaperPath $paperCopyPath

$jsonOut = Join-Path $runRoot "json\grade_result.json"
$textOut = Join-Path $runRoot "reports\grade_report.txt"
$visualOut = Join-Path $runRoot "visual"

$arguments = @(
    (Join-Path $projectRoot "app\grade_paper.py"),
    $resolvedPaperPath,
    "--stage",
    $Stage,
    "--visual-mode",
    $VisualMode,
    "--visual-model",
    $VisualModel,
    "--visual-output-dir",
    $visualOut,
    "--json-out",
    $jsonOut,
    "--text-out",
    $textOut
)

foreach ($reference in $ReferenceDoc) {
    if ($reference -and (Test-Path $reference)) {
        $resolvedReference = (Resolve-Path $reference).Path
        if ($resolvedReference -eq $resolvedPaperPath) {
            continue
        }
        $arguments += "--reference-doc"
        $arguments += $resolvedReference
    }
}

& $pythonExe @arguments
if ($LASTEXITCODE -ne 0) {
    throw "Grade script failed with exit code: $LASTEXITCODE"
}

$noteLines = @(
    "run_root=$runRoot",
    "paper=$resolvedPaperPath",
    "paper_copy=$paperCopyPath",
    "stage=$Stage",
    "visual_mode=$VisualMode",
    "visual_model=$VisualModel",
    "visual_dir=$visualOut",
    "json=$jsonOut",
    "report=$textOut",
    "references=$($ReferenceDoc -join ';')"
)
$noteLines | Set-Content -Encoding UTF8 (Join-Path $runRoot "notes\run_info.txt")

$repoRoot = Split-Path -Parent (Split-Path -Parent $projectRoot)
$pipelineScript = Join-Path $repoRoot "pipeline\run_pipeline.ps1"
if (Test-Path $pipelineScript) {
    try {
        & powershell.exe -ExecutionPolicy Bypass -File $pipelineScript refresh-log | Out-Null
    } catch {
        Write-Warning ("Tracking refresh failed: " + $_.Exception.Message)
    }
}

Write-Host "Grading completed."
Write-Host "Run directory: $runRoot"
Write-Host "Text report: $textOut"
Write-Host "JSON report: $jsonOut"
