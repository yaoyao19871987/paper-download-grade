param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PipelineArgs
)

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$workspaceRoot = Split-Path -Parent $projectRoot
$pipelineScript = Join-Path $projectRoot "pipeline.py"
$configPath = Join-Path $projectRoot "pipeline.config.json"

$gradingRoot = Join-Path $workspaceRoot "components\essaygrade"
$venvPython = Join-Path $gradingRoot ".venv\Scripts\python.exe"
$pythonExe = if ($env:PIPELINE_PYTHON) {
    $env:PIPELINE_PYTHON
} elseif (Test-Path $venvPython) {
    $venvPython
} else {
    "python"
}

if (-not $env:PYTHONUTF8) {
    $env:PYTHONUTF8 = "1"
}
if (-not $env:PYTHONIOENCODING) {
    $env:PYTHONIOENCODING = "utf-8"
}

& $pythonExe $pipelineScript --config $configPath @PipelineArgs
exit $LASTEXITCODE
