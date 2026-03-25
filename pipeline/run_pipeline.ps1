param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PipelineArgs
)

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $projectRoot
$pipelineScript = Join-Path $projectRoot "pipeline.py"
$configPath = Join-Path $repoRoot "config\pipeline\pipeline.config.json"

. (Join-Path $repoRoot "scripts\lib\project_paths.ps1")
$null = Import-ProjectEnvironment -StartPath $MyInvocation.MyCommand.Path

$gradingRoot = Join-Path $repoRoot "components\essaygrade"
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
