param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PipelineArgs
)

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pipelineScript = Join-Path $projectRoot "pipeline.py"
$configPath = Join-Path $projectRoot "pipeline.config.json"

$gradingRoot = Join-Path (Split-Path $projectRoot -Parent) "paper-grading-system"
$venvPython = Join-Path $gradingRoot ".venv\Scripts\python.exe"
$pythonExe = if (Test-Path $venvPython) { $venvPython } else { "python" }

& $pythonExe $pipelineScript --config $configPath @PipelineArgs
exit $LASTEXITCODE
