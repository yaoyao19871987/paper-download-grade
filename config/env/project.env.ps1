param(
    [Parameter(Mandatory = $true)]
    [string]$RepoRoot
)

Set-StrictMode -Version Latest

$runtimeRoot = Join-Path $RepoRoot "runtime"
$runtimeDirs = @(
    $runtimeRoot,
    (Join-Path $runtimeRoot "downloads\longzhi_batch_output"),
    (Join-Path $runtimeRoot "downloads\longzhi_batch_output\downloads"),
    (Join-Path $runtimeRoot "downloads\longzhi_batch_output\artifacts"),
    (Join-Path $runtimeRoot "downloads\longzhi_batch_output\state"),
    (Join-Path $runtimeRoot "grading\incoming_papers"),
    (Join-Path $runtimeRoot "grading\runs"),
    (Join-Path $runtimeRoot "pipeline\state"),
    (Join-Path $runtimeRoot "tracking\student_feedback"),
    (Join-Path $runtimeRoot "tracking"),
    (Join-Path $runtimeRoot "exports\case_exports"),
    (Join-Path $runtimeRoot "locks"),
    (Join-Path $runtimeRoot "secrets\credential_store")
)

foreach ($dir in $runtimeDirs) {
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
}

$env:PAPER_PIPELINE_REPO_ROOT = $RepoRoot
$env:PAPER_PIPELINE_CONFIG = Join-Path $RepoRoot "config\pipeline\pipeline.config.json"
$env:PAPER_PIPELINE_CREDENTIAL_STORE_DIR = Join-Path $runtimeRoot "secrets\credential_store"
$env:PAPERDOWNLOAD_OUTPUT_ROOT = Join-Path $runtimeRoot "downloads\longzhi_batch_output"
$env:ESSAYGRADE_INCOMING_DIR = Join-Path $runtimeRoot "grading\incoming_papers"
$env:ESSAYGRADE_RUNS_DIR = Join-Path $runtimeRoot "grading\runs"
$env:PIPELINE_STATE_DIR = Join-Path $runtimeRoot "pipeline\state"
$env:PAPER_PIPELINE_FEEDBACK_DIR = Join-Path $runtimeRoot "tracking\student_feedback"
$env:PAPER_PIPELINE_STUDENT_LOG_JSON = Join-Path $runtimeRoot "tracking\student_progress_log.json"
$env:PAPER_PIPELINE_STUDENT_LOG_MD = Join-Path $runtimeRoot "tracking\student_progress_log.md"
$env:PAPER_PIPELINE_CASE_EXPORTS_DIR = Join-Path $runtimeRoot "exports\case_exports"
$env:ESSAYGRADE_WORD_LOCK_FILE = Join-Path $runtimeRoot "locks\essaygrade_word_export.lock"

if (-not $env:PLAYWRIGHT_BROWSERS_PATH) {
    $env:PLAYWRIGHT_BROWSERS_PATH = Join-Path $RepoRoot "components\paperdownload\.pw-browsers"
}
if (-not $env:PYTHONUTF8) {
    $env:PYTHONUTF8 = "1"
}
if (-not $env:PYTHONIOENCODING) {
    $env:PYTHONIOENCODING = "utf-8"
}
