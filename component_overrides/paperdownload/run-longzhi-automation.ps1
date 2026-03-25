param(
  [int]$PageSize = 100,
  [int]$StartPage = 1,
  [string]$OutputRoot = ""
)

$repoRoot = $PSScriptRoot
$workspaceRoot = Split-Path -Parent (Split-Path -Parent $repoRoot)
$credentialStoreScript = Join-Path $workspaceRoot 'scripts\credentials\credential_store.ps1'
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
  $OutputRoot = if ($env:PAPERDOWNLOAD_OUTPUT_ROOT) {
    $env:PAPERDOWNLOAD_OUTPUT_ROOT
  } else {
    Join-Path $repoRoot 'longzhi_batch_output'
  }
}

$runLogFile = Join-Path $OutputRoot 'state\run_log.jsonl'
$stateDir = Join-Path $OutputRoot 'state'
$reportDir = Join-Path $stateDir 'reports'

if (-not (Test-Path $credentialStoreScript)) {
  throw "Credential store script not found: $credentialStoreScript"
}

New-Item -ItemType Directory -Force -Path $reportDir | Out-Null

. $credentialStoreScript
$credential = Read-CredentialStoreEntry -RepoRoot $workspaceRoot -Service 'longzhi'
$plainUsername = [string]$credential.fields.username
$plainPassword = [string]$credential.fields.password

$nodeScript = Join-Path $repoRoot 'longzhi-batch-download.mjs'
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = 'node'
$psi.Arguments = "`"$nodeScript`""
$psi.WorkingDirectory = $repoRoot
$psi.UseShellExecute = $false
$psi.RedirectStandardInput = $true

$psi.Environment['PLAYWRIGHT_BROWSERS_PATH'] = if ($env:PLAYWRIGHT_BROWSERS_PATH) {
  $env:PLAYWRIGHT_BROWSERS_PATH
} else {
  Join-Path $repoRoot '.pw-browsers'
}
$psi.Environment['LZ_CREDENTIAL_STDIN'] = '1'
$psi.Environment['OUTPUT_ROOT'] = [string]$OutputRoot
$psi.Environment['PAGE_SIZE'] = [string]$PageSize
$psi.Environment['START_PAGE'] = [string]$StartPage
$psi.Environment['REVIEW_ENTER_WAIT_MS'] = '10000'
$psi.Environment['POST_VISIBLE_WAIT_MS'] = '4000'
$psi.Environment['HUMAN_READY_WAIT_MIN_MS'] = '1000'
$psi.Environment['HUMAN_READY_WAIT_MAX_MS'] = '2000'
$psi.Environment['DOWNLOAD_STEP_WAIT_MIN_MS'] = '450'
$psi.Environment['DOWNLOAD_STEP_WAIT_MAX_MS'] = '1100'
$psi.Environment['POST_ALL_DOWNLOAD_WAIT_MIN_MS'] = '500'
$psi.Environment['POST_ALL_DOWNLOAD_WAIT_MAX_MS'] = '1200'
$psi.Environment['RETURN_LIST_WAIT_MIN_MS'] = '600'
$psi.Environment['RETURN_LIST_WAIT_MAX_MS'] = '1400'
$psi.Environment['DOWNLOAD_REQUEST_TIMEOUT_MS'] = '120000'
$psi.Environment['DOWNLOAD_RETRY_COUNT'] = '2'
$psi.Environment['NAV_TIMEOUT_MS'] = '90000'
$psi.Environment['NAV_RETRY_COUNT'] = '2'

$process = $null
try {
  $process = [System.Diagnostics.Process]::Start($psi)
  $credentialPayload = @{
    username = $plainUsername
    password = $plainPassword
  } | ConvertTo-Json -Compress
  $process.StandardInput.WriteLine($credentialPayload)
  $process.StandardInput.Close()
  $process.WaitForExit()
} finally {
  $plainUsername = $null
  $plainPassword = $null
}

if (-not $process) {
  throw "Failed to start download process."
}

if ($process.ExitCode -ne 0) {
  throw "Incremental runner failed with exit code $($process.ExitCode)"
}

$lines = Get-Content -Path $runLogFile -Encoding UTF8
$endRecord = $null
for ($i = $lines.Count - 1; $i -ge 0; $i--) {
  $item = $lines[$i] | ConvertFrom-Json
  if ($item.type -eq 'run_end') {
    $endRecord = $item
    break
  }
}

if (-not $endRecord) {
  throw "Could not locate run_end record in $runLogFile"
}

$runId = $endRecord.runId
$downloadEvents = @()
foreach ($line in $lines) {
  $item = $line | ConvertFrom-Json
  if ($item.runId -eq $runId -and $item.type -eq 'student_downloaded') {
    $downloadEvents += $item
  }
}

$newFiles = @()
foreach ($event in $downloadEvents) {
  if ($event.files) {
    $newFiles += @($event.files)
  }
}

$summary = [ordered]@{
  runId = $runId
  status = $endRecord.status
  startedAt = $endRecord.startedAt
  endedAt = $endRecord.endedAt
  processed = $endRecord.summary.processed
  downloaded = $endRecord.summary.downloaded
  skippedExisting = $endRecord.summary.skippedExisting
  noAttachment = $endRecord.summary.noAttachment
  failed = $endRecord.summary.failed
  indexedStudents = $endRecord.summary.indexedStudents
  newFiles = @($newFiles)
}

$summaryJson = $summary | ConvertTo-Json -Depth 5
$summaryJson | Set-Content -Path (Join-Path $stateDir 'latest_automation_summary.json') -Encoding UTF8
$summaryJson | Set-Content -Path (Join-Path $reportDir "$runId.json") -Encoding UTF8

$summaryLines = @(
  "# Longzhi Automation Summary"
  ""
  "- runId: $($summary.runId)"
  "- status: $($summary.status)"
  "- startedAt: $($summary.startedAt)"
  "- endedAt: $($summary.endedAt)"
  "- processed: $($summary.processed)"
  "- downloaded: $($summary.downloaded)"
  "- skippedExisting: $($summary.skippedExisting)"
  "- noAttachment: $($summary.noAttachment)"
  "- failed: $($summary.failed)"
  "- indexedStudents: $($summary.indexedStudents)"
  "- newFiles: $([string]::Join(', ', $summary.newFiles))"
)
$summaryLines -join "`r`n" | Set-Content -Path (Join-Path $stateDir 'latest_automation_summary.md') -Encoding UTF8
($summaryLines -join "`r`n") | Set-Content -Path (Join-Path $reportDir "$runId.md") -Encoding UTF8

$summaryJson

