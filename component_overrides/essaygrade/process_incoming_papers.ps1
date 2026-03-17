param(
    [ValidateSet("initial_draft", "final")]
    [string]$Stage = "initial_draft",

    [ValidateSet("auto", "openai", "moonshot", "siliconflow", "expert", "heuristic", "off")]
    [string]$VisualMode = "auto",

    [string]$VisualModel = "gpt-5.4",

    [int]$Limit = 0
)

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$incomingDir = Join-Path $projectRoot "assets\incoming_papers"
$runsDir = Join-Path $projectRoot "grading_runs"
$runScript = Join-Path $projectRoot "run_grade.ps1"

if (-not (Test-Path $incomingDir)) {
    throw "Incoming papers directory not found: $incomingDir"
}

function Get-ProcessedPaperPaths {
    $processed = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if (-not (Test-Path $runsDir)) {
        return $processed
    }

    Get-ChildItem -Path $runsDir -Recurse -Filter run_info.txt -ErrorAction SilentlyContinue | ForEach-Object {
        $paperLine = Select-String -Path $_.FullName -Pattern '^paper=' -SimpleMatch:$false | Select-Object -First 1
        if ($paperLine) {
            $paperPath = $paperLine.Line.Substring(6).Trim()
            if ($paperPath) {
                $null = $processed.Add($paperPath)
            }
        }
    }

    return $processed
}

function Get-IncomingPaperFiles {
    Get-ChildItem -Path $incomingDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @('.doc', '.docx') } |
        Sort-Object LastWriteTime, Name
}

$processedPaths = Get-ProcessedPaperPaths
$incomingFiles = Get-IncomingPaperFiles

if (-not $incomingFiles) {
    Write-Host "No incoming papers found in $incomingDir"
    exit 0
}

$pendingFiles = @()
foreach ($file in $incomingFiles) {
    $resolved = (Resolve-Path $file.FullName).Path
    if (-not $processedPaths.Contains($resolved)) {
        $pendingFiles += $file
    }
}

if ($Limit -gt 0) {
    $pendingFiles = $pendingFiles | Select-Object -First $Limit
}

if (-not $pendingFiles) {
    Write-Host "No unprocessed papers found."
    exit 0
}

$results = @()
foreach ($file in $pendingFiles) {
    $label = ($file.BaseName -replace '[^0-9A-Za-z_-]', '_').Trim('_')
    if ([string]::IsNullOrWhiteSpace($label)) {
        $label = "paper"
    }

    Write-Host "Processing: $($file.FullName)"
    & PowerShell -ExecutionPolicy Bypass -File $runScript `
        -PaperPath $file.FullName `
        -RunLabel $label `
        -Stage $Stage `
        -VisualMode $VisualMode `
        -VisualModel $VisualModel

    if ($LASTEXITCODE -ne 0) {
        throw "Grading failed for $($file.FullName)"
    }

    $results += $file.FullName
}

Write-Host ""
Write-Host "Processed papers:"
$results | ForEach-Object { Write-Host "- $_" }
