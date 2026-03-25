param(
    [ValidateSet("initial_draft", "final")]
    [string]$Stage = "initial_draft",

    [ValidateSet("auto", "openai", "moonshot", "siliconflow", "expert", "heuristic", "off")]
    [string]$VisualMode = "auto",

    [string]$VisualModel = "gpt-5.4",

    [ValidateSet("off", "auto", "expert", "siliconflow", "moonshot")]
    [string]$TextMode = "expert",

    [string]$TextPrimaryModel = "deepseek-ai/DeepSeek-V3.2",

    [string]$TextSecondaryModel = "kimi-for-coding",

    [int]$Limit = 0
)

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$incomingDir = if ($env:ESSAYGRADE_INCOMING_DIR) { $env:ESSAYGRADE_INCOMING_DIR } else { Join-Path $projectRoot "assets\incoming_papers" }
$runsDir = if ($env:ESSAYGRADE_RUNS_DIR) { $env:ESSAYGRADE_RUNS_DIR } else { Join-Path $projectRoot "grading_runs" }
$runScript = Join-Path $projectRoot "run_grade.ps1"

if (-not (Test-Path $incomingDir)) {
    throw "Incoming papers directory not found: $incomingDir"
}

function Get-ProcessedPaperPaths {
    $processed = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $processedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if (-not (Test-Path $runsDir)) {
        return @{
            Paths = $processed
            Names = $processedNames
        }
    }

    Get-ChildItem -Path $runsDir -Recurse -Filter run_info.txt -ErrorAction SilentlyContinue | ForEach-Object {
        $paperLine = Select-String -Path $_.FullName -Pattern '^paper=' -SimpleMatch:$false | Select-Object -First 1
        if ($paperLine) {
            $paperPath = $paperLine.Line.Substring(6).Trim()
            if ($paperPath) {
                $null = $processed.Add($paperPath)
                $paperName = [System.IO.Path]::GetFileName($paperPath)
                if (-not [string]::IsNullOrWhiteSpace($paperName)) {
                    $null = $processedNames.Add($paperName)
                }
            }
        }
    }

    return @{
        Paths = $processed
        Names = $processedNames
    }
}

function Get-IncomingPaperFiles {
    Get-ChildItem -Path $incomingDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @('.doc', '.docx') } |
        Sort-Object LastWriteTime, Name
}

$processedInfo = Get-ProcessedPaperPaths
$processedPaths = $processedInfo.Paths
$processedNames = $processedInfo.Names
$incomingFiles = Get-IncomingPaperFiles

if (-not $incomingFiles) {
    Write-Host "No incoming papers found in $incomingDir"
    exit 0
}

$pendingFiles = @()
foreach ($file in $incomingFiles) {
    $resolved = (Resolve-Path $file.FullName).Path
    if (-not $processedPaths.Contains($resolved) -and -not $processedNames.Contains($file.Name)) {
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
        -VisualModel $VisualModel `
        -TextMode $TextMode `
        -TextPrimaryModel $TextPrimaryModel `
        -TextSecondaryModel $TextSecondaryModel

    if ($LASTEXITCODE -ne 0) {
        throw "Grading failed for $($file.FullName)"
    }

    $results += $file.FullName
}

Write-Host ""
Write-Host "Processed papers:"
$results | ForEach-Object { Write-Host "- $_" }
