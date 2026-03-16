param(
    [string]$PaperDownloadRepoUrl = "https://github.com/yaoyao19871987/paperdownload.git",
    [string]$EssayGradeRepoUrl = "https://github.com/yaoyao19871987/essaygrade.git",
    [switch]$SkipPull,
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$componentsDir = Join-Path $repoRoot "components"
$paperdownloadDir = Join-Path $componentsDir "paperdownload"
$essaygradeDir = Join-Path $componentsDir "essaygrade"

function Ensure-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Ensure-GitRepo {
    param(
        [Parameter(Mandatory = $true)][string]$RepoDir,
        [Parameter(Mandatory = $true)][string]$RepoUrl,
        [Parameter(Mandatory = $false)][switch]$NoPull
    )

    if (-not (Test-Path $RepoDir)) {
        git clone $RepoUrl $RepoDir
        if ($LASTEXITCODE -ne 0) {
            throw "Clone failed: $RepoUrl"
        }
        return
    }

    if (-not (Test-Path (Join-Path $RepoDir ".git"))) {
        throw "Path exists but is not a git repo: $RepoDir"
    }

    if (-not $NoPull) {
        git -C $RepoDir pull --ff-only
        if ($LASTEXITCODE -ne 0) {
            throw "git pull failed in $RepoDir"
        }
    }
}

function Resolve-CommandName {
    param([Parameter(Mandatory = $true)][string[]]$Candidates)
    foreach ($name in $Candidates) {
        if (Get-Command $name -ErrorAction SilentlyContinue) {
            return $name
        }
    }
    return $null
}

Ensure-Directory -Path $componentsDir

Ensure-GitRepo -RepoDir $paperdownloadDir -RepoUrl $PaperDownloadRepoUrl -NoPull:$SkipPull
Ensure-GitRepo -RepoDir $essaygradeDir -RepoUrl $EssayGradeRepoUrl -NoPull:$SkipPull

if (-not $SkipInstall) {
    $npmCmd = Resolve-CommandName -Candidates @("npm.cmd", "npm")
    $npxCmd = Resolve-CommandName -Candidates @("npx.cmd", "npx")
    if (-not $npmCmd -or -not $npxCmd) {
        throw "npm/npx not found. Please install Node.js 18+ first."
    }

    Push-Location $paperdownloadDir
    try {
        & $npmCmd install
        if ($LASTEXITCODE -ne 0) {
            throw "npm install failed in paperdownload"
        }
        $env:PLAYWRIGHT_BROWSERS_PATH = Join-Path $paperdownloadDir ".pw-browsers"
        & $npxCmd playwright install chromium
        if ($LASTEXITCODE -ne 0) {
            throw "playwright install failed in paperdownload"
        }
    } finally {
        Pop-Location
    }

    Push-Location $essaygradeDir
    try {
        $setupScript = Join-Path $essaygradeDir "setup_windows_env.ps1"
        if (Test-Path $setupScript) {
            PowerShell -ExecutionPolicy Bypass -File $setupScript
            if ($LASTEXITCODE -ne 0) {
                throw "setup_windows_env.ps1 failed in essaygrade"
            }
        } else {
            $pythonCmd = Resolve-CommandName -Candidates @("python", "py")
            if (-not $pythonCmd) {
                throw "Python not found. Please install Python 3.11+ first."
            }
            & $pythonCmd -m venv .venv
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to create venv in essaygrade"
            }
            & ".\.venv\Scripts\python.exe" -m pip install --upgrade pip
            if (Test-Path ".\requirements.txt") {
                & ".\.venv\Scripts\python.exe" -m pip install -r .\requirements.txt
            }
        }
    } finally {
        Pop-Location
    }
}

Write-Host ""
Write-Host "Setup completed."
Write-Host "1) Save credential (one-time):"
Write-Host "   PowerShell -ExecutionPolicy Bypass -File .\save_longzhi_credential.ps1 -Username `"YOUR_USER`" -Password `"YOUR_PASS`""
Write-Host "2) Health check:"
Write-Host "   PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 doctor"
Write-Host "3) Trial run (1 student):"
Write-Host "   PowerShell -ExecutionPolicy Bypass -File .\pipeline\run_pipeline.ps1 run-all --max-students 1 --stage initial_draft --visual-mode heuristic --limit 1"
