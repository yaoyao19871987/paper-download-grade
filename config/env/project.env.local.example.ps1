param(
    [Parameter(Mandatory = $true)]
    [string]$RepoRoot
)

# Copy this file to project.env.local.ps1 for machine-specific overrides.
# Example:
# $env:PIPELINE_PYTHON = "D:\Python311\python.exe"
# $env:ESSAYGRADE_PYTHON = "D:\CodeX\paper-download-grade\components\essaygrade\.venv\Scripts\python.exe"
# $env:OPENAI_API_KEY = "<your-openai-key>"
# $env:MOONSHOT_API_KEY = "<your-moonshot-key>"
# $env:SILICONFLOW_API_KEY = "<your-siliconflow-key>"
