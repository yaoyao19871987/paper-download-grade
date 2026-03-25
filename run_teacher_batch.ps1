param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ForwardArgs
)

$ErrorActionPreference = "Stop"

$target = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "scripts\run\run_teacher_batch.ps1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $target @ForwardArgs
exit $LASTEXITCODE
