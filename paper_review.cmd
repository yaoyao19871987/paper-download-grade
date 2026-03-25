@echo off
setlocal
set "CODEX_RUN_TEACHER_BATCH_ARGS=%*"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$argsLine = [Environment]::GetEnvironmentVariable('CODEX_RUN_TEACHER_BATCH_ARGS', 'Process'); Invoke-Expression ""& '%~dp0run_teacher_batch.ps1' $argsLine""; exit $LASTEXITCODE"
set "EXITCODE=%ERRORLEVEL%"
endlocal & exit /b %EXITCODE%
