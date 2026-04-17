@echo off
REM AEDI-S launcher shim for cmd.exe — delegates to launch.py.
setlocal
set "REPO=%~dp0.."
where python >nul 2>&1 && set "PY=python" || set "PY=python3"
%PY% "%REPO%\launch.py" %*
exit /b %ERRORLEVEL%
