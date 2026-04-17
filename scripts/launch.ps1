#!/usr/bin/env pwsh
# AEDI-S launcher shim — delegates to cross-platform launch.py.
# Windows PowerShell 7+ and pwsh on macOS/Linux.
$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
$py = if (Get-Command python -ErrorAction SilentlyContinue) { "python" } else { "python3" }
& $py (Join-Path $repo "launch.py") @args
exit $LASTEXITCODE
