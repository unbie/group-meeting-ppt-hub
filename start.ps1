$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
if (-not (Test-Path '.venv\Scripts\python.exe')) { throw 'Create .venv and install dependencies as described in README.md first.' }
if (-not (Test-Path 'config.json') -and -not ($env:PPT_HUB_SECRET_KEY -and $env:PPT_HUB_ADMIN_PASSWORD_HASH)) { throw 'Run .venv\Scripts\python.exe setup.py first.' }
& .venv\Scripts\python.exe app.py
