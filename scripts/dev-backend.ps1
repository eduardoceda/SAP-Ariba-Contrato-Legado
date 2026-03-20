$ErrorActionPreference = 'Stop'

$backendPath = (Resolve-Path (Join-Path $PSScriptRoot '..\backend')).Path
Set-Location $backendPath

$venvPython = Join-Path $backendPath '.venv\Scripts\python.exe'
if (-not (Test-Path $venvPython)) {
    python -m venv .venv
}

& $venvPython -m pip install -r requirements.txt
& $venvPython -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8001
