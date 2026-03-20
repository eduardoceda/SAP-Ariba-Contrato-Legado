$ErrorActionPreference = 'Stop'

$frontendPath = (Resolve-Path (Join-Path $PSScriptRoot '..\frontend')).Path
Set-Location $frontendPath

if (-not (Test-Path 'node_modules')) {
    npm install
}

$env:VITE_API_BASE_URL = 'http://localhost:8001/api'
npm run dev
