# Starts the SLIDIFY app and opens it in your browser.
# Run from anywhere:  powershell -ExecutionPolicy Bypass -File <repo>\app\start.ps1

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot

if (-not (Get-Command claude -ErrorAction SilentlyContinue)) {
    Write-Host "Claude Code was not found. Install it and sign in first." -ForegroundColor Yellow
    exit 1
}

Write-Host "Checking Python libraries..."
python -m pip install --quiet --user -r (Join-Path $here 'requirements.txt')

$port = if ($env:SLIDIFY_PORT) { $env:SLIDIFY_PORT } else { 8765 }
Start-Process "http://127.0.0.1:$port"
python (Join-Path $here 'server.py')
