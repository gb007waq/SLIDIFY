# Starts the SLIDIFY app and opens it in your browser.
# Double-click SLIDIFY.bat in the repo folder, or run:
#   powershell -ExecutionPolicy Bypass -File <repo>\app\start.ps1

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$port = if ($env:SLIDIFY_PORT) { $env:SLIDIFY_PORT } else { 8765 }
$url = "http://127.0.0.1:$port"

# Already running? Just open it.
try {
    Invoke-WebRequest -UseBasicParsing -TimeoutSec 2 $url | Out-Null
    Write-Host "SLIDIFY is already running - opening it."
    Start-Process $url
    exit 0
} catch { }

if (-not (Get-Command claude -ErrorAction SilentlyContinue)) {
    Write-Host "Claude Code was not found. Install it and sign in first." -ForegroundColor Yellow
    exit 1
}
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python was not found. Install Python 3 first." -ForegroundColor Yellow
    exit 1
}

Write-Host "Checking Python libraries..."
python -m pip install --quiet --user --disable-pip-version-check -r (Join-Path $here 'requirements.txt')

# Open the browser once the server answers.
Start-Job -ArgumentList $url {
    param($u)
    for ($i = 0; $i -lt 60; $i++) {
        try { Invoke-WebRequest -UseBasicParsing -TimeoutSec 1 $u | Out-Null; Start-Process $u; break }
        catch { Start-Sleep -Milliseconds 500 }
    }
} | Out-Null

python (Join-Path $here 'server.py')
