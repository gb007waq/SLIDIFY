# Installs the Slidify skill for Claude Code on this computer.
# Run from the repo folder:  powershell -ExecutionPolicy Bypass -File .\install.ps1

$ErrorActionPreference = 'Stop'
$src = Join-Path $PSScriptRoot 'skills\slidify'
$dest = Join-Path $env:USERPROFILE '.claude\skills\slidify'

New-Item -ItemType Directory -Force (Split-Path $dest) | Out-Null
if (Test-Path $dest) { Remove-Item -Recurse -Force $dest }
Copy-Item -Recurse $src $dest

python -m pip install --quiet --user python-pptx pillow

Write-Host "Slidify installed to $dest"
Write-Host "In Claude Code, ask: 'slidify a presentation on <topic>'"
