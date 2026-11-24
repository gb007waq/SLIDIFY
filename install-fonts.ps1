# Installs the fonts in skills\slidify\fonts for the current Windows user,
# so PowerPoint and Word can show them (no admin rights needed).
# Run:  powershell -ExecutionPolicy Bypass -File .\install-fonts.ps1

$ErrorActionPreference = 'Stop'
$src = Join-Path $PSScriptRoot 'skills\slidify\fonts'
$dest = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
$reg = 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
New-Item -ItemType Directory -Force $dest | Out-Null
if (-not (Test-Path $reg)) { New-Item -Path $reg -Force | Out-Null }

foreach ($f in Get-ChildItem $src -Include *.ttf, *.otf -Recurse) {
    $target = Join-Path $dest $f.Name
    $same = (Test-Path $target) -and ((Get-FileHash $target).Hash -eq (Get-FileHash $f.FullName).Hash)
    if (-not $same) { Copy-Item $f.FullName $target -Force }   # already-installed fonts are in use; skip them
    $name = $f.BaseName -replace '-', ' '   # e.g. Mulish-Regular -> Mulish Regular
    $kind = if ($f.Extension -eq '.otf') { 'OpenType' } else { 'TrueType' }
    Set-ItemProperty -Path $reg -Name "$name ($kind)" -Value $target
    Write-Host "Installed $name"
}
Write-Host 'Done. Restart PowerPoint or Word if they are open.'
