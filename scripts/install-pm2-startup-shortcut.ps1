param(
  [string]$FileName = "BusbarApi_PM2_Resurrect.cmd"
)

$repoRoot = (Resolve-Path "$PSScriptRoot\..").Path
$startupDir = [Environment]::GetFolderPath('Startup')
$targetPath = Join-Path $startupDir $FileName

$content = @"
@echo off
cd /d "$repoRoot"
npx pm2 resurrect
"@

Set-Content -Path $targetPath -Value $content -Encoding ASCII
Write-Host "Oppstartfil opprettet: $targetPath"
