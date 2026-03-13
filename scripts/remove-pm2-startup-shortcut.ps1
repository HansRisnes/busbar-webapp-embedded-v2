param(
  [string]$FileName = "BusbarApi_PM2_Resurrect.cmd"
)

$startupDir = [Environment]::GetFolderPath('Startup')
$targetPath = Join-Path $startupDir $FileName

if (Test-Path $targetPath) {
  Remove-Item -Path $targetPath -Force
  Write-Host "Oppstartfil fjernet: $targetPath"
} else {
  Write-Host "Ingen oppstartfil funnet: $targetPath"
}
