param(
  [string]$TaskName = "BusbarApi_PM2_Resurrect"
)

Write-Host "Fjerner scheduled task '$TaskName'..."
schtasks /Delete /TN $TaskName /F | Out-Host

if ($LASTEXITCODE -ne 0) {
  throw "Kunne ikke fjerne scheduled task '$TaskName'."
}

Write-Host "Task fjernet."
