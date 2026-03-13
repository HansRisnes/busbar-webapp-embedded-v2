param(
  [string]$TaskName = "BusbarApi_PM2_Resurrect"
)

$repoRoot = (Resolve-Path "$PSScriptRoot\..").Path
$taskCommand = "cmd.exe /c cd /d `"$repoRoot`" && npx pm2 resurrect"

Write-Host "Registerer scheduled task '$TaskName'..."
schtasks /Create /TN $TaskName /SC ONLOGON /TR $taskCommand /RL LIMITED /F | Out-Host

if ($LASTEXITCODE -ne 0) {
  throw "Kunne ikke opprette scheduled task '$TaskName'."
}

Write-Host "Task opprettet. Verifiser med: schtasks /Query /TN $TaskName /V /FO LIST"
