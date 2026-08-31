# Quita la tarea programada creada por register-startup-task.ps1.
# Uso: powershell -ExecutionPolicy Bypass -File uninstall-startup-task.ps1
$ErrorActionPreference = "Stop"

$taskName = "WordGPT-Gateway"
$existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existing) {
  Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
  Write-Host "Tarea '$taskName' eliminada."
} else {
  Write-Host "No había ninguna tarea '$taskName' registrada."
}
