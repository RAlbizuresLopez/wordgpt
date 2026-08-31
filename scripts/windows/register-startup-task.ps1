# Registra una tarea programada que arranca el gateway de Word GPT al iniciar sesión de
# Windows, sin ventana visible, y que se reintenta sola si el proceso se cae.
# Uso: powershell -ExecutionPolicy Bypass -File register-startup-task.ps1
$ErrorActionPreference = "Stop"

$taskName = "WordGPT-Gateway"
$scriptPath = Join-Path $PSScriptRoot "run-server.ps1"

$action = New-ScheduledTaskAction -Execute "powershell.exe" `
  -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet `
  -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
  -ExecutionTimeLimit ([TimeSpan]::Zero) `
  -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
  -Settings $settings -Description "Gateway local de Word GPT (Codex)" -Force | Out-Null

Write-Host "Tarea '$taskName' registrada. Se iniciará la próxima vez que inicies sesión."
Write-Host "Para arrancarla ahora mismo sin reiniciar sesión: Start-ScheduledTask -TaskName '$taskName'"
