# Instalación de un solo paso para Word GPT en Windows.
# Antes de correr esto, autentica Codex CLI (requiere abrir el navegador, no se puede
# automatizar): codex login
#
# Uso: powershell -ExecutionPolicy Bypass -File scripts\windows\setup.ps1
$ErrorActionPreference = "Stop"
$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..")
Set-Location $projectRoot

function Test-Command($name) {
  return [bool](Get-Command $name -ErrorAction SilentlyContinue)
}

if (-not (Test-Command "node")) {
  throw "No se encontró Node.js. Instala Node.js 20+ desde https://nodejs.org antes de continuar."
}
if (-not (Test-Command "pnpm")) {
  throw "No se encontró pnpm. Instala Node.js 20+ y corre 'corepack enable', o 'npm install -g pnpm'."
}
if (-not (Test-Command "codex")) {
  throw "No se encontró Codex CLI. Instálalo con 'npm install -g @openai/codex' y luego corre 'codex login' antes de repetir este script."
}

Write-Host "==> Instalando dependencias (pnpm install)..."
pnpm install

Write-Host "==> Instalando y confiando el certificado local (pnpm run certs)..."
Write-Host "    Si el panel muestra un error de certificado más adelante, vuelve a correr esto"
Write-Host "    como Administrador agregando --machine: pnpm exec office-addin-dev-certs install --machine"
pnpm run certs

Write-Host "==> Generando el panel para uso diario (pnpm run build)..."
pnpm run build

Write-Host "==> Registrando el arranque automático al iniciar sesión..."
& (Join-Path $PSScriptRoot "register-startup-task.ps1")

Write-Host "==> Arrancando el gateway ahora mismo para esta sesión..."
Start-ScheduledTask -TaskName "WordGPT-Gateway"
Start-Sleep -Seconds 2

Write-Host "==> Registrando el complemento en Word (pnpm run sideload)..."
try {
  pnpm run sideload
} catch {
  Write-Warning "El sideload automático falló. Carga 'manifest.xml' manualmente desde Word: Insertar -> Complementos -> Mis complementos -> Cargar mis complementos."
}

Write-Host ""
Write-Host "Listo. Recuerda:"
Write-Host " - El gateway ya quedó registrado para iniciar solo la próxima vez que inicies sesión en Windows."
Write-Host " - Si Codex CLI no está autenticado en esta cuenta de Windows, corre 'codex login' una vez."
Write-Host " - La primera vez que abras el complemento en Word, acepta el aviso de confianza si aparece."
