# Lo ejecuta la tarea programada "WordGPT-Gateway". Se posiciona en la raíz del proyecto
# (relativo a la carpeta donde vive este script) y arranca el gateway en modo producción.
$ErrorActionPreference = "Stop"
$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..")
Set-Location $projectRoot
pnpm start
