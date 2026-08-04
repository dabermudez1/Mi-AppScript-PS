# Forzar a la consola de Windows a usar UTF-8 para pintar emojis y caracteres especiales
# Comando: Set-ExecutionPolicy Bypass -Scope Process; .\deploy-prod.ps1
# Comando simple: .\deploy-prod.ps1
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host " Iniciando despliegue masivo a Produccion (Formato Pulido)..." -ForegroundColor Cyan

$clinicas = @(
    @{ Name = "Bembrive";   Config = ".clasp.prod.bembrive.json" },
    @{ Name = "Moana";      Config = ".clasp.prod.moana.json" },
    @{ Name = "Navia";      Config = ".clasp.prod.navia.json" },
    @{ Name = "Olimpia";    Config = ".clasp.prod.olimpia.json" },
    @{ Name = "Ponteareas"; Config = ".clasp.prod.ponteareas.json" },
    @{ Name = "Redondela";  Config = ".clasp.prod.redondela.json" }
)

$globalAuth = "$env:USERPROFILE\.clasprc.json"
$globalAuthBak = "$env:USERPROFILE\.clasprc.json.bak"

if (Test-Path ".\.clasp.dev.bak") { Remove-Item ".\.clasp.dev.bak" -Force }
if (Test-Path $globalAuthBak) { Remove-Item $globalAuthBak -Force }

if (Test-Path ".\.clasp.json") { Rename-Item ".\.clasp.json" ".clasp.dev.bak" -Force }
if (Test-Path $globalAuth) { Rename-Item $globalAuth ".clasprc.json.bak" -Force }

if (Test-Path ".\.clasp.auth.prod.json") {
    Copy-Item ".\.clasp.auth.prod.json" $globalAuth -Force
} else {
    Write-Host " ERROR CRITICO: No se encuentra .clasp.auth.prod.json" -ForegroundColor Red
    if (Test-Path ".\.clasp.dev.bak") { Rename-Item ".\.clasp.dev.bak" ".clasp.json" -Force }
    if (Test-Path $globalAuthBak) { Rename-Item $globalAuthBak ".clasprc.json" -Force }
    exit
}

$erroresDetectados = 0

foreach ($clinica in $clinicas) {
    $name = $clinica.Name
    $configFile = $clinica.Config

    Write-Host "`n--------------------------------------------------" -ForegroundColor Gray
    Write-Host " Trabajando en Clinica: [$name]" -ForegroundColor Yellow
    Write-Host "--------------------------------------------------" -ForegroundColor Gray

    if (-not (Test-Path ".\$configFile")) {
        Write-Host " Error: No se encuentra el archivo $configFile" -ForegroundColor Red
        $erroresDetectados++
        continue
    }

    Copy-Item ".\$configFile" ".\.clasp.json" -Force

    Write-Host " Subiendo codigo fuente..." -ForegroundColor Cyan
    $output = clasp push -f 2>&1

    foreach ($line in $output) {
        # Limpieza estetica de los caracteres de arbol manglados por Windows
        $cleanLine = $line -replace "ÔööÔöÇ", "" -replace "Ôöö", ""
        
        if ($cleanLine -match "error" -or $cleanLine -match "permission" -or $cleanLine -match "GaxiosError") {
            Write-Host " Detalle: $cleanLine" -ForegroundColor Red
        } else {
            Write-Host "   $cleanLine" -ForegroundColor Gray
        }
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Host " OCURRIO UN ERROR EN [$name]" -ForegroundColor Red
        $erroresDetectados++
    } else {
        Write-Host " ¡Subida exitosa para [$name]!" -ForegroundColor Green
    }

    if (Test-Path ".\.clasp.json") { Remove-Item ".\.clasp.json" -Force }
}

if (Test-Path $globalAuth) { Remove-Item $globalAuth -Force }
if (Test-Path $globalAuthBak) { Rename-Item $globalAuthBak ".clasprc.json" -Force }
if (Test-Path ".\.clasp.dev.bak") { Rename-Item ".\.clasp.dev.bak" ".clasp.json" -Force }

Write-Host "`n==================================================" -ForegroundColor Gray
if ($erroresDetectados -eq 0) {
    Write-Host " FIN DEL PROCESO: ¡Todo desplegado con exito sin errores!" -ForegroundColor Green
} else {
    Write-Host " FIN DEL PROCESO: Se detectaron $erroresDetectados fallo(s)." -ForegroundColor Red
}
Write-Host "==================================================" -ForegroundColor Gray
