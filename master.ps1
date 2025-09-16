# --- CONFIGURACION DE RUTAS (Modificar segun sea necesario) ---
# Directorio donde se encuentra el script de recolección
$directorioRecoleccion = "C:\Users\SALV\Desktop\PROACTIVAS\PROACTIVAS\devops-powershell-portable 1.7.4()\devops-powershell"
# Script de conversión
$scriptConversion  = "C:\Users\SALV\Desktop\PROACTIVAS\PROACTIVAS\devops-powershell-portable 1.7.4()\JSONtoExcels.ps1"
# Directorio donde el script de recolección guarda los reportes JSON
$directorioReportesJson = "C:\Users\SALV\Desktop\PROACTIVAS\PROACTIVAS\devops-powershell-portable 1.7.4()\devops-powershell\reportes"
# Directorio donde se guardará el Excel final
$directorioExcelFinal   = "C:\Users\SALV\Desktop\PROACTIVAS\PROACTIVAS\devops-powershell-portable 1.7.4()\devops-powershell\reportes\old"


try {
    # --- PASO 1: Ejecutar la recolección de datos ---
    Write-Host "[PASO 1 de 3] - Ejecutando el script de recoleccion de vSphere (puede requerir tu interaccion)..." -ForegroundColor Yellow
    
    # Guardamos la ubicación actual para poder volver después
    Push-Location
    
    # Nos movemos al directorio del script de recolección para que las rutas relativas funcionen
    Set-Location -Path $directorioRecoleccion
    
    # Ejecutamos el script de recolección
    .\portable.ps1
    
    # Regresamos a la ubicación original
    Pop-Location
    
    Write-Host "Recoleccion finalizada." -ForegroundColor Green
    Write-Host ""

    # --- PASO 2: Encontrar el reporte JSON recién creado ---
    Write-Host "[PASO 2 de 3] - Buscando el reporte JSON mas reciente..." -ForegroundColor Yellow
    
    $jsonMasReciente = Get-ChildItem -Path $directorioReportesJson -Filter *.json | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if (-not $jsonMasReciente) {
        throw "No se encontró ningún archivo JSON en el directorio '$directorioReportesJson'. El script no puede continuar."
    }

    $rutaJsonEntrada = $jsonMasReciente.FullName
    Write-Host "Archivo encontrado: '$($jsonMasReciente.Name)'" -ForegroundColor Green
    Write-Host ""

    # --- PASO 3: Convertir el JSON a Excel ---
    Write-Host "[PASO 3 de 3] - Ejecutando la conversion a Excel..." -ForegroundColor Yellow

    $nombreExcel = "$($jsonMasReciente.BaseName).xlsx"
    $rutaExcelFinalCompleta = Join-Path -Path $directorioExcelFinal -ChildPath $nombreExcel

    & $scriptConversion -RutaJsonEntrada $rutaJsonEntrada -RutaExcelSalida $rutaExcelFinalCompleta
    
    Write-Host "Conversion finalizada." -ForegroundColor Green
    Write-Host ""
    Write-Host "Proceso completado exitosamente." -ForegroundColor Cyan
    Write-Host "Excel guardado en: '$rutaExcelFinalCompleta'"
}
catch {
    Write-Error "Ocurrio un error durante la ejecucion: $($_.Exception.Message)"
}

Read-Host "Presiona Enter para salir..."