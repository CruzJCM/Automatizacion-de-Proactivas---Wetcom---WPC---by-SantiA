$baseDir = $PSScriptRoot #portable

$directorioReportesJson = Join-Path -Path $baseDir -ChildPath "devops-powershell\reportes" 

$directorioExcelFinal   = Join-Path -Path $baseDir -ChildPath "devops-powershell\reportes\proactiva-excel"

try {
    Import-Module ImportExcel -ErrorAction Stop
}
catch {
    Write-Error "El módulo 'ImportExcel' no está instalado en PowerShell portable. Por favor descargue el modulo y peguelo en el directorio \devops-powershell-portable 1.7.4()\Modules."
    Read-Host "Presiona Enter para salir."
    return
}

Write-Host "Iniciando la conversión de JSON a Excel..." -ForegroundColor Green

try {
    $jsonMasReciente = Get-ChildItem -Path $directorioReportesJson -Filter *.json | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if (-not $jsonMasReciente) {
        throw "No se encontró ningún archivo JSON en el directorio '$directorioReportesJson'."
    }

    $rutaJsonEntrada = $jsonMasReciente.FullName
    Write-Host "Archivo a convertir encontrado: '$($jsonMasReciente.Name)'"
    
    $nombreExcel = "$($jsonMasReciente.BaseName).xlsx"
    $rutaExcelFinalCompleta = Join-Path -Path $directorioExcelFinal -ChildPath $nombreExcel

    $jsonData = Get-Content -Path $rutaJsonEntrada -Raw | ConvertFrom-Json
    $reportObject = $jsonData.Report
    
    if (Test-Path $rutaExcelFinalCompleta) {
        Remove-Item $rutaExcelFinalCompleta
    }

    Write-Host "Exportando datos a: $rutaExcelFinalCompleta"
    foreach ($sheet in $reportObject.PSObject.Properties) {
        if ($sheet.Value -and $sheet.Value.Count -gt 0) {
            Write-Host "  -> Creando hoja: '$($sheet.Name)'..."
            
            $exportParams = @{
                Path          = $rutaExcelFinalCompleta
                WorksheetName = $sheet.Name
                AutoSize      = $true
                FreezeTopRow  = $true
                AutoFilter    = $true
            }

            if ($sheet.Name -eq 'ESXi') {
                $exportParams['NoNumberConversion'] = @('NtpServer', 'DnsServer')
            }
            if ($sheet.Name -eq 'vCenter') {
                $exportParams['NoNumberConversion'] = @('Version')
            }
            $sheet.Value | Export-Excel @exportParams
        }
    }

    Write-Host "Aplicando estilos al archivo Excel..." -ForegroundColor Yellow
    
    $excelPackage = Open-ExcelPackage -Path $rutaExcelFinalCompleta
    
    foreach ($ws in $excelPackage.Workbook.Worksheets) {
        $headerRange = $ws.Cells[1, 1, 1, $ws.Dimension.End.Column]
        
        $headerRange.Style.Font.Bold = $true
        $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
        $headerRange.Style.Fill.PatternType = 'Solid'
        $colorVerde = [System.Drawing.Color]::FromArgb(0, 176, 80) # RGB para el verde 
        $headerRange.Style.Fill.BackgroundColor.SetColor($colorVerde)
    }
    
    Close-ExcelPackage $excelPackage

    Write-Host "¡Listo! El archivo Excel ha sido generado y estilizado exitosamente ." -ForegroundColor Green
}
catch {
    Write-Error "Ha ocurrido un error critico: $($_.Exception.Message)"
}