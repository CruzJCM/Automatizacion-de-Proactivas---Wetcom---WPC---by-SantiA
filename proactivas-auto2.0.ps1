if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Install-Module -Name ImportExcel -Force -AllowClobber
}

$baseDir = $PSScriptRoot #portable

if (-not $path){
    $path = (Get-Item -Path ".\" -Verbose).FullName
}

$nombreCliente = Read-Host "Escriba el Nombre del cliente"
$mes = Read-Host "Escriba el mes correspondiente a la tarea Proactiva (ej: Enero, Febrero...)"

$rutaArchivos = Join-Path -Path $baseDir -ChildPath "devops-powershell\reportes\proactiva-excel"    
$archivoSalida = "${baseDir}\devops-powershell\reportes\Anexo\Anexo Tecnico - ${nombreCliente} - ${mes}.xlsx"

$excelMasReciente = Get-ChildItem -Path $rutaArchivos -Filter *.xlsx | Sort-Object LastWriteTime -Descending | Select-Object -First 1


$datosSalida = @()


function Exportar-InformeConEstilo {
    param(
        [Parameter(Mandatory=$true)]
        [object[]]$Datos, 

        [Parameter(Mandatory=$true)]
        [string]$RutaArchivo, 

        [Parameter(Mandatory=$true)]
        [string]$NombreHoja, 

        [string[]]$ColumnasSinConversion 
    )
    
    $exportParams = @{
        Path = $RutaArchivo
        WorksheetName = $NombreHoja
        AutoSize = $true
        FreezeTopRow = $true
        PassThru = $true 
    }

    if ($PSBoundParameters.ContainsKey('ColumnasSinConversion')) {
        $exportParams['NoNumberConversion'] = $ColumnasSinConversion
    }
    
    $excelPackage = $Datos | Export-Excel @exportParams

    $ws = $excelPackage.Workbook.Worksheets[$NombreHoja]
    
    $headerRange = $ws.Cells[1, 1, 1, $ws.Dimension.End.Column]
    
    $headerRange.Style.Font.Bold = $true
    $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $headerRange.Style.Fill.PatternType = 'Solid'
    
    $colorVerde = [System.Drawing.Color]::FromArgb(0, 176, 80)
    $headerRange.Style.Fill.BackgroundColor.SetColor($colorVerde)

    $excelPackage.Save()
    $excelPackage.Dispose()
}


#LA SIGUIENTE FUNCION TIENE UN COMPORTAMIENTO INCORRECTO, LA MEMORIA RECOMENDA VARIA SEGUN LA VERSION DE VCENTER 
#DADO QUE EL DEVOPS PORTABLE NO TRAE INFORMACION DEL LA VERSION DEL VCENTER ESTO NO ES POSIBLE DE REALIZAR
#LAS UNICAS OPCIONES SON, MODIFICAR EL SCRIPT ORIGNAL DE CLOUD&AUT Y AGREGAR UNA HOJA QUE RECOPILE LAS VERSIONES DEL LOS VCENTERS AGREGADOS POR PARAMETRO (PEDIR PERMISO)
#O SE PUEDE TRAER ESA INFORMACIÓN DESDE LAS RVTOOLS, SIENDO ESTA OPCION LA MAS FACIL PERO TAMBIEN AQUELLA QUE ROMPE EL ESQUEMA DE AUTOMATIZACION TOTAL EN ESTOS CHEQUEOS DE PROACTIVA
function AnalizarSize {
    $sizingWeights = @{
        'tiny'    = 1; 'small'   = 2; 'medium'  = 3
        'large'   = 4; 'x-large' = 5
    }
    $recomendacionesSizing = @{
        'tiny'    = @{ Cores = 2;  MemoriaGB = 12 }; 'small'   = @{ Cores = 4;  MemoriaGB = 19 }
        'medium'  = @{ Cores = 8;  MemoriaGB = 28 }; 'large'   = @{ Cores = 16; MemoriaGB = 37 }
        'x-large' = @{ Cores = 24; MemoriaGB = 56 }
    }
    $limitesSizing = @{
        'tiny'    = @{ VMs = 100 };  'small'   = @{ VMs = 1000 }
        'medium'  = @{ VMs = 4000 }; 'large'   = @{ VMs = 10000 }
        'x-large' = @{ VMs = 35000 }
    }

    $mapaDeEstadosVM = @{}
    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $datosVM = Import-Excel -Path $_.FullName -WorksheetName "VM"
        foreach ($vm in $datosVM) { if (-not $mapaDeEstadosVM.ContainsKey($vm.VM)) { $mapaDeEstadosVM.Add($vm.VM, $vm.State) } }
    }

    $datosSizing = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        Import-Excel -Path $_.FullName -WorksheetName "Sizing"
    }

    $datosCombinados = @()
    foreach ($fila in $datosSizing) {
        $estado = $mapaDeEstadosVM[$fila.VM]
        if ($estado -ne "PoweredOff") {
            $datosCombinados += $fila | Add-Member -MemberType NoteProperty -Name "State" -Value $estado -PassThru
        }
    }

    $vcentersUnicos = $datosCombinados | Group-Object -Property @{ Expression = { $_.vCenter + '|' + $_.VM } } | ForEach-Object {
        $_.Group | Sort-Object -Property @{
            Expression = {
                $sizing = ""
                if (-not [string]::IsNullOrEmpty($_."Sizing actual")) { $sizing = $_."Sizing actual".ToLower().Trim() }
                if ($sizingWeights.ContainsKey($sizing)) { return $sizingWeights[$sizing] } else { return 0 }
            }
        } | Select-Object -Last 1
    }

    $informeFinal = @()
    foreach ($vcenter in $vcentersUnicos) {
        $coresActuales = 0; $memoriaActual = 0.0; $vmsActuales = 0
        
        $sizingRecomendado = ""; if (-not [string]::IsNullOrEmpty($vcenter."Sizing recomendado")) { $sizingRecomendado = $vcenter."Sizing recomendado".ToLower().Trim() }
        $sizingActual = ""; if (-not [string]::IsNullOrEmpty($vcenter."Sizing actual")) { $sizingActual = $vcenter."Sizing actual".ToLower().Trim() }

        $estaInfraDimensionado = $false
        $estaSobrecargado = $false

        if ($recomendacionesSizing.ContainsKey($sizingRecomendado)) {
            $rec = $recomendacionesSizing[$sizingRecomendado]
            [int]::TryParse($vcenter.'vCPU', [ref]$coresActuales)
            [double]::TryParse(($vcenter.'Memory GB'.ToString()).Replace(',','.'), [ref]$memoriaActual)

            if (($coresActuales -lt $rec.Cores) -or ($memoriaActual -lt $rec.MemoriaGB)) {
                $estaInfraDimensionado = $true
            }
        }
        
        if ($limitesSizing.ContainsKey($sizingActual)) {
            $limites = $limitesSizing[$sizingActual]
            [int]::TryParse($vcenter.'Cantidad de VMs', [ref]$vmsActuales)

            if ($vmsActuales -gt $limites.VMs) {
                $estaSobrecargado = $true
            }
        }
        
        if ($estaInfraDimensionado -or $estaSobrecargado) {
            $informeFinal += $vcenter
        }
    }

    if ($informeFinal) {
    
        $datosParaExportar = $informeFinal | Select-Object -Property @(
            'State', 'VM', 'vCenter',
            @{ Name = 'VMs';          Expression = { $_.'Cantidad de VMs' } },
            'Sizing actual', 'Sizing recomendado',
            @{ Name = 'Cores';        Expression = { $_.'vCPU' } },
            @{ Name = 'Memoria (GB)'; Expression = { $_.'Memory GB' } }
        )

        Exportar-InformeConEstilo -Datos $datosParaExportar `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Sizing Incorrecto" `
                                  -ColumnasSinConversion "VM", "vCenter"
    } else {
        Write-Host "Análisis finalizado. Todos los vCenters están correctamente dimensionados." -ForegroundColor Yellow
    }
}


function Particiones { 
    $datosSalida = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "Partitions"
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['Disk'] -and $fila.PSObject.Properties['Free %']) {
                $valorNormalizado = ([string]$fila."Free %").Replace(',', '.')
                $valorLimpio = $valorNormalizado -replace '[^0-9.]'
                $freePercentNumeric = 0
                [void][double]::TryParse($valorLimpio, [ref]$freePercentNumeric)
                $diskPath = $fila.Disk.Trim()
                $systemDisks = @("/storage/core", "/storage/archive")
                if (($freePercentNumeric -lt 30) -and ($diskPath -notin $systemDisks)) {
                    [PSCustomObject]@{
                        "VM"         = $fila."VM"
                        "Annotation" = $fila."Annotation"
                        "Disk"       = $fila."Disk"
                        "Free %"     = $fila."Free %"
                    }
                }
            }
        }
    }
    if ($datosSalida) {
                Exportar-InformeConEstilo   -Datos $datosSalida `
                                            -RutaArchivo $archivoSalida `
                                            -NombreHoja "Particiones" `
    }
}


function SyslogCheck {
    $datosSalida = @()
    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
        
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['SyslogGlobalLogDir'] -and $fila.PSObject.Properties['SyslogGlobalLogHost']) {

                $logDir = $fila.SyslogGlobalLogDir
                $logHost = $fila.SyslogGlobalLogHost

                $dirEsIncorrecto = ($logDir -match "/scratch/log" -or $logDir -match "local" -or [string]::IsNullOrEmpty($logDir))
                $hostEsIncorrecto = ($logHost -match "/scratch/log" -or $logHost -match "local" -or [string]::IsNullOrEmpty($logHost))

                if ($dirEsIncorrecto -and $hostEsIncorrecto) {
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"             = $fila."vCenter"
                        "Hostname"            = $fila."Hostname"
                        "Datacenter"          = $fila."Datacenter"
                        "Cluster"             = $fila."Cluster"
                        "SyslogGlobalLogDir"  = $logDir
                        "SyslogGlobalLogHost" = $logHost
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    if ($datosSalida) {
        Exportar-InformeConEstilo -Datos $datosSalida `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Syslog" `
                                  -ColumnasSinConversion "Hostname"

    }
}


function Multipath {
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "Datastores"
    
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['Datastore'] -and $fila.PSObject.Properties['Policy']) {

                $datastoreName = $fila.Datastore
                $policy = $fila.Policy.Trim()

                $esLocal = ($datastoreName -match "local" -or $datastoreName -match "datastore1")
                $esCompartido = -not($esLocal)
                
                $policyRecomendada = ""
                $esInconsistente = $false

                if ($esCompartido -and $policy -ne "RoundRobin") {
                    $esInconsistente = $true
                    $policyRecomendada = "RoundRobin"
                }

                if ($esLocal -and $policy -eq "RoundRobin") {
                    $esInconsistente = $true
                    $policyRecomendada = "MRU o Fixed"
                }

                if ($esInconsistente) {
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"             = $fila."vCenter"
                        "Hostname"            = $fila."Hostname"
                        "Datastore"           = $datastoreName
                        "Policy"              = $policy
                        "Policy recomendado" = $policyRecomendada
                        "Documento"           = ""
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    if ($datosSalida) {
                Exportar-InformeConEstilo -Datos $datosSalida `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "Multipath" `
                                  -ColumnasSinConversion "Hostname"
    }
}


function ConsVer { 
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
        
        $gruposCluster = @{}

        foreach ($fila in $vPartition) {
            if ($null -ne $fila -and -not [string]::IsNullOrEmpty($fila.Cluster) -and -not [string]::IsNullOrEmpty($fila.vCenter)) {
                $claveUnica = "$($fila.vCenter.Trim())|$($fila.Cluster.Trim())"
                
                if (-not $gruposCluster.ContainsKey($claveUnica)) {
                    $gruposCluster[$claveUnica] = @()
                }
                $gruposCluster[$claveUnica] += $fila
            }
        }
        
        foreach ($clave in $gruposCluster.Keys) {
            $hostsDelCluster = $gruposCluster[$clave]
            
            $versionesUnicas = $hostsDelCluster.EsxiVersion | ForEach-Object { $_.Trim() } | Select-Object -Unique
            
            if ($versionesUnicas.Count -gt 1) {
                $datosSalida += $hostsDelCluster
            }
        }
    }
    
    if ($datosSalida) {
        $informeFinal = $datosSalida | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"     = $_."vCenter"
                "Hostname"    = $_."Hostname"
                "Datacenter"  = $_."Datacenter"
                "Cluster"     = $_."Cluster"
                "EsxiVersion" = $_."EsxiVersion"
            }
        }
        Exportar-InformeConEstilo -Datos $informeFinal `
                                  -RutaArchivo $archivoSalida `
                                  -NombreHoja "ConsVer" `
                                  -ColumnasSinConversion "Hostname"
    }
}

function ConsRec {
    param(
        [double]$toleranciaMemoriaGB = 1,
        [double]$toleranciaVelocidadCPU = 1 
    ) 
        $datosSalida = @() 

        Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
            $archivoEntrada = $_.FullName
            $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
            $gruposCluster = @{}

            foreach ($fila in $vPartition) {
                if ($null -ne $fila -and -not [string]::IsNullOrEmpty($fila.Cluster) -and -not [string]::IsNullOrEmpty($fila.vCenter)) {
                    $claveUnica = "$($fila.vCenter.Trim())|$($fila.Cluster.Trim())"
                    if (-not $gruposCluster.ContainsKey($claveUnica)) {
                        $gruposCluster[$claveUnica] = @()
                    } 
                    $gruposCluster[$claveUnica] += $fila 
                } 
            }

            foreach ($clave in $gruposCluster.Keys) { 
            $hostsDelCluster = $gruposCluster[$clave]
                $memorias = @()
                $velocidades = @()
                $modelos = @()
                foreach ($h in $hostsDelCluster) {
                    $memStr = ([string]$h.MemoryGB) -replace '[^0-9\.,]', '' -replace ',', '.'
                    $spdStr = ([string]$h.CpuSpeed) -replace '[^0-9\.,]', '' -replace ',', '.'
                    $memorias += [double]$memStr 
                    $velocidades += [double]$spdStr
                    $modelos += $h.CpuModel.ToString().Trim()
                }
 
                $memMin = ($memorias | Measure-Object -Minimum).Minimum
                $memMax = ($memorias | Measure-Object -Maximum).Maximum
                $spdMin = ($velocidades | Measure-Object -Minimum).Minimum
                $spdMax = ($velocidades | Measure-Object -Maximum).Maximum
                $modelosUnicos = $modelos | Select-Object -Unique

                $inconsistente = $false
                if ($modelosUnicos.Count -gt 1)                                 { $inconsistente = $true }
                if (($memMax - $memMin) -gt $toleranciaMemoriaGB)               { $inconsistente = $true }
                if (($spdMax - $spdMin) -gt $toleranciaVelocidadCPU)            { $inconsistente = $true }

                if ($inconsistente) {
                     $datosSalida += $hostsDelCluster
                }
            }
        }
        if ($datosSalida) {
            $informeFinal = $datosSalida | ForEach-Object {
                [PSCustomObject]@{
                    "vCenter"   = $_."vCenter"
                    "Hostname"  = $_."Hostname"
                    "Cluster"   = $_."Cluster"
                    "MemoryGB"  = $_."MemoryGB"
                    "CpuModel"  = $_."CpuModel"
                    "CpuSpeed"  = $_."CpuSpeed"
                }
            }
            Exportar-InformeConEstilo   -Datos $informeFinal `
                                        -RutaArchivo $archivoSalida `
                                        -NombreHoja "ConsRec" `
                                        -ColumnasSinConversion "Hostname"
        }
}


function PlacaRed { 
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
    
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "VM"
    
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['SO (vCenter)']) {
                $soActual = $fila."SO (vCenter)".Trim()

                if (($soActual -like "Microsoft Windows*") -and ($soActual -notlike "*2003*") -and ($soActual -notlike "*2000*")) {
                    
                    $encontroE1000 = $false 
                    $todosVacios = $true    

                    for ($i = 1; $i -le 10; $i++) {
                        $nombreAdapter = "Adapter_{0:D2}" -f $i 
                        
                        if ($fila.PSObject.Properties[$nombreAdapter]) {
                            $valorAdapter = ([string]$fila.$nombreAdapter).Trim()
                            
                            if ($valorAdapter -eq "e1000" -or $valorAdapter -eq "e1000e" -or $valorAdapter -eq "Flexible") {
                                $encontroE1000 = $true
                            }
                            
                            if (-not [string]::IsNullOrEmpty($valorAdapter)) {
                                $todosVacios = $false
                            }
                        }
                    }
                    
                    if ($encontroE1000 -or $todosVacios) {
                        $objetoPersonalizado = [PSCustomObject]@{
                            "vCenter"      = $fila."vCenter"
                            "VM"           = $fila."VM"
                            "Cluster"      = $fila."Cluster"
                            "Host"         = $fila."Host"
                            "State"        = $fila."State"
                            "ToolsStatus"  = $fila."ToolsStatus"
                            "SO (vCenter)" = $fila."SO (vCenter)"
                            "SO (Tools)"   = $fila."SO (Tools)"
                            "Adapter_01"   = $fila."Adapter_01"
                            "Adapter_02"   = $fila."Adapter_02"
                            "Adapter_03"   = $fila."Adapter_03"
                            "Adapter_04"   = $fila."Adapter_04"
                            "Adapter_05"   = $fila."Adapter_05"
                            "Adapter_06"   = $fila."Adapter_06"
                            "Adapter_07"   = $fila."Adapter_07"
                            "Adapter_08"   = $fila."Adapter_08"
                            "Adapter_09"   = $fila."Adapter_09"
                            "Adapter_10"   = $fila."Adapter_10"
                        }
                        $datosSalida += $objetoPersonalizado
                    }
                }
            }
        }
    }
   
    if ($datosSalida) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "PlacaRed" `
                                    -ColumnasSinConversion "Host"
    }
}

function TSMCheck { 
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
    
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
    
        foreach ($fila in $vPartition) {
            
            if ($fila.PSObject.Properties['ESXIShellTimeOut'] -and $fila.PSObject.Properties['ESXIShellinteractiveTimeOut']) {

                if ( ($fila.ESXIShellTimeOut -lt 300 -or $fila.ESXIShellTimeOut -gt 1800) -or `
                     ($fila.ESXIShellinteractiveTimeOut -lt 300 -or $fila.ESXIShellinteractiveTimeOut -gt 1800) ) {
                
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"                     = $fila."vCenter"
                        "Hostname"                    = $fila."Hostname"
                        "Datacenter"                  = $fila."Datacenter"
                        "Cluster"                     = $fila."Cluster"
                        "ESXIShellTimeOut"            = $fila."ESXIShellTimeOut"
                        "ESXIShellinteractiveTimeOut" = $fila."ESXIShellinteractiveTimeOut"
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida) {
                Exportar-InformeConEstilo   -Datos $datosSalida `
                                            -RutaArchivo $archivoSalida `
                                            -NombreHoja "TSM" `
                                            -ColumnasSinConversion "Hostname"
    }
}

function pManagement { 
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
    
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"

        foreach ($fila in $vPartition) {
            
            if ($null -ne $fila -and $fila.PSObject.Properties['PowerManagement']) {

                if ($fila.PowerManagement.Trim() -ne "High performance") {
                
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"         = $fila."vCenter"
                        "Hostname"        = $fila."Hostname"
                        "Datacenter"      = $fila."Datacenter"
                        "Cluster"         = $fila."Cluster"
                        "PowerManagement" = $fila."PowerManagement"
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida) {
                Exportar-InformeConEstilo   -Datos $datosSalida `
                                            -RutaArchivo $archivoSalida `
                                            -NombreHoja "pManagement" `
                                            -ColumnasSinConversion "Hostname"
    }
}


function vmtools { 
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
    
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "VM"
    
        foreach ($fila in $vPartition) {

            if ($fila.PSObject.Properties['ToolsStatus'] -and $fila.PSObject.Properties['State']) {

                $status = $fila.ToolsStatus.Trim()
                $state = $fila.State.Trim()

                if (-not ( ($status -eq "toolsOk") -or (($state -eq "PoweredOff") -and ($status -eq "toolsNotRunning")) )) {
                
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"              = $fila."vCenter"
                        "VM"                   = $fila."VM"
                        "Cluster"              = $fila."Cluster"
                        "Host"                 = $fila."Host"
                        "ConnectionState"      = $fila."ConnectionState"
                        "State"                = $fila."State"
                        "ToolsStatus"          = $fila."ToolsStatus"
                        "ToolsVersion"         = $fila."ToolsVersion"
                        "ToolsRequiredVersion" = $fila."ToolsRequiredVersion"
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida) {
                Exportar-InformeConEstilo   -Datos $datosSalida `
                                            -RutaArchivo $archivoSalida `
                                            -NombreHoja "vmtools" `
                                            -ColumnasSinConversion "Host"
    }
}


function isos { 
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
    
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "VM"
    
        foreach ($fila in $vPartition) {
            
            if ($fila.PSObject.Properties['IsoConnected']) {
                
                $isoValue = $fila.IsoConnected

                if (-not [string]::IsNullOrEmpty($isoValue) -and $isoValue.Trim() -ne "[]") {
                
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"      = $fila."vCenter"
                        "VM"           = $fila."VM"
                        "Cluster"      = $fila."Cluster"
                        "Host"         = $fila."Host"
                        "IsoConnected" = $fila."IsoConnected"
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida) {
                Exportar-InformeConEstilo   -Datos $datosSalida `
                                            -RutaArchivo $archivoSalida `
                                            -NombreHoja "isos" `
                                            -ColumnasSinConversion "Host"
    }
}

function placasDeRed { 
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
    
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "vNetwork"
    
        foreach ($fila in $vPartition) {

            if ($fila.PSObject.Properties['Status'] -and $fila.PSObject.Properties['Connected'] -and $fila.PSObject.Properties['StartsConnected']) {

                if (([string]$fila.Status -eq "1") -and ($fila.Connected -eq "True") -and ($fila.StartsConnected -eq "False")) {
                
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"         = $fila."vCenter"
                        "VM"              = $fila."VM"
                        "Cluster"         = $fila."Cluster"
                        "Host"            = $fila."Host"
                        "Status"          = $fila."Status"
                        "Mac"             = $fila."Mac"
                        "Connected"       = $fila."Connected"
                        "StartsConnected" = $fila."StartsConnected"
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida) {
                Exportar-InformeConEstilo   -Datos $datosSalida `
                                            -RutaArchivo $archivoSalida `
                                            -NombreHoja "placasDeRed" `
                                            -ColumnasSinConversion "Host"
    }
}


function snapshotsCheck {
    $datosSalida = @()
    $totalSizeMB = 0

    $fechaLimite = (Get-Date).AddDays(-14)

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
    
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "Snapshot"
    
        foreach ($fila in $vPartition) {
            if ($fila.PSObject.Properties['Fecha'] -and $fila.PSObject.Properties['SizeMB']) {
                try {
                    $fechaSnapshot = [datetime]$fila.Fecha
                    
                    if ($fechaSnapshot -lt $fechaLimite) {
                        $objetoPersonalizado = [PSCustomObject]@{
                            "vCenter"  = $fila."vCenter"
                            "VM"       = $fila."VM"
                            "Snapshot" = $fila."Snapshot"
                            "Fecha"    = $fila.Fecha
                            "SizeMB"   = $fila."SizeMB"
                        }
                        $datosSalida += $objetoPersonalizado
                        $totalSizeMB += [double]$fila.SizeMB
                    }
                } catch {
                    Write-Warning "No se pudo convertir la fecha '$($fila.Fecha)' para la VM '$($fila.VM)'. Se omite esta fila."
                }
            }
        }
    }

    if ($datosSalida.Count -gt 0) {
        $snapshotMasGrande = $datosSalida | Sort-Object -Property @{ Expression = { [double]$_.SizeMB } } | Select-Object -Last 1
        $snapshotMasGrande.SizeMB = $snapshotMasGrande.SizeMB / 1024
        $filaResumenGrande = [PSCustomObject]@{
            vCenter  = ""
            VM       = "SNAPSHOT MAS GRANDE (GB)"
            Snapshot = $snapshotMasGrande.Snapshot
            Fecha    = ""
            SizeMB   = [math]::Round($snapshotMasGrande.SizeMB,2)
        }
        
        $totalSizeGB = $totalSizeMB / 1024
        $filaResumenTotal = [PSCustomObject]@{
            vCenter  = ""
            VM       = "TOTAL (GB)"
            Snapshot = ""
            Fecha    = ""
            SizeMB   = [math]::Round($totalSizeGB, 2) 
        }

        $datosSalida += $filaResumenGrande
        $datosSalida += $filaResumenTotal
    }

    if ($datosSalida) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "snapshots"
    }
}

function endOfSupport { #done
    $datosSalida = @()
    $eosDates = @{
        '8.0' = '2027-10-11'
        '7.0' = '2025-10-02'
        '6.7' = '2022-10-15'
        '6.5' = '2022-10-15'
    }

    $fechaLimite = (Get-Date).AddYears(1)

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
    
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
    
        foreach ($fila in $vPartition) {

            if ($fila.PSObject.Properties['ESXiVersion']) {
                $versionCompleta = $fila.ESXiVersion
                $versionPrincipal = $null # Se reinicia la variable

                if ($versionCompleta -match '(\d+\.\d+)') {
                    $versionPrincipal = $matches[1]
                }

                if ($versionPrincipal -and $eosDates.ContainsKey($versionPrincipal)) {
                    $fechaEosString = $eosDates[$versionPrincipal]
                    $fechaEos = [datetime]$fechaEosString

                    if ($fechaEos -lt $fechaLimite) {
                        $objetoPersonalizado = [PSCustomObject]@{
                            "vCenter"      = $fila."vCenter"
                            "Hostname"     = $fila."Hostname"
                            "Datacenter"   = $fila."Datacenter"
                            "Cluster"      = $fila."Cluster"
                            "Version"      = $versionCompleta
                            "Fecha de EoS" = $fechaEos.ToString('yyyy-MM-dd') # Se formatea la fecha
                        }
                        $datosSalida += $objetoPersonalizado
                    }
                }
            }
        }
    }
    
    if ($datosSalida) {
                Exportar-InformeConEstilo   -Datos $datosSalida `
                                            -RutaArchivo $archivoSalida `
                                            -NombreHoja "endOfSupport" `
                                            -ColumnasSinConversion "Hostname"}
}

function compatComponentes { #done
    $datosSalida = @()

    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
    
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
    
        foreach ($fila in $vPartition) {

            if ($fila.PSObject.Properties['Supported']) {

                if ($fila.Supported.Trim() -eq "False") {
                
                    $objetoPersonalizado = [PSCustomObject]@{
                        "vCenter"                   = $fila."vCenter"
                        "Hostname"                  = $fila."Hostname"
                        "Datacenter"                = $fila."Datacenter"
                        "Cluster"                   = $fila."Cluster"
                        "Version"                   = $fila."ESXiVersion"
                        "Version minima soportada"  = $fila."Supported Releases"
                        "Supported"                 = $fila."Supported"
                    }
                    $datosSalida += $objetoPersonalizado
                }
            }
        }
    }
    
    if ($datosSalida) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "compatComponentes" `
    }
}

function ntpCheck { #Done
    $datosSalida = @()
    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
        $gruposCluster = @{}
        # Agrupar por vCenter + Cluster
        foreach ($fila in $vPartition) {
            if ($null -ne $fila -and -not [string]::IsNullOrEmpty($fila.Cluster) -and -not [string]::IsNullOrEmpty($fila.vCenter)) {
                $claveUnica = "$($fila.vCenter.Trim())|$($fila.Cluster.Trim())"
                if (-not $gruposCluster.ContainsKey($claveUnica)) {
                    $gruposCluster[$claveUnica] = @()
                }
                $gruposCluster[$claveUnica] += $fila
            }
        }
        foreach ($clave in $gruposCluster.Keys) {
            $hostsDelCluster = $gruposCluster[$clave]
            $ntpsUnicos = $hostsDelCluster | ForEach-Object { $_.NtpServer.Trim() } | Select-Object -Unique
            $hayInconsistencia = ($ntpsUnicos.Count -gt 1)
            $hayNtpdFalse = $hostsDelCluster | Where-Object { $_.NtpdRunning -ne $true }
            if ($hayInconsistencia) {
                $datosSalida += $hostsDelCluster
            } elseif ($hayNtpdFalse) {
                $datosSalida += $hayNtpdFalse
            }
        }
    }
    if ($datosSalida) {
        $informeFinal = $datosSalida | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"      = $_."vCenter"
                "Hostname"     = $_."Hostname"
                "Datacenter"   = $_."Datacenter"
                "Cluster"      = $_."Cluster"
                "NtpServer"    = $_."NtpServer"
                "NtpdRunning"  = $_."NtpdRunning"
            }
        }
                Exportar-InformeConEstilo   -Datos $informeFinal `
                                            -RutaArchivo $archivoSalida `
                                            -NombreHoja "ntpCheck" `
                                            -ColumnasSinConversion "Hostname", "NtpServer"
    }
}


function Drivers { #en algun momento lo resolveremos
    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
    
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi IO"

        $datosSalida += $vPartition | ForEach-Object {
            $obj = [PSCustomObject]@{
                "vCenter" = $_."vCenter"
                "Hostname" = $_."Hostname"
                "Version del Host" = $_."ESXi Release"
                "Placa" = $_."Placa"
                "Controlador" = $_."Controlador"
                "Vendor" = $_."Vendor"
                "Driver" = $_."Driver"
                "Version" = $_."Version"
                "Firmware" = $_."Firmware"
                "Vid" = $_."Vid"
                "Did" = $_."Did"
                "Svid" = $_."Svid"
                "ssid" = $_."ssid"
                "URL" = $_."URL"
            }
            $obj
            
        }
    }
    
    if ($datosSalida) {
            Exportar-InformeConEstilo   -Datos $datosSalida `
                                        -RutaArchivo $archivoSalida `
                                        -NombreHoja "Drivers" `
                                        -ColumnasSinConversion "Hostname"}
}


function DNSConfig {
    $datosSalida = @()
    Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "ESXi"
        $gruposCluster = @{}
        foreach ($fila in $vPartition) {
            if ($null -ne $fila -and -not [string]::IsNullOrEmpty($fila.Cluster) -and -not [string]::IsNullOrEmpty($fila.vCenter)) {
                $claveUnica = "$($fila.vCenter.Trim())|$($fila.Cluster.Trim())"
                if (-not $gruposCluster.ContainsKey($claveUnica)) {
                    $gruposCluster[$claveUnica] = @()
                }
                $gruposCluster[$claveUnica] += $fila
            }
        }
        foreach ($clave in $gruposCluster.Keys) {
            $hostsDelCluster = $gruposCluster[$clave]
            
            $hostsSinDns = $hostsDelCluster.Where({ [string]::IsNullOrEmpty($_.DnsServer) })
            $hayDnsVacio = $hostsSinDns.Count -gt 0

            $dnsConfigurados = $hostsDelCluster.DnsServer.Trim() | Select-Object -Unique
            $hayInconsistencia = $dnsConfigurados.Count -gt 1
            
            if ($hayDnsVacio) {
                Write-Host "  Resultado: Se encontraron $($hostsSinDns.Count) hosts con DNS vac o." -ForegroundColor Yellow
            }
            if ($hayInconsistencia) {
                Write-Host "  Resultado: Se detect  inconsistencia en la configuraci n de DNS." -ForegroundColor Yellow
            }

            if ($hayDnsVacio -or $hayInconsistencia) {
                $datosSalida += $hostsDelCluster
            }
        }
    }
    if ($datosSalida) {
        $informeFinal = $datosSalida | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"         = $_."vCenter"
                "Hostname"        = $_."Hostname"
                "Datacenter"      = $_."Datacenter"
                "Cluster"         = $_."Cluster"
                "ConnectionState" = $_."ConnectionState"
                "DnsServer"       = $_."DnsServer"
            }
        }
        Exportar-InformeConEstilo   -Datos $informeFinal `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "DNS" `
                                    -ColumnasSinConversion "Hostname"
    }
}

function Licencia {
    $datosSalida = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $fechaLimite = Get-Date
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "vLicense"
        $vPartition | ForEach-Object {
            $expiraRaw = $_.ExpirationDate

            if ($expiraRaw -eq "Evaluation") {
                [PSCustomObject]@{
                    Licencia           = $_.Name
                    Componente         = $_.ProductName
                    Datacenter         = $_.vCenter
                    "En Uso"           = $_.Used
                    Limite             = $_.Total
                    "Fecha Expiracion" = $expiraRaw
                }
            }
            else {
                try {
                    $expiraConvertida = [datetimeoffset]::Parse($expiraRaw)
                    if ($expiraConvertida.DateTime -lt $fechaLimite) {
                        [PSCustomObject]@{
                            Licencia           = $_.Name
                            Componente         = $_.ProductName
                            Datacenter         = $_.vCenter
                            "En Uso"           = $_.Used
                            Limite             = $_.Total
                            "Fecha Expiracion" = $expiraConvertida.ToString("yyyy-MM-dd HH:mm")
                        }
                    }
                } catch {
                    Write-Warning "No se pudo analizar la fecha: $expiraRaw"
                }
            }
        }
    }

    if ($datosSalida) {
        Exportar-InformeConEstilo   -Datos $datosSalida `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "Licencia" `
    }
}

function NIOC { #done
    $datosSalida = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        $archivoEntrada = $_.FullName
        $vPartition = Import-Excel -Path $archivoEntrada -WorksheetName "vDS"
        
        $vPartition | ForEach-Object {
            if ($_."NIOC Enabled" -ne $true) {
                [PSCustomObject]@{
                    "Name"         = $_."Name"
                    "MTU"          = $_."MTU"
                    "NIOC Enabled" = $_."NIOC Enabled"
                }
            }
        }
    }
    
    if ($datosSalida) {
                Exportar-InformeConEstilo   -Datos $datosSalida `
                                            -RutaArchivo $archivoSalida `
                                            -NombreHoja "NIOC" `
    }
}


function vss {

    $informeFinal = @()

    $todosLosDatos = Get-ChildItem -Path $excelMasReciente -Filter *.xlsx | ForEach-Object {
        Import-Excel -Path $_.FullName -WorksheetName "Standard Switch"
    } | Where-Object { -not [string]::IsNullOrEmpty($_.Cluster) }

    $gruposCluster = $todosLosDatos | Group-Object -Property @{ Expression = { $_.vCenter + '|' + $_.Cluster } }

    foreach ($cluster in $gruposCluster) {
        
        $hostsUnicos = $cluster.Group.ESXi | Select-Object -Unique
        if ($hostsUnicos.Count -le 1) {
            continue
        }
        
        $gruposPorHost = $cluster.Group | Group-Object -Property ESXi
        
        $huellasDePortGroups = $gruposPorHost | ForEach-Object {
            ($_.Group.PortGroup | Select-Object -Unique | Sort-Object) -join ';'
        } | Select-Object -Unique
        
        $listaDePortGroupsEsInconsistente = $huellasDePortGroups.Count -gt 1

        $configuracionEsInconsistente = $false
        if (-not $listaDePortGroupsEsInconsistente) {
            $portgroupsEnCluster = $cluster.Group | Group-Object -Property PortGroup
            foreach ($pg in $portgroupsEnCluster) {
                $configuraciones = $pg.Group | ForEach-Object { "$($_.Switch)|$($_.vLAN)" } | Select-Object -Unique
                if ($configuraciones.Count -gt 1) {
                    $configuracionEsInconsistente = $true
                    break
                }
            }
        }
        
        if ($listaDePortGroupsEsInconsistente -or $configuracionEsInconsistente) {
            $informeFinal += $cluster.Group
        }
    }

    if ($informeFinal) {
        $informeLimpio = $informeFinal | ForEach-Object {
            [PSCustomObject]@{
                "vCenter"   = $_."vCenter"
                "ESXi"      = $_."ESXi"
                "Cluster"   = $_."Cluster"
                "PortGroup" = $_."PortGroup"
                "Switch"    = $_."Switch"
                "vLAN"      = $_."vLAN"
            }
        } 

        Exportar-InformeConEstilo   -Datos $informeLimpio `
                                    -RutaArchivo $archivoSalida `
                                    -NombreHoja "vss"
    }
}

Write-Host "Seleccione las tareas a ejecutar:"
Write-Host "1. Tareas mensuales"
Write-Host "2. Tareas trimestrales"
Write-Host "3. Tareas semestrales"
Write-Host "4. Salir"

$opcion = Read-Host "Ingrese el numero de la tarea (separe multiples opciones con comas)"

$tareasSeleccionadas = $opcion -split ','

foreach ($tarea in $tareasSeleccionadas) {
    switch ($tarea.Trim()) {
        1 {
            Write-Host "Ejecutando tareas mensuales..."
            AnalizarSize
            Particiones
            SyslogCheck
            Multipath
            ConsVer
            ConsRec
            PlacaRed
            TSMCheck
            pManagement
            vmtools
            isos
            placasDeRed
            snapshotsCheck
            
            break
        }
        2 {
            Write-Host "Ejecutando tareas trimestrales..."
            endOfSupport
            compatComponentes
            ntpCheck
            Drivers
            DNSConfig
            Licencia
            NIOC
            vss
            
            break
        }
        3 {
            Write-Host "Esta tarea se realiza manualmente, leer documentacion: How to Proactivas"
            break
        }
        4 {
            Write-Host "Saliendo del script."
            exit
        }
        default {
            Write-Host "Opcion no valida. Por favor, seleccione una opcion valida."
        }
    }
}

Write-Host "Proceso completado. Se ha creado un nuevo archivo Excel en: $archivoSalida"
