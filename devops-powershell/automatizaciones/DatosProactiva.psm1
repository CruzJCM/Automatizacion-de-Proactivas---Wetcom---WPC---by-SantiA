using module ".\lib\proactivas.psm1"

$credenciales = "cHJvYWN0aXZhOlBhc3N3b3JkMTIzJA=="

function Get-DatosProactivas($vcenters) {
    Start-DatosProactivas($vcenters)

    <#
	.Synopsis
    Recolección de Datos para Proactiva
	.Component
	vcenter
    .Role
    ui
	#>
}

function Start-DatosProactivas($vcenters){
    $proactiva = New-Object Proactiva

    # --- [NUEVO] Se inicializa una lista para guardar los datos de los vCenters ---
    $vCenterData = @()

    foreach($vcenter in $vcenters.conn){
        $proactiva.setCurrentVcenter($vcenter)
        Write-Host "Processing vCenter: $vcenter"
        Write-Host "`tGathering Hosts..." -NoNewLine
        $hosts = Get-VMHost -server $vcenter
        Write-Host "`t`t"$hosts.length"Hosts found"
        Write-Host "`tGathering VMs..." -NoNewLine
        $vms = Get-VM -server $vcenter
        Write-Host "`t`t"$vms.length"VMs found"
        Write-Host "`tGathering VDSwitches..." -NoNewLine
        $vdswitches = Get-VDSwitch -Server $vcenter
        Write-Host "`t`t"$vdswitches.length" Virtual Distributed Switches found"
        Write-Host "`tGathering Clusters..." -NoNewLine
        $clusters = Get-Cluster -Server $vcenter
        Write-Host "`t`t"$clusters.length" Clusters found"

        # --- [NUEVO] Bloque para recopilar la información del vCenter ---
        Write-Host "`tGathering vCenter Appliance Info..."
        # Buscamos la VM cuyo nombre coincida con el del servidor vCenter
        $vcenter_vm = $vms | Where-Object { $_.Name -eq $vcenter.Name }
        
        $vCenterData += [PSCustomObject]@{
            "vCenter Server" = $vcenter.Name
            "VM Name"        = if ($vcenter_vm) { $vcenter_vm.Name } else { "No Encontrada" }
            "Version"        = $vcenter.Version
            "Build"          = $vcenter.Build
        }
        # --- Fin del bloque nuevo ---

        $proactiva.processEsxi($hosts)
        $proactiva.processNic($hosts, $vdswitches)
        $proactiva.processVm($vms, $clusters)
        $proactiva.processDatastore($hosts)
        $proactiva.processSwitch($hosts)
        $proactiva.processKernelAdapters($hosts)
        $proactiva.processSnapshot($vms)
        $proactiva.processPartitions($vms)
        $proactiva.processVcenterSizing($vms, $hosts)
        $proactiva.processvDS($vdswitches)
        $proactiva.processLicense()
        $proactiva.executeAlarm($hosts)
    }
    
    # --- [NUEVO] Se procesa la nueva información recolectada ---
    $proactiva.processVcenter($vCenterData)

    $file = [PSCustomObject] @{
        Result="OK";
        Name="Proactiva";
        Version = $global:APP_VERSION;
        DateTime= (Get-Date -Format "yyyy-MM-dd HH:mm");
        LocalHost= [system.environment]::MachineName;
        User = whoami;
        Endpoint=$vcenters.host;
        Component="vcenter";
        Report = $proactiva.getReport();
        IdAutomatizacion=$credenciales;
    }

    $file | ConvertTo-Json -Depth 99| Set-Content ($global:CONFIG.REPORTS_FOLDER + "/" + (Get-Date).toString('yyyy-MM-dd HHmmss') + "_" + $file.Name + ".json")
}