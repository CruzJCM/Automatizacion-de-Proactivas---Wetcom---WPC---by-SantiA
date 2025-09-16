. .\automatizaciones\lib\check-hcl.ps1
. .\automatizaciones\lib\create-alarm.ps1
. .\automatizaciones\lib\check-vsan-hcl.ps1

class Proactiva {

    [PSCustomObject[]] $toolsRefJson = (Get-Content -Path (".\automatizaciones\lib\data\vmtools-ref.json") | ConvertFrom-Json)
    [PSCustomObject[]] $vCenterSizing = (Get-Content -Path (".\automatizaciones\lib\data\vcenter-sizing.json") | ConvertFrom-Json)
    [PSCustomObject[]] $vsanhcl = (Get-Content -Path (".\automatizaciones\lib\data\vsan-hcl.json") | ConvertFrom-Json)

    [String[]] $annotations = @("VMware vCenter Server Appliance")
    
    [PSCustomObject] $VMHostNetworkAdapters = [PSCustomObject] @{}
    [PSCustomObject] $VMHostStandardSwitches = [PSCustomObject] @{}
    [PSCustomObject] $toolsReference = [PSCustomObject] @{}

    [PSCustomObject] $ioHclRef = $null
    [String] $currentVCenter = "No vCenter"

    $esxiReport = [System.Collections.ArrayList]@()
    $nicReport = [System.Collections.ArrayList]@()
    $vmReport = [System.Collections.ArrayList]@()
    $datastoreReport = [System.Collections.ArrayList]@()
    $switchReport = [System.Collections.ArrayList]@()
    $kernelAdaptersReport = [System.Collections.ArrayList]@()
    $snapshotReport = [System.Collections.ArrayList]@()
    $partitionReport = [System.Collections.ArrayList]@()
    $sizingReport = [System.Collections.ArrayList]@()
    $vdsReport = [System.Collections.ArrayList]@()
    $vNetworkReport = [System.Collections.ArrayList]@()
    $vLicenseReport = [System.Collections.ArrayList]@()
    $vCenterReport = [System.Collections.ArrayList]@()


    processEsxi($hosts) {
        Write-Host "`tProcessing ESXi..." -NoNewline
        for ($count = 0; $count -lt $hosts.length; $count++) {
            Show-Progress $hosts.length ($count + 1)
            $h = $hosts[$count]
            try {
                $hclResult = $h | Check-HCL
            }
            catch {
                $hclResult = [PSCustomObject]@{
                    Model             = "No data available"
                    Supported         = $false
                    SupportedReleases = "No data available"
                    Reference         = "No data available"
                    Note              = "Host cannot be processed"
                }
            }
            
            
            if (@("Connected", "Maintenance") -contains $h.ConnectionState) {
                $advSett = $h | Get-AdvancedSetting
                $ESXIShellTimeOut = ($advSett | Where-Object { $_.name -eq 'UserVars.ESXIShellTimeOut' }).Value
                $ESXIShellinteractiveTimeOut = ($advSett | Where-Object { $_.name -eq 'UserVars.ESXIShellinteractiveTimeOut' }).Value
                $SyslogGlobalLogDir = ($advSett | Where-Object { $_.name -eq 'Syslog.global.logDir' }).Value
                $SyslogGlobalLogHost = ($advSett | Where-Object { $_.name -eq 'Syslog.global.logHost' }).Value
                $ntpdRunning = ($h | Get-VMHostService | Where-Object { $_.key -eq "ntpd" }).Running.ToString()
                $hostModel = $hclResult.Model
                $supported = if ($hclResult.Supported) { "True" } else { "False" }
                $supportedReleases = $hclResult.SupportedReleases -join ","
                $reference = $hclResult.Reference
                $note = $hclResult.Note
            }
            else {
                $ESXIShellTimeOut = $null
                $ESXIShellinteractiveTimeOut = $null
                $SyslogGlobalLogDir = $null
                $SyslogGlobalLogHost = $null
                $ntpdRunning = $null
                $hostModel = $hclResult.Model
                $supported = if ($hclResult.Supported) { "True" } else { "False" }
                $supportedReleases = $hclResult.SupportedReleases -join ","
                $reference = $hclResult.Reference
                $note = $hclResult.Note
            }

            $this.esxiReport += [PSCustomObject] @{
                vCenter                     = $this.currentVCenter
                Hostname                    = $h.Name
                Model                       = $hostModel
                Datacenter                  = ($h | Get-Datacenter).Name
                Cluster                     = ($h | Get-Cluster).Name
                ESXiVersion                 = $h.ExtensionData.Summary.Config.Product.FullName
                ConnectionState             = $h.ConnectionState.ToString()
                MemoryGB                    = [math]::Round($h.MemoryTotalGB, 3)
                CpuModel                    = $h.ProcessorType
                CpuSpeed                    = $h.ExtensionData.summary.hardware.CpuMhz
                DnsServer                   = $h.ExtensionData.Config.Network.DnsConfig.Address -join ","
                NtpServer                   = ($h | Get-VMHostNtpServer) -join ","
                NtpdRunning                 = $ntpdRunning
                ESXIShellTimeOut            = $ESXIShellTimeOut
                ESXIShellinteractiveTimeOut = $ESXIShellinteractiveTimeOut
                PowerManagement             = $h.ExtensionData.Hardware.CpuPowerManagementInfo.CurrentPolicy
                SyslogGlobalLogDir          = $SyslogGlobalLogDir
                SyslogGlobalLogHost         = $SyslogGlobalLogHost
                Supported                   = $supported
                "Supported Releases"        = $supportedReleases
                Note                        = $note
                Reference                   = $reference
            }
        }
    }
    processNic($hosts, $vdswitches) {
        Write-Host "`tProcessing ESXi IO..." -NoNewline
        $vmHosts = $hosts | Where-Object { @("Connected", "Maintenance") -contains $_.ConnectionState }
        $devices = $vmHosts | Get-VMHostPciDevice | Where-Object { $_.DeviceClass -eq "MassStorageController" -or $_.DeviceClass -eq "NetworkController" -or $_.DeviceClass -eq "SerialBusController" } 
        for ($count = 0; $count -lt $devices.length; $count++) {
            Show-Progress $devices.length ($count + 1)
            $device = $devices[$count]
            if ($device.DeviceName -like "*USB*" -or $device.DeviceName -like "*iLO*" -or $device.DeviceName -like "*iDRAC*") {
                continue
            }
            
           
            $vid = [String]::Format("{0:x4}", $device.VendorId)
            $did = [String]::Format("{0:x4}", $device.DeviceId)
            $svid = [String]::Format("{0:x4}", $device.SubVendorId)
            $ssid = [String]::Format("{0:x4}", $device.SubDeviceId)
            $hclReference = $this.getHCLReference()
            foreach ($entry in $hclReference.data.ioDevices) {
                if (($vid -eq $entry.vid) -and ($did -eq $entry.did) -and ($svid -eq $entry.svid) -and ($ssid -eq $entry.ssid)) {
                    $isNic = $device.DeviceClass -eq "NetworkController"
                    $isHba = $device.DeviceClass -eq "MassStorageController" -or $device.DeviceClass -eq "SerialBusController"
                    $deviceData = $null
                    $vmnicdetail = $null
                    $standardSwitches = $null
                    $version = $null
                    $vsanCompatibility = $null
                    $esxicli = $device.VMHost | Get-EsxCli -V2
                    if ($isNic) {
                        $standardSwitches = $this.getHostStandardSwitches($device.VMHost)
                        $deviceData = $esxicli.network.nic.list.Invoke() | Where-Object { $_.PCIDevice -like '*' + $device.Id }
                        $vmnicdetail = $esxicli.network.nic.get.Invoke(@{nicname = $deviceData.Name })
                        $version = $vmnicdetail.DriverInfo.Version
                        $vsanCompatibility = $null
                    }
                    elseif ($isHba) {
                        $deviceData = $device.VMHost | Get-VMHostHba -ErrorAction SilentlyContinue | Where-Object { $_.PCI -like '*' + $device.Id } 
                        $vibname = $deviceData.Driver -replace "_", "-"
                        $version = ($esxicli.software.vib.list.invoke() | Where-Object { $_.Name -match "^(scsi-|sata-|)$vibname" }).Version -join ", "
                        $vsanCompatibility = Get-VsanHclDatabase $vid $did $svid $ssid $($this.vsanhcl)
                    }
                    $this.nicReport += [PSCustomObject] @{
                        "vCenter"                 = $this.currentVCenter;
                        "Hostname"                = $device.VMHost.name;
                        "Placa"                   = if ($isNic) { $deviceData.Name } elseif ($isHba) { $deviceData.Device };
                        "Controlador"             = $device.DeviceName;
                        "Vendor"                  = $device.VendorName;
                        "Driver"                  = if ($isNic) { $vmnicdetail.DriverInfo.Driver } elseif ($isHba) { $deviceData.Driver };
                        "Version"                 = $version;
                        "Firmware"                = if ($isNic) { $vmnicdetail.DriverInfo.FirmwareVersion } else { $null };
                        "Vid"                     = $vid;
                        "Did"                     = $did;
                        "Svid"                    = $svid;
                        "ssid"                    = $ssid;
                        "ESXi Release"            = $device.VMHost.ExtensionData.Summary.Config.Product.FullName;
                        "ESXi Supported Releases" = $entry.releases -join ",";
                        "URL"                     = $entry.url;
                        "vSAN Compatibility"      = $vsanCompatibility -join ","
                        "Switch"                  = if ($isNic) { $this.getSwitchNameForPNic($deviceData, $vdswitches, $standardSwitches) }else { $null };
                    }
                    break;
                }
            }
        }
    }
    processVm($vms, $clusters) {
        Write-Host "`tProcessing network adapters..." -NoNewline
        $snapshots = $vms | Get-Snapshot
        $connectedIsos = $vms | Get-CDDrive | Where-Object { $null -ne $_.IsoPath }
        $allNetworkAdapters = $vms | Get-NetworkAdapter
        $this.processvNetwork($allNetworkAdapters)
        Write-Host "`tProcessing VMs..." -NoNewline
        for ($count = 0; $count -lt $vms.length; $count++) {
            $vm = $vms[$count]
            Show-Progress $vms.length ($count + 1)
            $networkAdapters = $allNetworkAdapters | Where-Object { $_.ParentId -eq $vm.Id }
            $nicCount = $networkAdapters.length
            $toolsRequredVersion = $this.getToolsReference($vm.VMHost)
            $this.vmReport += [PSCustomObject]@{
                vCenter              = $this.currentVCenter
                VM                   = $vm.Name
                Cluster              = ($clusters | Where-Object { $_.ExtensionData.Host -contains $vm.VMHost.Id }).name
                Host                 = $vm.VMHost.name
                ConnectionState      = $vm.ExtensionData.Runtime.ConnectionState.ToString()
                State                = $vm.PowerState.ToString()
                vCPU                 = $vm.NumCpu
                "Memory MB"          = $vm.MemoryMB
                HardwareVersion      = $vm.HardwareVersion
                Snapshots            = ($snapshots | Where-Object { $_.VM -eq $vm }).length
                ToolsStatus          = try { $vm.ExtensionData.Guest.ToolsStatus.ToString() } Catch { "VM has not been scanned" };
                ToolsVersion         = $vm.ExtensionData.Guest.ToolsVersion
                ToolsRequiredVersion = $toolsRequredVersion
                "SO (vCenter)"       = $vm.ExtensionData.Config.GuestFullName
                "SO (Tools)"         = $vm.ExtensionData.Guest.GuestFullName
                IsoConnected         = ($connectedIsos | Where-Object { $_.ParentId -eq $vm.Id }).IsoPath
                Adapter_01           = if ($nicCount -gt 0) { $networkAdapters[0].Type.ToString() } else { "" } 
                Adapter_02           = if ($nicCount -gt 1) { $networkAdapters[1].Type.ToString() } else { "" } 
                Adapter_03           = if ($nicCount -gt 2) { $networkAdapters[2].Type.ToString() } else { "" } 
                Adapter_04           = if ($nicCount -gt 3) { $networkAdapters[3].Type.ToString() } else { "" } 
                Adapter_05           = if ($nicCount -gt 4) { $networkAdapters[4].Type.ToString() } else { "" } 
                Adapter_06           = if ($nicCount -gt 5) { $networkAdapters[5].Type.ToString() } else { "" } 
                Adapter_07           = if ($nicCount -gt 6) { $networkAdapters[6].Type.ToString() } else { "" } 
                Adapter_08           = if ($nicCount -gt 7) { $networkAdapters[7].Type.ToString() } else { "" } 
                Adapter_09           = if ($nicCount -gt 8) { $networkAdapters[8].Type.ToString() } else { "" } 
                Adapter_10           = if ($nicCount -gt 9) { $networkAdapters[9].Type.ToString() } else { "" } 
            }
        }
    }
    processvNetwork($allNetworkAdapters) {
        for ($count = 0; $count -lt $allNetworkAdapters.length; $count++) {
            Show-Progress $allNetworkAdapters.length ($count + 1)
            $networkadapter = $allNetworkAdapters[$count]
            $this.vNetworkReport += [PSCustomObject]@{
                vCenter         = $this.currentVCenter
                VM              = $networkadapter.Parent.Name
                Cluster         = ($networkadapter.Parent.VMHost | Get-Cluster).Name
                Host            = $networkadapter.Parent.VMHost.Name
                Status          = $networkadapter.Parent.PowerState
                Mac             = $networkadapter.MacAddress
                Connected       = if ($networkadapter.ConnectionState.Connected) { "True" } else { "False" }
                StartsConnected = if ($networkadapter.ConnectionState.StartConnected) { "True" } else { "False" }
            }
        }
    }
    processLicense() {
        Write-Host "`tProcessing licenses..."
        foreach ($licenseManager in (Get-View LicenseManager)) {
            #-Server $vCenter.Name
            foreach ($license in $licenseManager.Licenses) {
                $licenseProp = $license.Properties
                $licenseExpiryInfo = $licenseProp | Where-Object { $_.Key -eq 'expirationDate' } | Select-Object -ExpandProperty Value
                if ($license.Name -eq 'Product Evaluation') {
                    $expirationDate = 'Evaluation'
                } #if ($license.Name -eq 'Product Evaluation')
                elseif ($null -eq $licenseExpiryInfo) {
                    $expirationDate = 'Never'
                } #elseif ($null -eq $licenseExpiryInfo)
                else {
                    $expirationDate = $licenseExpiryInfo
                } #else #if ($license.Name -eq 'Product Evaluation')
    
                if ($license.Total -eq 0) {
                    $totalLicenses = 'Unlimited'
                } #if ($license.Total -eq 0)
                else {
                    $totalLicenses = $license.Total
                } #else #if ($license.Total -eq 0)

                $productName = $licenseProp | Where-Object { $_.Key -eq 'ProductName' } | Select-Object -ExpandProperty Value
                $productVersion = $licenseProp | Where-Object { $_.Key -eq 'ProductVersion' } | Select-Object -ExpandProperty Value
    
                $this.vLicenseReport += [PSCustomObject]@{
                    Name           = $license.Name
                    LicenseKey     = $license.LicenseKey
                    ExpirationDate = $expirationDate
                    ProductName    = if ($null -eq ($productName)) { "No product name" } else { $productName } 
                    ProductVersion = if ($null -eq ($productName)) { "No product version" } else { $productVersion }
                    EditionKey     = $license.EditionKey
                    Total          = $totalLicenses
                    Used           = $license.Used
                    CostUnit       = $license.CostUnit
                    vCenter        = $this.currentVCenter
                }
            } #foreach ($license in $licenseManager.Licenses)
        }
    }
    processDatastore($hosts) {
        Write-Host "`tProcessing Datastores..." -NoNewline

        for ($count = 0; $count -lt $hosts.length; $count++) {
            Show-Progress $hosts.length ($count + 1)
            $h = $hosts[$count]
            $datastores = $h | Get-Datastore | Where-Object { $_.Type -notlike "vsan" }
            $esxName = $h.Name
            $esx = Get-View -ViewType HostSystem -Property Name, Config.StorageDevice -Filter @{"Name" = "^$esxName" }
            foreach ($lun in $esx.Config.StorageDevice.MultipathInfo.Lun) {
                $scsiLun = $esx.Config.StorageDevice.ScsiLun | Where-Object { $_.Key -eq $lun.Lun }
                $canon = $scsiLun.CanonicalName
                $datastore = ($datastores | Where-Object { ($_.extensiondata.info.vmfs.extent | Select-Object -expand diskname) -like $canon }).name

                if ($null -ne $datastore) {
                    $policy = if ($lun.Policy.Policy -match "_FIXED") { "Fixed" }
                    elseif ($lun.Policy.Policy -match "_MRU") { "MostRecentlyUsed" }
                    elseif ($lun.Policy.Policy -match "_RR") { "RoundRobin" }
                    else { "Unknown" }

                    $this.datastoreReport += [PSCustomObject] @{
                        vCenter   = $this.currentVCenter
                        Hostname  = $esxName
                        Datastore = $datastore
                        Policy    = $policy
                    }
                }
            }
        }
    }
    processSwitch($hosts) {
        Write-Host "`tProcessing Standard Switches..." -NoNewline
        $totalPG = 0
        for ($count = 0; $count -lt $hosts.length; $count++) {
            Show-Progress $hosts.length ($count + 1)
            $h = $hosts[$count]
            $cluster = ($h | Get-Cluster).Name
            foreach ($sw in ($h | Get-VirtualSwitch -Standard)) {
                $portGroups = $sw | Get-VirtualPortGroup
                $totalPG += $portGroups.length
                foreach ($pg in $portGroups) {
                    if ($null -ne $pg.vLanId) {
                        $this.switchReport += [PSCustomObject] @{
                            vCenter   = $this.currentVCenter
                            ESXi      = $h.name
                            Cluster   = $cluster
                            PortGroup = $pg.Name
                            Switch    = $sw.Name
                            vLAN      = $pg.vLanId
                        }
                    }
                }
            }
        }
    }
    processSnapshot($vms) {
        Write-Host "`tProcessing Snapshots...";
        $snap = $vms | Get-Snapshot

        for ($count = 0; $count -lt $snap.length; $count++) {
            Show-Progress $snap.length ($count + 1)
            $s = $snap[$count]
            $this.snapshotReport += [PSCustomObject] @{
                vCenter  = $this.currentVCenter
                VM       = $s.VM.Name
                Snapshot = $s.Name
                Fecha    = ($s.Created | Get-Date -Format "yyyy-MM-dd HH:mm")
                SizeMB   = [int]$s.SizeMB
            }
        }
    }
    processPartitions($vms) {
        Write-Host "`tProcessing Partitions..." -NoNewline
        for ($count = 0; $count -lt $vms.length; $count++) {
            $vm = $vms[$count];
            Show-Progress $vms.length ($count + 1)
            if ($vm.ExtensionData.Config.Annotation -in $this.annotations) {
                foreach ($partition in $vm.ExtensionData.Guest.Disk) {
                    $freePercentage = [math]::Round(($partition.FreeSpace / $partition.Capacity) * 100, 2);
                    $this.partitionReport += [PSCustomObject]@{
                        vCenter    = $this.currentVCenter
                        VM         = $vm.name
                        Annotation = $vm.ExtensionData.Config.Annotation
                        Disk       = $partition.DiskPath
                        "Free %"   = $freePercentage
                    }
                }
            }
        }
    }
    processKernelAdapters($hosts) {
        Write-Host "`tProcessing VMkernel Adapters..." -NoNewline
        $kernelAdapters = $hosts | Get-VMHostNetworkAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "vmk[0-9]+" }
        for ($count = 0; $count -lt $kernelAdapters.length; $count++) {
            Show-Progress $kernelAdapters.length ($count + 1)
            $ka = $kernelAdapters[$count]
            $this.kernelAdaptersReport += [PSCustomObject] @{
                Host       = $ka.VMHost.Name
                Name       = $ka.Name
                IP         = $ka.IP
                MTU        = $ka.MTU
                PortGroup  = $ka.PortGroupName 
                Management = $ka.ManagementTrafficEnabled
                vMotion    = $ka.VMotionEnabled
            }
        }
    }
    processVcenterSizing($vms, $hosts) {
        Write-Host "`tProcessing Sizing...";
        foreach ($vm in $vms) {
            if ($vm.ExtensionData.Config.Annotation -in $this.annotations) {
                for ($i = $this.vCenterSizing.vsphere.length - 1; $i -ge 0; $i--) {
                    if (($vm.NumCpu -ge $this.vCenterSizing.vsphere[$i].vcpus) -and ($vm.MemoryGB -ge $this.vCenterSizing.vsphere[$i].ram)) {
                        $this.sizingReport += [PSCustomObject]@{
                            vCenter             = $this.currentVCenter;
                            VM                  = $vm.name;
                            Annotation          = $vm.ExtensionData.Config.Annotation;
                            vCPU                = $vm.NumCpu;
                            "Memory GB"         = $vm.MemoryGB;
                            "Cantidad de VMs"   = $vms.length;
                            "Cantidad de Hosts" = $hosts.length;
                            "Sizing actual"     = $this.vCenterSizing.vsphere[$i].ToString
                        }
                    }
                }
            }
        } 
    }
    processvDS($vdswitches) {
        Write-Host "`tProcessing vDS...";
        foreach ($vds in $vdswitches) {
            #TODO: Agregar nombre de cliente al nombre del zip
            Export-VDSwitch -VDSwitch $vds -Destination ($global:CONFIG.REPORTS_FOLDER + "/vds_configuration/" + $vds.Name + "_" + (Get-Date).toString('yyyy-MM-dd HHmmss') + ".zip")
            $niocEnabled = if ($vds.ExtensionData.Config.NetworkResourceManagementEnabled) { "True" } else { "False" }
            $this.vdsReport += [PSCustomObject]@{
                Name           = $vds.Name
                MTU            = $vds.Mtu
                "NIOC Enabled" = $niocEnabled
            }
        }
    }
    processVcenter($vcenterData){
        # Este método recibe la lista de objetos de vCenter que ya fue creada en el script principal.
        # Su única responsabilidad es añadir esos datos al contenedor del reporte ($this.vCenterReport).
        # Usamos .AddRange() porque es una forma eficiente de añadir todos los elementos de una colección a un ArrayList.
        $this.vCenterReport.AddRange($vcenterData)
    }
    executeAlarm($hosts) {
        Write-Host "`tEjecutando falso positivo..."
        $method = Read-Host "Ingresar el metodo (Script/Email) o N para omitir"  
        if ($method -ne "N") {
            $action = Read-Host "Ingrese la linea de script o la casilla de alarmas segun corresponda"
            foreach ($h in $hosts) {
                if (@("Connected", "Maintenance") -contains $h.ConnectionState) {
                    Write-Host "Creando alarma en $($h.Name).."
                    New-Alarm $h $this.currentVCenter $method $action
                    Start-Sleep -Seconds 5
                    Write-Host "Removiendo alarma.."
                    Remove-AlarmDefinition "Falso positivo $($this.currentVCenter)"
                    break
                }
            }
        }
    }




    setCurrentVcenter($vcenter) {
        $this.currentVCenter = $vcenter
    }

    [PSCustomObject] getReport() {
        return [PSCustomObject] @{
            "vCenter"           = $this.vCenterReport;
            "ESXi"              = $this.esxiReport;
            "ESXi IO"           = $this.nicReport;
            "VM"                = $this.vmReport;
            "Datastores"        = $this.datastoreReport;
            "Standard Switch"   = $this.switchReport;
            "VMkernel Adapters" = $This.kernelAdaptersReport;
            "Snapshot"          = $this.snapshotReport;
            "Partitions"        = $this.partitionReport;
            "Sizing"            = $this.sizingReport;
            "vDS"               = $this.vdsReport;
            "vNetwork"          = $this.vNetworkReport;
            "vLicense"          = $this.vLicenseReport
        }
    }

    [PSCustomObject] getHCLReference() {    
        if ($null -eq $this.ioHclRef) {
            $this.ioHclRef = Get-Content -Path ($global:CONFIG.PLUGINS_FOLDER + "\lib\data\vmware-iohcl.json")
            $this.ioHclRef = $this.ioHclRef | ConvertFrom-Json
            return $this.ioHclRef
        }
        else {
            return $this.ioHclRef
        }
    }

    [Array] getHostNetworkAdaptersForVDS($vds) {
        if ($null -eq ($this.VMHostNetworkAdapters | Get-Member -Name $vds.name)) {
            $this.VMHostNetworkAdapters | Add-Member -NotePropertyName $vds.name -NotePropertyValue (Get-VMHostNetworkAdapter -DistributedSwitch $vds)
        }
        return $this.VMHostNetworkAdapters.($vds.name)
    }

    [String] getSwitchNameForPNic($pnic, $vdswitches, $standardSwitches) {
        foreach ($vdswitch in $vdswitches) {
            $networkAdapters = $this.getHostNetworkAdaptersForVDS($vdswitch)
            $thisAdapter = $networkAdapters | Where-Object { $_.mac -eq $pnic.MACAddress }
            if ($null -ne $thisAdapter) {
                return $vdswitch.name
            }
        }
        foreach ($sSwitch in $standardSwitches) {
            if ($sSwitch.nic -contains $pnic.name) {
                return $sSwitch.name
            }
        }
        return "None"
    }

    [Array] getHostStandardSwitches($vmhost) {
        if ($null -eq ($this.VMHostStandardSwitches | Get-Member -Name $vmhost.name)) {
            $this.VMHostStandardSwitches | Add-Member -NotePropertyName $vmhost.name -NotePropertyValue ($vmhost | Get-VirtualSwitch -Standard)
        }
        return $this.VMHostStandardSwitches.($vmhost.name)
    }

    [String] getToolsReference($esxi) {
        $toolsbuild = $this.toolsReference | Get-Member -Name $esxi.name
        if ($null -eq $toolsbuild) {
            $this.toolsReference | Add-Member -NotePropertyName $esxi.name -NotePropertyValue ($this.toolsRefJson | Where-Object { $_.esxiBuild -eq $esxi.ExtensionData.Config.Product.build }).toolsBuild
        }
        return $this.toolsReference.($esxi.name)
    }
}