class UnprotectedVMs {
    [String] $currentVCenter = "No vCenter"
    $unprotectedReport = [System.Collections.ArrayList]@()

    processUnprotectedVMs($vms, $backupCategory) {
        Write-Host "`tProcesando VMs..." -NoNewLine
        for ($i = 0; $i -lt $vms.length; $i++) {
            Show-Progress $vms.length ($i + 1)
            $hasTag = Get-TagAssignment -Entity $vms[$i] -Category $backupCategory
            if ([string]::IsNullOrEmpty($hasTag)) {
                $this.unprotectedReport += [PSCustomObject]@{
                    vCenter         = $vms[$i].Uid.Split(":")[0].Split("@")[1]
                    VMName          = $vms[$i].Name
                    Cluster         = $vms[$i].VMHost.Parent.Name
                    Host            = $vms[$i].VMHost.Name
                    PowerState      = if ($vms[$i].PowerState) { "PoweredOn" }else { "PoweredOff" }
                    $backupCategory = ""
                }
            }
        }
    }


    setCurrentVcenter($vcenter) {
        $this.currentVCenter = $vcenter
    }

    [PSCustomObject] getReport() {
        return [PSCustomObject] @{
            "Unprotected VMs" = $this.unprotectedReport;
        }
    }
}