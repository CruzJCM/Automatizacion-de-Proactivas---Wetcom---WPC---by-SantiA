using module .\lib\unprotectedvms.psm1


function Get-UnprotectedVMs($vcenters) {
    Start-UnprotectedVMs($vcenters)

    <#
	.Synopsis
    VMs no protegidas (Veeam)
	.Component
	vcenter
    .Role
    ui
	#>
}

function Start-UnprotectedVMs ($vcenters) {
    $unprotectedVMs = New-Object UnprotectedVMs
    $backupCategory = Read-Host "Ingrese la categoria usada para backup"

    foreach($vcenter in $vcenters.conn){
        $unprotectedVMs.setCurrentVcenter($vcenter)
        
        $vms = Get-VM -Server $vcenter
        $vms += Get-Template -Server $vcenter
        $unprotectedVMs.processUnprotectedVMs($vms,$backupCategory)
    }

    $file = [PSCustomObject] @{
		Result="OK";
        Name="Unprotected VMs";
		DateTime= (Get-Date -Format "yyyy-MM-dd HH:mm");
		LocalHost= [system.environment]::MachineName;
		User = whoami;
		Endpoint=$vcenters.host;
		Component="vcenter";
		Report = $unprotectedVMs.getReport();
		IdAutomatizacion=$credenciales;
    }

	$file | ConvertTo-Json -Depth 99| Set-Content ($global:CONFIG.REPORTS_FOLDER + "/" + (Get-Date).toString('yyyy-MM-dd HHmmss') + "_" + $file.Name + ".json")

}