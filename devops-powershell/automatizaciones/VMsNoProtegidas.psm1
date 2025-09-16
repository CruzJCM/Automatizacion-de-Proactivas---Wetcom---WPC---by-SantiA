function Get-VMsNoProtegidasPorHA($vcenter) {
    Start-VMsNoProtegidasPorHA($vcenter)

    <#
	.Synopsis
    VMs no protegidas por HA
	.Component
	vcenter
	#>
}


function Start-VMsNoProtegidasPorHA($vcenter){
	$clusters = Get-Cluster
	$report = @()
	$index = 0
	Write-Host "Progress " -NoNewline
	foreach($cluster in $clusters){
		Show-Progress $clusters.length $index
		$vms = @()
		if($cluster.HAEnabled){
			$vms =  $cluster | Get-VM | Where-Object {$_.HARestartPriority -eq "Disabled"} | Where-Object {$_.PowerState -eq "PoweredOn"}
		}else{
			$vms =  $cluster | Get-VM | Where-Object {$_.PowerState -eq "PoweredOn"}
		}
		foreach($vm in $vms){
			$report += [PSCustomObject] @{
				VM = $vm.name
				vcenter = $vm.Uid.Substring($vm.Uid.IndexOf('@')+1).Split(":")[0]
				cluster = $cluster.name
				powerState = $vm.PowerState.ToString()
				dasProtection = $cluster.HAEnabled
				HARestartPriority = $vm.HARestartPriority.ToString()
			}
		}
		$index++
		Show-Progress $clusters.length $index
	}

	Write-Host "Done!"

	Write-Host "RESUMEN:"
	Write-Host "VMs no protegidas por HA: " $report.length

	$file = [PSCustomObject] @{
		Result=if($report.length -eq 0){"Ok"}else{"Alert"};
		Name="VMs no protegidas por HA";
		DateTime= (Get-Date -Format "yyyy-MM-dd HH:mm");
		LocalHost= [system.environment]::MachineName;
		User = whoami
		Endpoint=$vcenter.host;
		Component="vcenter";
		Report=@($report)
		IdAutomatizacion="dm1zRGVzcHJvdGVnaWRhczpQYXNzd29yZDEyMyQ="
	}

	$file | ConvertTo-Json | Set-Content ($global:CONFIG.REPORTS_FOLDER + "/" + (Get-Date).toString('yyyy-MM-dd HHmmss') + "_" + $file.Name + ".json")
}