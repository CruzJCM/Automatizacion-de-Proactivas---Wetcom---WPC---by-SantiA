$CURRENT_FOLDER = Split-Path $script:MyInvocation.MyCommand.Path

$ESPACIO_DATASTORES=@{
	THRESHOLD_ALERT= 20
	THRESHOLD_WARNING= 30
}

Import-Module $CURRENT_FOLDER\lib\vrops.psm1

$credenciales = "ZXNwYWNpb19kYXRhc3RvcmVzOlBhc3N3b3JkMTIzJA=="

function Get-EspacioDatastores($vrops) {
    Start-EspacioDatastores($vrops)

    <#
	.Synopsis
    Espacio en datastores
	.Component
	vrops
	#>
}

function Start-EspacioDatastores($vrops) {

	$uri_vcenters = $URI_RESOURCES -f ($vrops.host, "VMwareAdapter Instance") 
	$vcenters = Get-RequestResult $uri_vcenters $vrops.conn
	
	$report = @()
	$index = 0
	Write-Host "Progress " -NoNewline
	foreach ($vcenter in $vcenters.resourceList) {
		Show-Progress $vcenters.resourceList.length $index
		$uri_datacenters = $URI_RELATED -f ($vrops.host, $vcenter.identifier)
		$result = Get-RequestResult $uri_datacenters $vrops.conn
		$datacenters =  Get-RelatedByResourceKind $result "Datacenter"
		foreach ($dc in $datacenters) {
			$uri_datastores = $URI_RELATED -f ($vrops.host, $dc.identifier)
			$result = Get-RequestResult $uri_datastores $vrops.conn


			$datastores =  Get-RelatedByResourceKind $result "Datastore"
			foreach($ds in $datastores){
				$uri_stats = $URI_LATEST_STATS -f ($vrops.host, $ds.identifier)

				$stats = Get-RequestResult $uri_stats $vrops.conn
				$space = Get-ValueOfLatestStat $stats "diskspace|freespace"
				if($space -le $ESPACIO_DATASTORES.THRESHOLD_WARNING){
					if($space -le $ESPACIO_DATASTORES.THRESHOLD_ALERT){
						$state = "Alert"
					}else{
						$state = "Warning"
					}
				}else{
					$state = "Ok"
				}
				if($state -ne "Ok"){
					$report += [PSCustomObject] @{
						vcenter = $vcenter.resourceKey.name
						datastore = $ds.resourceKey.name
						remaining_GB = [math]::Round($space,2)
						state = $state
					}
				}
			}
		}
		
		$index ++
		Show-Progress $vcenters.resourceList.length $index
	}
	Write-Host "Done!"

	$warned = ($report | Where-Object{$_.state -eq "Warning"}).length
	$alarmed = ($report | Where-Object{$_.state -eq "Alert"}).length
	

	Write-Host "RESUMEN:"
	Write-Host "Datastores con espacio alarmado: $warned"
	Write-Host "Datastores con espacio critico:  $alarmed"

	$file = [PSCustomObject] @{
		Result=if($report.length - $warned - $alarmed -eq $report.length){"Ok"}else{"Alert"};
		Name="Espacio en datastores";
		DateTime= (Get-Date -Format "yyyy-MM-dd HH:mm");
		LocalHost= [system.environment]::MachineName;
		User = whoami
		Endpoint=$vrops.host;
		Component="vRops";
		Report=@($report)
		IdAutomatizacion=$credenciales
	}

	$file | ConvertTo-Json | Set-Content ($global:CONFIG.REPORTS_FOLDER + "/" + (Get-Date).toString('yyyy-MM-dd HHmmss') + "_" + $file.Name + ".json")
}