$CURRENT_FOLDER = Split-Path $script:MyInvocation.MyCommand.Path

Import-Module $CURRENT_FOLDER\lib\vrops.psm1

$credenciales = "ZGF0YXN0b3JlX2xvY2FsOlBhc3N3b3JkMTIzJA=="
function Get-VMsDatastoresLocales($vrops) {
    Start-VMsLocales($vrops)

    <#
	.Synopsis
    VMs en datastores locales
	.Component
	vrops
    .Role
	#>
}

function Assert-IsLocalVM ($props, $datastore) {
    $IsLocal = $false
    foreach ($p in $props) {
        if ($p.name -match "virtualDisk:.+\|datastore" -and $p.value -eq $datastore) {
            $IsLocal = $true
            break
        }
    }
    return $IsLocal
}


# function Remove-DuplicatedVMs ($vms) {
# 	$revisar = $vms | Where-Object {$_.status -eq "Local"}
# 	$repetidas = @()
# 	foreach($vm in $revisar){
# 		$repetidas += $vms | Where-Object{$_.vm -eq $vm.vm -and $_.status -eq "Revisar"}
# 	}
# 	return $vms | Where-Object {$repetidas -notcontains $_}
# }

function Start-VMsLocales($vrops) {
    $localhost = [system.environment]::MachineName

    $URI_DATASTORES_LOCALES = 'https://{0}/suite-api/api/resources/?resourceKind=Datastore&propertyName=summary|isLocal&propertyValue=true'
    

    $result = @()

    # Datastores Locales
    $uri = $URI_DATASTORES_LOCALES -f $vrops.host
    
    $datastores = Get-RequestResult $uri $vrops.conn
	$vms = @()
	$index = 0
	Write-Host "Progress " -NoNewline
    foreach ($ds in $datastores.resourceList) {
		Show-Progress $datastores.resourceList.length $index
        # VMS relacionadas
        $uri = $URI_RELATED -f $vrops.host, $ds.identifier
        $data = Get-RequestResult $uri $vrops.conn
        $vms = $data.resourceList | Where-Object { $_.resourceKey.resourceKindKey -eq "VirtualMachine" }
        # Verify VMs
        foreach ($vm in $vms) {
            $uri = $URI_RESOURCE_PROPERTIES -f $vrops.host, $vm.identifier
            $properties = Get-RequestResult $uri $vrops.conn
			
			$isTemplate = if((Get-ValueOfProperty $properties.property "summary|config|isTemplate") -eq "true")
							{$true}else{$false}

            if (-not $isTemplate) {
                $vCenter = Get-ValueOfProperty $properties.property "summary|parentVcenter"
                $rep = [PSCustomObject] @{
                    vcenter = $vCenter
                    datastore = $ds.resourceKey.name
                    vm = $vm.resourceKey.name
                    status = if (Assert-IsLocalVM $properties.property $ds.resourceKey.name) {"Local"} else {"Revisar"}
                }
				$result += $rep
			}
		}
		$index ++
		Show-Progress $datastores.resourceList.length $index
	}
	Write-Host "Done!"

	#$result = Remove-DuplicatedVMs $result
    
	$localvms = $result | Where-Object {$_.status -eq "Local"}

    Write-Host "RESUMEN:"
    Write-Host "Local VMs:     " $localvms.Length
    Write-Host "VMs a revisar: " ($result.length - $localvms.Length)

    $report = @{
        "Name" = "VMs en Datastores Locales"
        "DateTime" = (Get-Date).toString('yyyy-MM-dd HH:mm')
		"LocalHost" = $localhost
		"User" = whoami
        "Endpoint" = $vrops.host
        "Component" = "vrops"
        "Result" = if ($localvms.Length -gt 0) {"Alert"} elseif($result.length -gt 0) {"Warning"} else {"Ok"}
		"Report" = @($result)
		"IdAutomatizacion"=$credenciales
    }

    $report | ConvertTo-Json | Set-Content ($global:CONFIG.REPORTS_FOLDER + "/" + (Get-Date).toString('yyyy-MM-dd HHmmss') + "_" + $report.Name + ".json")
}