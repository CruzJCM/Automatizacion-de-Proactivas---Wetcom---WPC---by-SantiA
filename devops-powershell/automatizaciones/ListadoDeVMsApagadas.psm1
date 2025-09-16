$CURRENT_FOLDER = Split-Path $script:MyInvocation.MyCommand.Path

Import-Module $CURRENT_FOLDER\lib\vrops.psm1

$credenciales = "dm1zX2FwYWdhZGFzOlBhc3N3b3JkMTIzJA=="

function Get-VMsApagadas($vrops) {
    Start-ListadoVMsApagadas($vrops)

    <#
	.Synopsis
    Listado de VMs apagadas
	.Component
	vrops
	.Role
	#>
}

function Start-ListadoVMsApagadas($vrops) {
    $localhost = [system.environment]::MachineName
    
    $vms = get-VMsOff($vrops)

    $result = @()
    $propertyname = "summary|runtime|powerState"

    Write-Host "Progress " -NoNewline
    $index = 0

    foreach ($vm in $vms){
        Show-Progress $vms.length $index
        $vmIdentifier = $vm.identifier

        $uri = $URI_RESOURCE_PROPERTIES -f $vrops.host, $vmIdentifier

        $properties = (Get-RequestResult $uri $vrops.conn).property
            
        $vmName = Get-ValueOfProperty $properties "config|name"
        $dataStore = Get-ValueOfProperty $properties "summary|datastore"
        $vCenter = Get-ValueOfProperty $properties "summary|parentVcenter"
        $rep = [PSCustomObject] @{
            VM = $vmName
            Datastore = $dataStore
            vCenter = $vCenter
        }
        $result += $rep
        
        $index ++
		Show-Progress $vms.length $index
    }
    Write-Host "Done!"

    Write-Host "RESUMEN:"
    Write-Host "VMs Apagadas:     " $result.Length
    
    $report = @{
        "Name" = "Listado de VMs apagadas"
        "DateTime" = (Get-Date).toString('yyyy-MM-dd HH:mm')
        "LocalHost" = $localhost
        "User" = whoami
        "Endpoint" = $vrops.host
        "Component" = "vrops"
        "Result" = if ($result.Length -eq 0) {"Success"} else {"Alert"}
        "Report" = @($result)
        "IdAutomatizacion"=$credenciales
    }
    
    $report | ConvertTo-Json | Set-Content ($global:CONFIG.REPORTS_FOLDER + "/" + (Get-Date).toString('yyyy-MM-dd HHmmss') + "_" + $report.Name + ".json")
}


function get-VMsOff($vrops) {
    $resource = 'virtualmachine' + "&propertyName=summary|runtime|powerState&propertyValue=Powered Off"
    $uri = $URI_RESOURCES -f $vrops.host, $resource
    return (Get-RequestResult $uri $vrops.conn).resourceList
}