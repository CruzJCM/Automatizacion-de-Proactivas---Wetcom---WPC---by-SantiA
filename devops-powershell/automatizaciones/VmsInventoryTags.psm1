$CURRENT_FOLDER = Split-Path $script:MyInvocation.MyCommand.Path

Import-Module $CURRENT_FOLDER\lib\vrops.psm1

function Get-TagsInventory($vrops) {
    Start-TagsInventory($vrops)

    <#
	.Synopsis
    Inventario de Tags en VMs
	.Component
	vrops
	.Role
	ui
	#>
}

function Start-TagsInventory($vrops) {
    $category_file_route = Read-Host "Ingrese la ruta del archivo con las categorias a analizar"
    $category_file = Get-Content $category_file_route
    Write-Host "Obteniendo VMs..."
    $resources = Get-NotDeletedVms($vrops)
    Write-Host "Obteniendo IDs..."
    $ids = Get-ResourceIds($resources)
    Write-Host "Obteniendo Properties..."

    $props = @( "config|name", "summary|parentVcenter", 
                "summary|tagJson", "summary|parentCluster",
                "summary|datastore"
                "summary|parentDatacenter", "summary|parentHost",
                "config|guestFullName", "summary|guest|fullName")
    $body = @{"resourceIds"=@($ids); "propertyKeys" = $props} | ConvertTo-Json
    $uri = $URI_QUERY_PROPERTIES -f $vrops.host
   
    $vms = Get-PostRequestResult $uri $body $vrops.conn
    $index = 0
    Write-Host "Analizando Properties..."
	Write-Host "Progress " -NoNewline
    $result = @()

    foreach ($vm in $vms) {
        Show-Progress $vms.length $index
        $properties = $vm.'property-contents'.'property-content'
        
        $tags = Get-ValueOfLastestProperties $properties "summary|tagJson"
        $name = Get-ValueOfLastestProperties $properties "config|name"        
       
        $result += [PSCustomObject] @{
            VM = $name
            CreationTime = Get-CreationTime $resources $vm.'resourceId'
            vCenter = Get-ValueOfLastestProperties $properties "summary|parentVcenter"
            Datacenter = Get-ValueOfLastestProperties $properties "summary|parentDatacenter"
			Cluster = Get-ValueOfLastestProperties $properties "summary|parentCluster"
            Host = Get-ValueOfLastestProperties $properties "summary|parentHost"
            Datastore = Get-ValueOfLastestProperties $properties "summary|datastore"
            "OS segun config file" = Get-ValueOfLastestProperties $properties "config|guestFullName"
            "OS segun VMWare Tools" = Get-ValueOfLastestProperties $properties "summary|guest|fullName"
        }
        
        foreach ($category in $category_file) {
            $result | Select-Object -Last 1 | Add-Member -MemberType NoteProperty -Name $category -Value (Get-ValueOfTag $tags $category)
        }
        
        $index++
        Show-Progress $vms.length $index
    }

    Write-Host "Done!"
    Write-Host "RESUMEN:"
	Write-Host "VMs analizadas: " $vms.length

    $report  = @{
        "Name" = "Listado de vSphere Tags asignados"
        "DateTime" = (Get-Date).toString('yyyy-MM-dd HH:mm')
        "LocalHost" = [system.environment]::MachineName
        "User" = whoami
        "Endpoint" = $vrops.host
        "Component" = "vrops"
        "Result" = "Success"
        "Report" = @($result)
        "IdAutomatizacion"= "dnNwaGVyZVRhZ3M6UGFzc3dvcmQxMjMk"
    }
    
    $report | ConvertTo-Json | Set-Content ($global:CONFIG.REPORTS_FOLDER + "/" + (Get-Date).toString('yyyy-MM-dd HHmmss') + "_" + $report.Name + ".json")
}



function Get-NotDeletedVms($vrops) {
    $uri = $URI_EXISTING_VMS -f $vrops.host
    return (Get-RequestResult $uri $vrops.conn).resourceList
}

function Get-ResourceIds($resourceList) {
    $result = @()
    foreach ($res in $resourceList) {
        $result += $res.identifier
    }
    return $result
}

function Get-CreationTime($resources, $resourceId) {
    $vm = $resources | Where-Object {$_.identifier -eq $resourceId}
    $date = (Get-Date 01.01.1970)+([System.TimeSpan]::FromMilliseconds($vm.'creationTime'))
    return $date
}