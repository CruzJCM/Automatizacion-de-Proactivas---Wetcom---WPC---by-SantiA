$URI_RELATED = 'https://{0}/suite-api/api/resources/{1}/relationships'
$URI_RESOURCE_PROPERTIES = 'https://{0}/suite-api/api/resources/{1}/properties'
$URI_RESOURCES = 'https://{0}/suite-api/api/resources/?resourceKind={1}'
$URI_LATEST_STATS = 'https://{0}/suite-api/api/resources/{1}/stats/latest'
# resourceStatus=DATA_RECEIVING en lugar de resourceState=STARTED para vROps <7.5
$URI_EXISTING_VMS = 'https://{0}/suite-api/api/resources/?resourceKind=VirtualMachine&resourceState=STARTED&propertyName=summary|config|isTemplate&propertyValue=false&pageSize=9999'
$URI_QUERY_PROPERTIES = 'https://{0}/suite-api/api/resources/properties/latest/query'


function Get-RequestResult ($uri, [pscredential]$cred) {
    $headers = @{"Accept" = "application/json"}
    $ProgressPreference = 'silentlyContinue'
    $result = Invoke-WebRequest -SkipCertificateCheck -Uri $uri -Credential $cred -Headers $headers
	$ProgressPreference = 'Continue'
	$jsonString = $result.Content.Replace("stat-list", "statList")
    $json = ConvertFrom-Json $jsonString
    return $json
}

function Get-PostRequestResult ($uri, $body, [pscredential]$cred) {
    $headers = @{"Accept" = "application/json"; "Content-Type" = "application/json"}
    $ProgressPreference = 'silentlyContinue'
    $result = Invoke-RestMethod -Method "POST" -Uri $uri -Body $body -Headers $headers -Credential $cred -SkipCertificateCheck
	
	$ProgressPreference = 'Continue'
	
    return $result.values
}


function Get-ValueOfProperty ($props,$propertyname) {
    foreach ($p in $props) {
        if ($p.name -eq $propertyname) {
            $result = $p.value
            break
        }
    }
    return $result
}


function Get-ValueOfLatestStat ($stats, $statname){
	foreach($stat in $stats.values[0].statList.stat){
		if($stat.statKey.key -eq $statname){
			$result = $stat.data[0]
			break
		}
	}
	return $result
}


function Get-RelatedByResourceKind ($data, $resourcekind) {
	$result = @()
	foreach ($resource in $data.resourceList) {
		if($resource.resourceKey.resourceKindKey -eq $resourcekind){
			$result += $resource
		}
	}
	return $result
}

function Get-ValueOfLastestProperties($properties, $propName){
	foreach ($prop in $properties) {
		if($prop.statKey -eq $propName){
			$result = $prop.values
			}
	}
	return $result
}

function Get-ValueOfTag($tags, $tagName){
	$result = @()
	if ($tags -ne "none" -and $tags -ne $null){
		$tags = ConvertFrom-Json $tags
		foreach ($tag in $tags) {
			if($tag.category -eq $tagName){
				$result += $tag.name
			}
		}
	}	
	return ($result -join ", ")
	
}


Export-ModuleMember -Variable * -Function *