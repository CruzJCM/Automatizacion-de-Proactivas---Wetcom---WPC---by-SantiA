$global:connections = @()

function Set-Endpoints($modules) {
	$components = @()

	foreach ($module in $modules) {
			$cs = $module.component.Split(";")
			foreach ($c in $cs) {
				if (!$components.Contains($c)) {
					$components += $c
					$Global:connections += [PSCustomObject] @{component=$c; host=""; conn=$null}
				}
			}
	}
}

function Test-vRopsConnection($hostName, $cred){
	$uri = "https://{0}/suite-api/api/versions/current/" -f $hostName
	$headers = @{"Accept" = "application/json"}
	try{
		$result = Invoke-WebRequest -SkipCertificateCheck -Uri $uri -Credential $cred -Headers $headers
	}catch{
		Write-Host "Connection Failure!"
		if($_ -like "No such host is known"){
			Write-Host "Unkown host. Try again."
			return $false
		}

		if($_.Exception.Response.ReasonPhrase -like 401 -or $_.Exception.Response.statusCode -like "Unauthorized"){
			Write-Host "Invalid credentials. Try again."
			return $false
		}

		if($_.FullyQualifiedErrorId -like "CannotConvertArgumentNoMessage,Microsoft.PowerShell.Commands.InvokeWebRequestCommand"){
			Write-Host "Invalid hostname. Try again."
			return $false
		}

		if($_.Exception.Message -like "A connection attempt failed because the connected party did not properly respond after a period of time, or established connection failed because connected host has failed to respond"){
			Write-Host "Response timeout, the hostname did not respond anything. Try again."
			return $false
		}

		
		$result = $_
	}
	if($result.statusCode -like 200){
		return $true
	}
	Write-Host $result
	Write-Host "Unexpected exception!"
	return $false
}

function Test-vCenterConnection(){
	try{
		$connection = Connect-VIServer -ea Stop
	}catch{
		Write-Host "Connection Failure!"
	
		Write-Host "Failed to connect. Try again."
		return $false
	}

	return $connection
}

function Connect-Endpoints($modules) {
	foreach ($module in $modules) {
		$endpoints = $Global:connections | 
						Where-Object {$_.component -in $module.component.Split(";") -and $_.host -eq ""}
		foreach ($endpoint in $endpoints) {	
			if($endpoint.component -eq "vrops"){
				$hostName = Read-Host "vROps hostname"
				$cred = Get-Credential -Message "Enter vRops credentials"
				if(Test-vRopsConnection $hostName $cred){
					$endpoint.host = $hostName
					$endpoint.conn = $cred
				}else{
					return $false
				}
			}elseif ($endpoint.component -eq "vcenter") {
				$connection =  Test-vCenterConnection
				if($connection){
					$endpoint.conn = $connection
					$endpoint.host = $Global:DefaultVIServers -join ", "
				}else{
					return $false
				}
			}
		}	
	}
	return $true
}

function Connect-Endpoint($component){
	# if($component -eq "vcenter"){
	# 	"Debe conectarse a un vCenter. Ingrese los parametros para conectarse"
	# 	$server = Read-Host "Host"
	# 	$user = Read-Host "Usuario"
	# 	$pass = Read-Host "Password" -AsSecureString
	# 	$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass)
	# 	$UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
	# 	$connection = Connect-VIServer -Server $server -u $user -pass $UnsecurePassword
	# 	$newConnection = [PSCustomObject] @{component=$component; host=$connection.Name; conn=$connection}
	# 	$global:connections += $newConnection
	# }
}

function Disconnect-Endpoints {
	# $Global:connections|Where-Object{$_.Component -eq "vcenter"} |
	# 	ForEach-Object {Disconnect-VIServer -Server $_.host}
	foreach ($conn in $global:connections) {
		if($conn.component -eq "vcenter"){
			Disconnect-VIServer -Confirm
		}
		$conn.host = ""
		$conn.conn = $null
	} 
}