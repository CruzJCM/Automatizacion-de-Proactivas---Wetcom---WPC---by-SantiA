function Submit-Reports{
	$zip = Compress-Reports
	Send-Reports $zip $Global:CONFIG.API_REPORTES.URL
}

function Compress-Reports{
	$user = whoami
	$user = $user -replace "\\", "-"
	$reports = @(Get-ChildItem -Path $Global:CONFIG.REPORTS_FOLDER | Where-Object{$_.Extension -eq ".json"})
	if ($reports.Length -gt 0){
		$filename = $Global:CONFIG.REPORTS_FOLDER + "/" + (Get-Date -Format "yyyy-MM-dd-HH-mm-ss-") + $user + ".zip"
		Compress-Archive -Path $reports -DestinationPath $filename
		$reports = @($reports)
		foreach($report in $reports){
			Remove-Item $report
		}
		$compressed = $reports.Length
		Write-Host "$compressed files compressed into $filename"
		return $filename
	}else{
		Write-Host "There are no new reports to send."
		return $null
	}
}

function Send-Reports ($File, $Uri){
	if($File -ne $null){
		Write-Host "Sending file $File to $Uri"
		$response = Invoke-WebRequest -Uri $Uri -Method POST -Form @{"file" = Get-Item $File} -UseBasicParsing
		if($response.StatusCode -eq 200){
			Write-Host "$File received OK"
			Move-Item -Path $File -Destination $Global:CONFIG.OLD_REPORTS_FOLDER
			$url = (ConvertFrom-Json $response.Content).url
			Write-Host "Opening $url"
			start-process $url
		}else{
			Write-Host "Failed to send $File"
			Write-Host "Make sure to be connected to the correct network and have access to devops.wetcom.net and press (r) to try again"
		}
	}
}

function Submit-PendingReports {
	$unsent = @(Get-ChildItem -Path $Global:CONFIG.REPORTS_FOLDER | Where-Object{$_.Extension -eq ".zip"})
	if ($unsent.Length -eq 0){
		Write-Host "There are no pending zip files to send"
	}else{
		foreach($report in $unsent){
			Send-Reports $report $Global:CONFIG.API_REPORTES.URL
		}
	}
}
