$CURRENT_FOLDER = Split-Path $script:MyInvocation.MyCommand.Path
If (Test-Path "$CURRENT_FOLDER/config_dev.json") {
	$global:CONFIG = Get-Content $CURRENT_FOLDER/config_dev.json | ConvertFrom-Json
} else {
	$global:CONFIG = Get-Content $CURRENT_FOLDER/config.json | ConvertFrom-Json
}
$global:PLUGINS_MODULES = @()
$global:APP_VERSION = "v1.7-1-g19aa5d0"

Import-Module $CURRENT_FOLDER/reports.psm1
Import-Module $CURRENT_FOLDER/ui.psm1
Import-Module $CURRENT_FOLDER/plugins.psm1
Import-Module $CURRENT_FOLDER/connections.psm1


function Start-App() {
    Start-Plugins
    Set-Endpoints($global:PLUGINS_MODULES)
    Register-ExecuteEventSubscription('Start-Tasks')
    Register-ClearConnectionsEventSubscription('Disconnect-Endpoints') # Method from connections.psm1
    Start-UI
}

function Start-Tasks {
	$modules = $global:PLUGINS_MODULES | Where-Object {$_.checked -eq $true}
    if(Connect-Endpoints $modules){
		$global:PLUGINS_MODULES | Show-Menu

		foreach ($module in $modules) {
			$params = $global:connections | Where-Object {$_.component -in $module.COMPONENT.Split(";") }
			
			Write-Title($module.Name)
			&($module.Name) $params
		}
	}else{
		Write-Host "Could not connect to all endpoints!"
	}
}