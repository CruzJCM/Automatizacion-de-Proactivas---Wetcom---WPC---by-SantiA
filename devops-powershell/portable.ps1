Push-Location (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)

Import-Module ..\Modules\VMware.VimAutomation.Cis.Core
Import-Module ..\Modules\VMware.VimAutomation.Common
Import-Module ..\Modules\VMware.VimAutomation.Core
Import-Module ..\Modules\VMware.VimAutomation.Vds
Import-Module ..\Modules\VMware.VimAutomation.Sdk
Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -confirm:$false
Set-PowerCLIConfiguration -InvalidCertificateAction:Ignore -confirm:$false

Import-Module ./app/app.psm1
Start-App

Pop-Location