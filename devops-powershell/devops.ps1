Push-Location (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)

Import-Module ./app/app.psm1
Start-App

Pop-Location