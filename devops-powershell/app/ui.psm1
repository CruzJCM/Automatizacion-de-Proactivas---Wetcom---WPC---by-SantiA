$global:UI_LINEH = "═"
$global:UI_UPLEFT = "╔"
$global:UI_UPRIGHT = "╗"
$global:UI_LINEV = "║"
$global:UI_DOWNLEFT = "╚"
$global:UI_DOWNRIGHT = "╝"
$global:UI_LINEVLEFTCROSS = "╠"
$global:UI_LINEVRIGHTROSS = "╣"

$global:ExecuteEvent = @()
$global:ClearConnectionsEvent = @()


function Register-ExecuteEventSubscription($commandName) {
    $global:ExecuteEvent += $commandName
}

function Invoke-ExecuteEvent {
    foreach ($subscription in $global:ExecuteEvent) {
        &$subscription
    }
}

function Register-ClearConnectionsEventSubscription($commandName) {
    $global:ClearConnectionsEvent += $commandName
}

function Invoke-ClearConnectionsEvent {
    foreach ($subscription in $global:ClearConnectionsEvent) {
        &$subscription
    }
}

function Start-UI {
    while($true) {
        Start-Sleep -MilliSeconds 200
        if ($Host.UI.RawUI.KeyAvailable) {
			$keydown = $Host.UI.RawUI.ReadKey("IncludeKeyUp")
			$command = $keydown.Character
			if($keydown.VirtualKeyCode -ne 13){
				Write-Host -NoNewline "`b `b" # Solución de Sebas al echo aún con "NoEcho"
			}
        } 
        if ($command -eq "x") { 
            break 
        } elseif ($command -eq "s") { 
            "Selected options"
            Write-Host (($global:PLUGINS_MODULES | Where-Object {$_.checked -eq $true}).Synopsis  | Out-String)
            Read-Host "(press <Enter> to continue)"
            $global:PLUGINS_MODULES | Show-Menu
        } elseif ($command -Match "[1-9]") {
            $index = [convert]::ToInt32($command, 10)
			$index--
            $global:PLUGINS_MODULES[$index].checked = !$global:PLUGINS_MODULES[$index].checked
            $global:PLUGINS_MODULES | Show-Menu
        } elseif ($command -eq "e"){
            Invoke-ExecuteEvent
            Read-Host "(press <Enter> to continue)"
            $global:PLUGINS_MODULES | Show-Menu
        } elseif ($command -eq "c"){
            Invoke-ClearConnectionsEvent
            $global:PLUGINS_MODULES | ForEach-Object {$_.checked = $false}
            $global:PLUGINS_MODULES | Show-Menu
        } elseif($command -eq "p"){
            Submit-PendingReports
			Submit-Reports
			Read-Host "(press <Enter> to continue)"
			$global:PLUGINS_MODULES | Show-Menu
		}
        
        # Resize
        if ($global:ScreenSize -ne $host.UI.RawUI.BufferSize.Width) {
            $global:ScreenSize = $host.UI.RawUI.BufferSize.Width
            $global:PLUGINS_MODULES | Show-Menu
        }

        $command = $null
    }
}

function Show-Menu() 
{
    begin { 
        Clear-Host 
        $index = 1
        $width = $host.UI.RawUI.BufferSize.Width
        $global:UI_UPLEFT + ($global:UI_LINEH * ($width - 2)) + $global:UI_UPRIGHT

        Write-MenuLine ($global:CONFIG.APP_HEADER + " " + $global:APP_VERSION)
        Write-Separator
        foreach ($conn in $global:connections) {
            $component = $conn.component
            $host = $conn.host
            Write-MenuLine ($component + ": " + $host)
        }
        Write-Separator
    }
    process {
        $text = $_.Synopsis

        $checked = if ($_.checked) {"X"} else {" "}
        $option = "[$checked] (" + $index + ") "
        $index++

        Write-MenuLine ($option + $text)
    }
    end {
        Write-Separator
		Write-MenuLine "(s) Show selected | (c) Clear data"
		Write-MenuLine "(e) Execute | (p) Process reports | (x) Exit"
        $global:UI_DOWNLEFT + ($global:UI_LINEH * ($width - 2)) + $global:UI_DOWNRIGHT
    }
}

function Write-MenuLine($text, [switch]$alignCenter)
{
    $text = (" " + $text + (" " * $width)).Substring(0, $width -2)
    $global:UI_LINEV + $text + $global:UI_LINEV
}

function Write-Separator {
    $global:UI_LINEVLEFTCROSS + ($global:UI_LINEH * ($width - 2)) + $global:UI_LINEVRIGHTROSS
}

function Write-Title($text) {
	Write-Host " "
    "=" * $text.Length
    $text
    "=" * $text.Length
}

function Show-Progress($total, $index){
	$current = [math]::Round(($index / $total) * 100)
	$whitespace = 3 - $current.ToString().Length
	Write-Host (" " * $whitespace  + $current + "%") -NoNewline
    
    if ($current -lt 100) {
        Write-Host "`b`b`b`b" -NoNewline
    } else {
        Write-Host ""
    }
}