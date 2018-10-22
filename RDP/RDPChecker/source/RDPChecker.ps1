<#
==================================
--- RDP Checker Tool
==================================
  Author:  Joshua Nasiatka
  Date:    Mar 2017
  Version: 1.0.0
==================================
#>

# Set-ExecutionPolicy RemoteSigned

# Set Paths Here
$wdpath = "."
$icofile = "rdp2.ico"
##################################

function ShowNotify {
    [cmdletbinding()]
        param(
        [parameter(Mandatory=$true)]
        [string]$Title,
        [ValidateSet("Info","Warning","Error")]
        [string]$MessageType = "Info",
        [parameter(Mandatory=$true)]
        [string]$Message,
        [string]$Duration=10000
    )

    [system.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
    $balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = Get-Process -id $pid | Select-Object -ExpandProperty Path
    #$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
    $balloon.Icon = "$wdpath\$icofile"
    $balloon.BalloonTipIcon = $MessageType
    $balloon.BalloonTipText = $Message
    $balloon.BalloonTipTitle = $Title
    $balloon.Visible = $true
    $balloon.ShowBalloonTip($Duration)
}

function checkPing {
    [cmdletbinding()]
        param(
        [parameter(Mandatory=$true)]
        [string]$CheckHost,
        [parameter(Mandatory=$true)]
        [ValidateSet("Drop","Alive")]
        [string]$CheckFor
    )

    if ($CheckFor -eq "Drop") {
        do {} until (!(Test-Connection $CheckHost -quiet))
        ShowNotify -Title "$CheckHost has gone down." -MessageType Error -Message "System has gone offline" -Duration 1000
    } elseif ($CheckFor -eq "Alive") {
        do {} until (Test-Connection $CheckHost -quiet)
        ShowNotify -Title "$CheckHost is back up." -MessageType Error -Message "System is responding to pings" -Duration 1000
    }

}

function waitForService {
    [cmdletbinding()]
        param(
        [parameter(Mandatory=$true)]
        [string]$CheckHost,
        [parameter(Mandatory=$true)]
        [int]$CheckPort
    )
    do {} until (TNC -ComputerName $CheckHost -Port 3389 -InformationLevel Quiet)
    ShowNotify -Title "$CheckHost is ready" -MessageType Info -Message "System is online and ready for RDP" -Duration 1500
}

function RebootRDPChecker {
    [cmdletbinding()]
        param(
        [parameter(Mandatory=$true)]
        [string]$CheckHost
    )
    
    ShowNotify -Title "Watching $CheckHost for a reboot" -MessageType Info -Message "Waiting for next signal" -Duration 1000
    checkPing -CheckHost $CheckHost -CheckFor Drop
    checkPing -CheckHost $CheckHost -CheckFor Alive
    waitForService -CheckHost $CheckHost -CheckPort 3389
}

function Btn_Click {
    if ($CheckHost_tb.Text-eq "") {
        $RebootRDPChecker.controls.Add($warn_lb)
    } else {
        $RebootRDPChecker.Hide()
        RebootRDPChecker -CheckHost $CheckHost_tb.Text
        if($RDP.checked -eq $true) {
            mstsc /v:$checkHost_tb.Text
        }
    }
}

function ShowForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Windows.Forms

    $RebootRDPChecker = New-Object system.Windows.Forms.Form
    $RebootRDPChecker.Text = "Reboot RDP Checker"
    $RebootRDPChecker.Icon = New-Object system.drawing.icon("$wdpath\$icofile")
    $RebootRDPChecker.Width = 450
    $RebootRDPChecker.Height = 200

    $CheckHost_tb = New-Object system.windows.Forms.TextBox
    $CheckHost_tb.Width = 385
    $CheckHost_tb.Height = 20
    $CheckHost_tb.location = new-object system.drawing.point(20,45)
    $CheckHost_tb.Font = "Consolas,12"
    $RebootRDPChecker.controls.Add($CheckHost_tb)

    $CheckHost_lb = New-Object system.windows.Forms.Label
    $CheckHost_lb.Text = "Hostname / IP Address:"
    $CheckHost_lb.AutoSize = $true
    $CheckHost_lb.Width = 25
    $CheckHost_lb.Height = 10
    $CheckHost_lb.location = new-object system.drawing.point(20,20)
    $CheckHost_lb.Font = "Microsoft Sans Serif,12,style=Bold"
    $RebootRDPChecker.controls.Add($CheckHost_lb)

    $RDP = New-Object system.windows.Forms.CheckBox
    $RDP.Text = "Automatically RDP?"
    $RDP.AutoSize = $true
    $RDP.Width = 95
    $RDP.Height = 20
    $RDP.location = new-object system.drawing.point(20,120)
    $RDP.Font = "Microsoft Sans Serif,12"
    $RebootRDPChecker.controls.Add($RDP)

    $submit_btn = New-Object system.windows.Forms.Button
    $submit_btn.Text = "Watch"
    $submit_btn.Width = 60
    $submit_btn.Height = 30
    $submit_btn.location = new-object system.drawing.point(342,110)
    $submit_btn.Font = "Microsoft Sans Serif,10"
    $RebootRDPChecker.controls.Add($submit_btn)

    $warn_lb = New-Object system.windows.Forms.Label
    $warn_lb.Text = "No! A hostname or IP address must be given."
    $warn_lb.AutoSize = $true
    $warn_lb.ForeColor = "#db2111"
    $warn_lb.Width = 25
    $warn_lb.Height = 10
    $warn_lb.location = new-object system.drawing.point(20,75)
    $warn_lb.Font = "Microsoft Sans Serif,10,style=Bold"

    $submit_btn.Add_Click({Btn_Click})

    $RebootRDPChecker.ShowDialog()
}

ShowForm
