# #############################################################################
# CITRIX RECEIVER - INSTALL TOOL
# NAME: Deploy-CitrixReceiver.ps1
#
# AUTHOR: Joshua Nasiatka
# DATE:   2017/12/09
#
# COMMENT:  This script will upgrade/downgrade receiver to desired version
#
# VERSION HISTORY
# 1.0 2017.12.09 Initial Version.
#
# #############################################################################

$citrixPreferredVersion = "4.7.0.6553"
$citrixRegKey           = "HKLM:\SOFTWARE\Wow6432Node\Citrix\ReceiverInside\Registrations\"
$citrixInstalledVersion = "0.0"
$citrixPath             = "C:\Program Files (x86)\Citrix\ICA Client\Receiver"
$skipUninstall          = $false
$installerLocation      = "\\domain.com\netlogon\Install\CitrixReceiver"
$deployLocation         = "C:\ABReceiver"
$logshare               = "\\domain.com\RDSDirectory\ReceiverDeploy"
#$installerFlags         = "/silent /includeSSON"
$installerFlags         = "/silent /INCLUDESSON ADDLOCAL=ReceiverInside,ICA_Client,SSON,SELFSERVICE,WebHelper,AM,USB,DesktopViewer STORE0='Citrix Internal;https://citrix.domain.com/Citrix/pnagent/config.xml;on;Citrix Internal' ENABLE_SSON=yes"
<#
$installerFlags         = @(
  "/silent"
  "/INCLUDESSON"
  "ADDLOCAL=ReceiverInside,ICA_Client,SSON,SELFSERVICE,WebHelper,AM,USB,DesktopViewer"
  "STORE0=Citrix Internal;https://citrix.domain.com/Citrix/pnagent/config.xml;on;Citrix Internal"
  "ENABLE_SSON=yes"
)
#>
# DONOTSTARTCC=1
# Check if Citrix Receiver is already installed
if (Test-Path $citrixPath) {
  $citrixInstalledVersion = $(Get-ChildItem $citrixRegKey | ForEach-Object { Get-ItemProperty $_.pspath }).Version
} else {
  Write-Host "Citrix not currently installed"
  $skipUninstall = $true
}

if (Compare-Object $citrixPreferredVersion $citrixInstalledVersion) {
  # Copy over the installer and cleanup files
  Write-Host "Copying installer files to $deployLocation"
  if (!(Test-Path -Path $deployLocation)) { md $deployLocation }
  Copy-Item "$installerLocation\CitrixReceiver.exe" $deployLocation -Force
  Copy-Item "$installerLocation\ReceiverCleanupUtility.exe" $deployLocation -Force

  # If citrix is already installed but incorrect version, run the cleanup utility
  if (!($skipUninstall)) {
    Write-Host "Cleaning up old receiver data..."
    Start-Process -FilePath "$deployLocation\ReceiverCleanupUtility.exe" -ArgumentList "/silent" -wait
  }

  # Stop conflicting MsiExec process if exists
  Stop-Process -Name msiexec.exe -Force -ErrorAction SilentlyContinue

  # Begin installation
  Write-Host "Installing..."
  Start-Process $deployLocation\CitrixReceiver.exe -Wait -ArgumentList $installerFlags

  # Clean up Citrix temp files
  Write-Host "Cleaning up..."
  if (Test-Path -Path "%SYSTEMDRIVE%\Users\Default\AppData\Roaming\ICAClient") {
    Remove-Item -Recurse -Force "%SYSTEMDRIVE%\Users\Default\AppData\Roaming\ICAClient"
  }

} else {
  Write-Host "Citrix Receiver already at latest preferred version."
  exit
}
