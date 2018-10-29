# #############################################################################
# MICROSOFT EXCHANGE 2013 - MENU
# NAME: Run-Exchange.ps1
#
# AUTHOR: Joshua Nasiatka
# DATE:   2017/11/30
#
# COMMENT:  This script handles menu system to drive the other Exchange scripts
#
# VERSION HISTORY
# 1.0 2017.11.30 Initial Version.
#
# #############################################################################

############################### CONFIG SETTINGS ###############################
$myDir = $env:UserProfile + "\AppData\Local\SysadminPOSH\Exchange"
if (-Not (Test-Path -Path $myDir)) { New-Item -Path $myDir -ItemType Directory }
#$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

if (!(Test-Path -Path "$MyDir\Settings.xml")) {
  Write-Warning "Unable to find settings.xml file"
  $createone = Read-Host "Would you like to create a new settings file? [Y/N]"
  if ($createone -eq "Y") {
    [string]$DomainFQDN           = Read-Host "Enter Domain Name (e.g. contoso.local)"
    [string]$DC                   = Read-Host "Enter Domain Controller hostname nearest Exchange (e.g. DC02)"
    [string]$DistributionGroupOU  = Read-Host "Enter distinguished name for distribution group location (e.g. CN=Users,DC=contoso,DC=local)"
    [string]$SharedMailboxOU      = Read-Host "Enter distinguished name for shared mailbox location (e.g. CN=Groups,DC=contoso,DC=local)"
    [string]$ExchangeSVR          = Read-Host "Enter hostname of exchange server (e.g. MSExch01)"

    Write-Host "Generating XML File..."
    [xml]$ConfigFile = (Get-Content ".\SettingsTemplate.xml")
    $ConfigFile.Settings.Domain.FQDN                = $DomainFQDN
    $ConfigFile.Settings.Domain.PreferredDC         = $DC
    $ConfigFile.Settings.Other.DistributionGroupOU  = $DistributionGroupOU
    $ConfigFile.Settings.Other.SharedMailboxOU      = $SharedMailboxOU
    $ConfigFile.Settings.Exchange.ServerName        = $ExchangeSVR
    $ConfigFile.Settings.Exchange.SMTPPort          = "25"
    $ConfigFile.Save("$MyDir\Settings.xml")

    Write-Host "File Created in '$MyDir'"
    pause
  } else {
    Write-Warning "Need a settings.xml file in order to proceed`nQuitting..."
    pause
    exit
  }
}

[xml]$ConfigFile = Get-Content "$MyDir\Settings.xml"

$DomainFQDN           = $ConfigFile.Settings.Domain.FQDN                # e.g. contoso.local
$DC                   = $ConfigFile.Settings.Domain.PreferredDC         # e.g. dc01
$DistributionGroupOU  = $ConfigFile.Settings.Other.DistributionGroupOU  # e.g. CN=Users,DC=contoso,DC=local
$SharedMailboxOU      = $ConfigFile.Settings.Other.SharedMailboxOU      # e.g. CN=Users,DC=contoso,DC=local
$ExchangeSVR          = $ConfigFile.Settings.Exchange.ServerName        # e.g. msexch101
###############################################################################

Function ClearScreen { [System.Console]::Clear() }

ClearScreen
Write-Host "Waiting for credentials..."

$titlemenu = @"
############################################################
#                        Main Menu                         #
#                                                          #
#                                                          #
#              Exchange 2013 Inteface System               #
#                                                          #
#                  PSIS by Joshua Nasiatka                 #
#                                                          #
#                                                          #
############################################################
#                                                          #
#  0. Create new mailbox       |   4. Create distribution  #
#  1. Unprovision mailbox      |   5. Bulk add to DL       #
#  2. Remove mailbox and user  |                           #
#  3. Provision mailbox ACL    |  99. Exit                 #
#                                                          #
############################################################
"@

# Get AD Credential w/ permission to access print servers (e.g. Domain Admin account)
if($cred = $host.ui.PromptForCredential('SA Credentials Required', 'Please enter your systems admin (SA) credentials in order to poll Active Directory and Exchange.',
'', "")){} else {
    Write-Warning "Need systems admin (SA) credentials in order to proceed.`r`nPlease re-run script and enter the appropriate credentials."
    exit
}

$username = $cred.username
$password = $cred.GetNetworkCredential().password
$CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
$domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain,$UserName,$Password)
if ($domain.name -eq $null) {
    Write-Warning "Authentication failed - please verify your username and password."
    exit
} else {
    Write-Host "Successfully authenticated with domain $domain.name"
}

$BulkParameters = `
@{
  Domain = $DomainFQDN
  ExchangeServer = $ExchangeSVR
  DomainController = $DC
  WinCredential = $cred
}

ClearScreen

Function ObtainMenuDecision() {
  [string]$menu = Read-Host "Please enter number to execute action"
  switch($menu) {
    0 { ProcessNewMailbox }
    1 { ProcessUnprovisionMailbox }
    2 { ProcessRemoveMailbox }
    3 { ProcessMailboxACL }
    4 { ProcessCreateDL }
    5 { ProcessUpdateDL }
    99 { exit }
    default { Write-Warning "Menu Option Not Found`nExiting"; break }
  }
}

Function ProcessNewMailbox {
  Write-Host "`nStarting new mailbox workflow..."
  $who = Read-Host "Enter w# or Display Name"
  $alias = Read-Host "Enter email alias"
  $shared = Read-Host "Shared Mailbox? OR User? [S/U]"
  if ($shared -eq 'S') {
      ./Create-UserMailbox.ps1 -User $who -Alias $alias -Shared @BulkParameters -OU $SharedMailboxOU
  } else {
    ./Create-UserMailbox.ps1 -User $who -Alias $alias @BulkParameters
  }
}

Function ProcessUnprovisionMailbox {
  Write-Host "`nStarting unprovision mailbox workflow..."
  ./Remove-UserMailbox.ps1 @BulkParameters
}

Function ProcessRemoveMailbox {
  Write-Host "`nBeginning removal workflow for mailbox and user..."
  $opt = Read-Host "Are you sure you want to delete a mailbox and user. This will destroy the record from Active Directory. [DELETE/NO]"
  if ($opt -eq "DELETE") {
    ClearScreen
    ./Remove-UserMailbox.ps1 -Delete -Force @BulkParameters
  } else {
    ClearScreen
  }
}

Function ProcessMailboxACL {
  Write-Host "`nBeginning mailbox ACL change workflow...`n"
  $dobulk = Read-Host "Will you be performing a bulk change? [Y/N]"

  if ($dobulk -eq "Y") {
    Write-Host "Workflow changed to bulk transaction`n"
    $filename = Read-Host "Enter filepath (e.g. C:\Temp\ExchangeThings.csv)"
    ./Set-MailboxPerms.ps1 -Bulk -Filepath $filename @BulkParameters
  } else {
    Write-Host "Workflow changed to single user modification`n"
    $user = Read-Host "Enter username of person being given access"
    $box = Read-Host "Enter mailbox alias (e.g. custsupport or j.doe)"
    ./Set-MailboxPerms.ps1 -User $user -MailboxAlias $box @BulkParameters
  }
}

Function ProcessCreateDL {
  Write-Host "`nBeginning new distribution list workflow...`n"
  $DisplayName = Read-Host "Enter Display Name"
  $Alias = Read-Host "Enter alias, aka stuff before '@'"
  $HaveFile = Read-Host "Did you create a new CSV w/ UserID and DL columns? [Y/N]"

  if ($HaveFile -eq "Y") {
    $UserListLocation = Read-Host "Enter filepath"
    ./Create-DistributionList.ps1 -DisplayName $DisplayName -Alias $Alias -UserList $UserListLocation -Domain $DomainFQDN -DomainController $DC -OU $DistributionGroupOU -ExchangeServer $ExchangeSVR -WinCredential $cred
  } else {
    Write-Host "`nChanging workflow to 'create only'..."
    ./Create-DistributionList.ps1 -DisplayName $DisplayName -Alias $Alias -Domain $DomainFQDN -DomainController $DC -OU $DistributionGroupOU -CreateOnly -ExchangeServer $ExchangeSVR -WinCredential $cred
  }
}

Function ProcessUpdateDL {
  Write-Host "`nBegnning update distribution list workflow...`n"
  $HaveFile = Read-Host "Did you create a new CSV w/ UserID and DL columns? [Y/N]"

  if ($HaveFile -eq "Y") {
    $UserListLocation = Read-Host "Enter filepath"
    ./Update-DistributionList.ps1 -UserList $UserListLocation -Domain $DomainFQDN -DomainController $DC -ExchangeServer $ExchangeSVR -WinCredential $cred
  } else {
    Write-Warning "Not sure how to update distribution lists if you don't give me a file to work with..."
  }
}

do { pause; ClearScreen; Write-Host $titlemenu; ObtainMenuDecision } while (1)
