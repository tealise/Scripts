<#
.SYNOPSIS
	Find/Replace departments/cost centers in AD
.DESCRIPTION
	Script to quickly Find/Replace departments/cost centers in AD
.PARAMETER InputFile
	File location of where to find the CSV for parsing
#>

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [string]$InputFile
)

############# VERIFY EXISTENCE OF CSV #########
If (!(Test-Path -Path $InputFile)) { Write-Error "You Dun Goofed! (No Valid Path)"; exit }
Else { Write-Output "File Exists!" }

############# OBTAIN CREDENTIALS ##############
# Get AD Credential w/ permission to access Active Directory (e.g. Domain Admin account)
if($cred = $host.ui.PromptForCredential('SA Credentials Required', 'Please enter your systems admin (SA) credentials in order to poll Active Directory and Exchange.',
'', "")){} else {
    Write-Warning "Need systems admin (SA) credentials in order to proceed.`r`nPlease re-run script and enter the appropriate credentials."
    exit
}

# Verify the entered credentials are valid
$username = $cred.username
$password = $cred.GetNetworkCredential().password
$CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
$domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain,$UserName,$Password)
if ($domain.name -eq $null) {
    Write-Warning "Authentication failed - please verify your username and password."
    exit
} else {
    Write-Output "Successfully authenticated with domain $domain.name"
}

############# PRIVATE FUNCTIONS ###############
function Verify-ADModule {
  # Search list of available modules for the AD module
  if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
      Write-Output "ERROR!!" -foregroundcolor "magenta"
      Write-Output "Active Directory CMDLET Not Installed:"
      Write-Output "Please install RSAT and enable Active Directory Powershell Module"
      return $false
  } else {
      Import-Module ActiveDirectory
      return $true
  }
}

function DepartmentPicker($dept) {
  switch -wildcard ($dept) {
    <# Complete the Switch case #>
    "*Accounting*" { "Finance" }
    "*Collections*" { "Credit Administration" }
    "*Credit*" { "Credit Administration" }
    "Retail*" { "Retail Administration" }

    default { "NO CHANGE" }
  }
}

############# BEGIN PROCESSING ################
$users = Import-Csv $InputFile
$users | Foreach-Object {
  $user = $_."Logon Name"
  $dept = $_.Department
  $newDept = DepartmentPicker($dept)
  if ($newDept -ne "NO CHANGE") {
    Set-ADUser -Identity $user -Department $newDept -Credential $cred
  } else {
    Write-Warning "No change for $user, no match for $dept"
  }
}

Write-Output "Department Update for All Users in AD is Completed`nRe-run .\Get-ADUser-List.ps1 or .\Get-ADUser-List-Wasp.ps1 to check."
