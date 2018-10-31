# #############################################################################
# ONBOARDING - WORKFLOW
# NAME: Create-NewHire.ps1
#
# AUTHOR: Joshua Nasiatka
# DATE:   2018/09/05
# EMAIL:  dev@joshuanasiatka.com
#
# COMMENT:  This script will create AD users w/w/o Exchange mailboxes both
#           manually and in bulk via CSV
#
# VERSION HISTORY
# 1.0 2018.09.05 Initial Version.
#
# #############################################################################

param (
  [string]$DC
)

############################### CONFIG SETTINGS #################################
# $templatePath 	= "C:\temp\UserDocsTemplate"	# Dummy folder created only to configure desired permissions (for the user folders)
$userDocsPath 	= "\\SERVER01\UserDocs"		# User Docs directory (e.g. \\SERVER\UserDocs)
$dc = "DC01" # Domain Controller (Remote/Same Subnet as Exchange)
$exchange = "EXCH01" # Exchange Server
$domain_fqdn = "example.com"
$domain_canonical = "dc=example,dc=com"
#################################################################################

# LOAD FUNCTIONS FROM FILE
. .\fn-NewHire.ps1

Clear-Host
Write-Output @"
############################################################
#                        Main Menu                         #
#                                                          #
#                                                          #
#           Employee Onboarding Inteface System            #
#                                                          #
#                 PSIS by Joshua Nasiatka                  #
#                                                          #
#                                                          #
############################################################
#                                                          #
#  0. Onboard New Hire         |   4.                      #
#  1.                          |   5.                      #
#  2.                          |                           #
#  3.                          |  99. Exit                 #
#                                                          #
############################################################
"@

Write-Output "Defaulting to menu option 1.`r`n"

Import-Module ActiveDirectory

$locations = Import-CSV .\locations.csv # BRANCH LOCATIONS
$roles = (Import-CSV .\roles.csv) | ?{$_.Location -ne "" -and $_.Department -ne ""} # ROLES AND DEPARTMENTS

$errors = @()   # EMPTY ERROR LOG
$warnings = @() # EMPTY WARNING LOG

# Obtain user information
$first_name = Read-Host "Enter First Name                    "
$last_name = Read-Host  "Enter Last Name                     "
$email_alias = $($first_name[0]).$last_name
$extension = ValidateExtension
$next_available_username = GetUserName
$role = RolePicker($roles)
$location = $locations | ?{$_.Name -eq $role.Location} | Select -First 1

ShowSummary

### TODO: ADD VERIFICATION STEP
$confirm = Read-Host "Does this information look correct? [Y/N]"

### UPLOAD USER TO AD ###
$password = ConvertTo-SecureString "P@ssw0rd1!" -AsPlainText -Force

# Read in data from template user
$template = Get-ADUser -Identity $(($role).Template) -Properties memberOf, PrimaryGroup, LogonHours, ScriptPath,
  Title, Department, Manager, Company, StreetAddress, City, State, PostalCode

$user_profile_data = @{
  'GivenName' = $first_name
  'Surname' = $last_name
  'Description' = $role.Description
  'Path' = $role.OU
  'SamAccountName' = $next_available_username
  'UserPrincipalName' = $next_available_username
  'Name' = "$($first_name) $($last_name)"
  'DisplayName' = "$($first_name) $($last_name)"
  'AccountPassword' = $password
  'ChangePasswordAtLogon' = $true
  'Enabled' = $true
}

Write-Output "Creating new AD account for $first_name $last_name with username $next_available_username"
New-ADUser @user_profile_data -Instance $template

### VERIFY UPLOAD TO AD ###
if ($extension) {
  # Check if extension is in valid DID region based on those configured in PBX system
  $telephone = $(if(($extension -le 1299) -and ($extension -ge 1100)) { "508-555-$extension" } elseif(($extension -le 3699) -and ($extension -ge 3500)){ "807-555-$extension" })
  Set-ADUser -Identity $next_available_username -add @{ipphone="$extension"}
  Set-ADUser -Identity $next_available_username -add @{telephoneNumber="$telephone"}
}

Set-ADUser -Identity $next_available_username -Manager $role.Manager
Set-ADUser -Identity $next_available_username -Department $role.Department

# Attach memberships
$template.memberOf | Add-ADGroupMember -Members $next_available_username
if ($role.AdditionalGroups) {($role.AdditionalGroups).Split(',') | Add-ADGroupMember -Members $next_available_username }
$newuser = Get-ADUser -Identity $next_available_username

Write-Output "Creating UserDocs Folder"
CreateUserDocs $userDocsPath $next_available_username

Write-Output "Configuring Mailbox."
$BulkParameters = `
@{
    Domain = $domain_fqdn
    ExchangeServer = $exchange
    DomainController = $dc
    WinCredential = $cred
}
./Create-UserMailbox.ps1 -User $next_available_username -Alias $email_alias @BulkParameters

Write-Output "Completed."
. .\fn-NewHire.ps1
ShowSummary
