# #############################################################################
# ACTIVE DIRECTORY - GET CURRENT USER LIST
# NAME: Get-ADUser-List.ps1
#
# AUTHOR: Joshua Nasiatka
# DATE:   2016/12/27
#
# COMMENT:  This script will retrieve a currently list of all users and w302 numbers
#
# VERSION HISTORY
# 1.0 2016.12.27 Initial Version.
# 1.1 2017.05.10 CSV formatting and more information
# 1.2 2017.06.22 Formatting
#
# FEATURES
# -Get-ADUser Properties
# #############################################################################

# VERIFY AD MODULE DEPENDENCY IS INSTALLED
if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "ERROR!!" -foregroundcolor "magenta"
    Write-Host "Active Directory CMDLET Not Installed:"
    Write-Host "Please install RSAT and enable Active Directory Powershell Module"
    exit 1
} else {
    Import-Module ActiveDirectory
}

$path = Split-Path -parent "\\SERVER01\Share\ADRecon\Output\*.*" 
$pathexist = Test-Path -Path $path
If ($pathexist -eq $false)
    {New-Item -type directory -Path $path}

$reportdate = Get-Date -Format yyyyMMdd

$csvreportfile = $path + "\ADExport_$reportdate.csv"

# Perform AD search. The quotes "" used in $SearchLoc is essential
# Without it, Export-ADUsers returuned error
Get-ADUser -Properties * -Filter * |
Select-Object @{Label = "First Name";Expression = {$_.GivenName}},
              @{Label = "Last Name";Expression = {$_.Surname}},
              #@{Label = "Display Name";Expression = {$_.DisplayName}},
              @{Label = "Logon Name";Expression = {$_.sAMAccountName}},
              #@{Label = "Full address";Expression = {$_.StreetAddress}},
              #@{Label = "City";Expression = {$_.City}},
              #@{Label = "State";Expression = {$_.st}},
              #@{Label = "Post Code";Expression = {$_.PostalCode}},
              @{Label = "Job Title";Expression = {$_.Title}},
              @{Label = "Description";Expression = {$_.Description}},
              @{Label = "Department";Expression = {$_.Department}},
              @{Label = "Phone";Expression = {$_.telephoneNumber}},
              @{Label = "Email";Expression = {$_.Mail}} |

#Export CSV report
Export-Csv -Path $csvreportfile -NoTypeInformation

Write-Host "Exported CSV available at $csvreportfile"
