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
    Write-Output "ERROR!!" -foregroundcolor "magenta"
    Write-Output "Active Directory CMDLET Not Installed:"
    Write-Output "Please install RSAT and enable Active Directory Powershell Module"
    exit 1
} else {
    Import-Module ActiveDirectory
}

$path = Split-Path -parent "\\SEVER01\Share\ADRecon\Output\*.*"
$pathexist = Test-Path -Path $path
If ($pathexist -eq $false)
    {New-Item -type directory -Path $path}

$reportdate = Get-Date -Format yyyyMMdd

$csvreportfile = $path + "\ADExport_WASP_$reportdate.csv"

# Perform AD search. The quotes "" used in $SearchLoc is essential
# Without it, Export-ADUsers returuned error
Get-ADUser -Filter * -Properties * |
Select-Object @{Label = "Employee No";Expression = {$_.sAMAccountName}},
              @{Label = "Last Name";Expression = {$_.Surname}},
              @{Label = "First Name";Expression = {$_.GivenName}},
              @{Label = "Department";Expression = {$_.Department}},
              @{Label = "Title";Expression = {$_.Title}},
              @{Label = "Email";Expression = {$_.Mail}},
              @{Label = "Manager No";Expression = {(Get-ADUser $_.Manager).sAMAccountName}},
              @{Label = "Ext";Expression = {$_.IPPhone}},
              @{Label = "Address 1";Expression = {$_.streetAddress}},
              @{Label = "City";Expression = {$_.City}},
              @{Label = "State";Expression = {"MA"}},
              @{Label = "Postal Code";Expression = {$_.PostalCode}},
              @{Label = "Country";Expression = {"USA"}} |


#Export CSV report
Export-Csv -Path $csvreportfile -NoTypeInformation

Write-Output "Exported CSV available at $csvreportfile"
