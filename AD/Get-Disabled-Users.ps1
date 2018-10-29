<#
.SYNOPSIS
    Export disabled user accounts older than 30 days
.DESCRIPTION
   Look at AD and pull a CSV of all disabled user accounts dating
   inactive for greater than 30 days
.NOTES
    Author     : Joshua Nasiatka
.LINK
    https://github.com/joshuanasiatka/
#>

# VERIFY AD MODULE DEPENDENCY IS INSTALLED
if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "ERROR!!" -foregroundcolor "magenta"
    Write-Host "Active Directory CMDLET Not Installed:"
    Write-Host "Please install RSAT and enable Active Directory Powershell Module"
    exit 1
} else {
    Import-Module ActiveDirectory
}

$domain = "domain.com"
$DaysInactive = 30
$time = (Get-Date).AddDays(-($DaysInactive))
$path = Split-Path -parent "\\SERVER01\Share\ADRecon\Output\*.*"

# Creating working directory
$pathexist = Test-Path -Path $path
If ($pathexist -eq $false)
    {New-Item -type directory -Path $path}

$reportdate = Get-Date -Format yyyyMMdd

# Place the following before "Enabled -eq $false" to check accounts that have not logged on specifically
# within the last 30 days, otherwise all accounts including never logged on ones
# LastLogonTimeStamp -lt $time -and

# Retrieve the users based on lastLogonTimestamp less than variable
Get-ADUser -Properties * -Filter {Enabled -eq $false} -searchbase "OU=Disabled Accounts,dc=domain,dc=com" |
Select-Object @{Label = "Display Name";Expression = {$_.DisplayName}},
              @{Label = "Logon Name";Expression = {$_.sAMAccountName}},
              @{Label = "Description";Expression = {$_.Description}},
              @{Label = "Last Logon Time";Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} |

Export-CSV $path\ADExport_Disabled_Users_$reportdate.csv -notypeinformation

Write-Host Exported CSV available at $path\ADExport_Disabled_Users_$reportdate.csv
