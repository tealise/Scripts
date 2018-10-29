<# 
.SYNOPSIS
    Export inactive users older than 30 days
.DESCRIPTION
   Filtering out service accounts, this will look at AD and pull a
   CSV of all user accounts dating inactive for greater than 30 days
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
$OUName = "Service Accounts"
$path = Split-Path -parent "\\SERVER01\Share\ADRecon\Output\*.*"

# Creating working directory
$pathexist = Test-Path -Path $path
If ($pathexist -eq $false)
    {New-Item -type directory -Path $path}

$reportdate = Get-Date -Format yyyyMMdd

# Retrieve the users based on lastLogonTimestamp less than variable
Get-ADUser -Properties * -Filter {LastLogonTimeStamp -lt $time -and enabled -eq $true} | ? {$_.DistinguishedName -notlike "*,OU=Service Accounts,*"} |

# Output to CSV
Select-Object @{Label = "Display Name";Expression = {$_.DisplayName}},
              @{Label = "Logon Name";Expression = {$_.sAMAccountName}},
              @{Label = "Description";Expression = {$_.Description}},
              @{Label = "Location";Expression = {$_.distinguishedName}},
              @{Label = "Last Logon Time";Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} |

Export-CSV $path\ADExport_Inactive_Users_$reportdate.csv -notypeinformation

Write-Host Exported CSV available at $path\ADExport_Inactive_Users_$reportdate.csv
