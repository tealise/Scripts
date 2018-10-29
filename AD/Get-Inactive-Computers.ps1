<#
.SYNOPSIS
    Export inactive computers older than 30 days
.DESCRIPTION
   Look at AD and pull a CSV of all computer accounts dating inactive
   for greater than 30 days
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

# Retrieve the users based on lastLogonTimestamp less than variable
Get-ADComputer -Properties * -Filter {LastLogonTimeStamp -lt $time -and enabled -eq $true} |

# Output to CSV
Select-Object @{Label = "Hostname";Expression = {$_.cn}},
              @{Label = "Description";Expression = {$_.description}},
              @{Label = "Location";Expression = {$_.distinguishedName}},
              @{Label = "Last Logon Time";Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} |

Export-CSV $path\ADExport_Inactive_Computers_$reportdate.csv -notypeinformation

Write-Host Exported CSV available at $path\ADExport_Inactive_Computers_$reportdate.csv
