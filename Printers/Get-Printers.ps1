# #############################################################################
# ACTIVE DIRECTORY - GET LIST OF PRINTERS
# NAME: Get-Printers.ps1
#
# AUTHOR: Joshua Nasiatka
# DATE:   2017/11/27
#
# COMMENT:  This script will retrieve a current list of all printers and their associated print server
#
# VERSION HISTORY
# 1.0 2017.11.27 Initial Version.
# 1.1 2017.11.27 Added Filter for only network printers; added credential check
#
# #############################################################################

# Get AD Credential w/ permission to access print servers (e.g. Domain Admin account)
if($cred = $host.ui.PromptForCredential('SA Credentials Required', 'Please enter your systems admin (SA) credentials in order to poll Active Directory and Print Servers.',
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
    Write-Output "Successfully authenticated with domain $domain.name"
}

# Generate array of print servers
Import-Module ActiveDirectory
$printServers = Get-ADObject -LDAPFilter "(&(&(&(uncName=*)(objectCategory=printQueue))))" -properties *|Sort-Object -Unique -Property servername |select servername

# Enumerate through and generate table
$printers = @()
$printServers | Foreach-Object {
    $printers += gwmi win32_printer -computer $_.servername -credential $cred  |  Select-Object PSComputerName,Caption,Comment,DriverName,Location,PortName,PrintProcessor | Where {$_.PortName -like "10.172*"}
}

# Export printer list to CSV
$printers | Export-CSV C:\Temp\SharedPrinters.csv -NoTypeInformation

# Open the new CSV
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$wb = $excel.Workbooks.Open("C:\Temp\SharedPrinters.csv")
