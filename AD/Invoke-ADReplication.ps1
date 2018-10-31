<#
.SYNOPSIS
    Runs a "replicate now" against Domain Controllers
.DESCRIPTION
    Forces replication to start on all DCs right now
.EXAMPLE
    C:\PS> Invoke-ADReplication
    < DCs replicate now >
#>

# Make sure that the Quest AD SnapIns are installed
if ( (Get-PSSnapin -Name Quest.ActiveRoles.ADManagement -ErrorAction silentlycontinue) -eq $null ) {
    # The QAD snapin is not active. Check it's installed
    if ( (Get-PSSnapin -Name Quest.ActiveRoles.ADManagement -Registered -ErrorAction SilentlyContinue) -eq $null) {
      Write-Error "You must install Quest ActiveRoles AD Tools to use this script!"
    } else {
      Write-Output "Importing QAD Tools"
      Add-PSSnapin -Name Quest.ActiveRoles.ADManagement -ErrorAction Stop
    }
}

Write-Output "Starting replication on ADDS"
# Find each domain controller, then do a foreach-object
Get-QADComputer -ComputerRole 'DomainController' | % {
    Write-Output "Replicating $($_.Name)"
    (repadmin /kcc $_.Name) | Out-null
    (repadmin /syncall /A /e $_.Name) | out-null
}

Write-Output "Completed ADDS Replication"
