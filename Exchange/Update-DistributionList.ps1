<#
.SYNOPSIS
	Creates a new mailbox for a user with quota and retention policy
.DESCRIPTION
	Script to quickly provision a new mailbox for an existing AD User and set quota and retention
.PARAMETER UserList
	Filepath of userlist.csv file (e.g. C:\Temp\userlist.csv)
	UserList.csv must contain two columns, DL and UserID
.PARAMETER Domain
	Domain name (e.g. domain.local)
.PARAMETER DomainController
	Preferred Domain Controller
.PARAMETER ExchangeServer
	Mando argument, need exchange server address
.PARAMETER WinCredential
	Take an argument of (Get-Credential) output
.EXAMPLE
	C:\PS> Create-UserMailbox -User person1 -RetPol "Default" -Quota "2gb"
	< Provisions mailbox for user "person1" with the "Default" retention policy and 2gb limit >
#>

[CmdletBinding()]
Param (
	[Parameter(Mandatory=$true)]
	[string]$UserList, # full filepath and name for userlist

	[Parameter(Mandatory=$true)]
	[string]$Domain, # AD Domain for User search space

	[Parameter(Mandatory=$false)]
	[string]$DomainController, # AD Domain Controller that's preferred

	[Parameter(Mandatory=$false)]
	[object]$WinCredential, # Get AD Credential for auth to Exchange box

	[Parameter(Mandatory=$true)]
	[string]$ExchangeServer # Must know where the script is talking to
)

process {

	# Check if required modules exist
	if (Get-Module -ListAvailable -Name ActiveDirectory) {
		Import-Module ActiveDirectory
	} else {
		$NopeError = [string]"ActiveDirectory powershell module is not installed. Please install RSAT and enable the AD PS module."
		Write-Warning $NopeError
		exit
	}

	# Check Preferred DC
	if (-Not $DomainController) { $DomainController = $env:LogonServer -replace '\\','' }

	# Function to check if mailbox exists
	Function DLLookup($DL) {
		if (!(Get-ADGroup -LDAPFilter "(mailNickname=$DL)" -Server $DomainController)) {
			return $false
		} else {
			return $true
		}
	}

	# Function to check user's existence
	Function UserLookup($User) {
		if (!(Get-ADUser -LDAPFilter "(sAMAccountName=$User)" -Server $DomainController)) {
			return $false
		} else {
			return $true
		}
	}

	$UpdateDL = Function UpdateDL($DL,$User) {
		Add-DistributionGroupMember -Identity $DL -Member $User -ErrorAction SilentlyContinue
	}

	# Try and establish connection to Exchange server
	try {
    $SessionParams = `
    @{
        ConfigurationName = "Microsoft.Exchange"
        ConnectionUri = "http://$ExchangeServer/PowerShell/"
        Credential = $WinCredential
    }
		$Session = New-PSSession @SessionParams
	} catch {
		$NopeError = [string]"Unable to connect to Microsoft Exchange Server. Check your entries."
		Write-Warning $NopeError
		exit
	}

	if (Test-Path -Path $UserList) {
		Write-Host "Updating distribution group members"
		Import-CSV $userlist | ForEach {
			if(!(DLLookup($_.DL)) -or !(UserLookup($_.UserID))) { Write-Warning "'$($_.DL)' or '$($_.UserID)' does not exist, skipping..." }
			else {
				$u = $_.UserID
				$d = $_.DL
				Write-Host "Adding $($u) to $($d)"
				Invoke-Command -Session $Session -ScriptBlock ${function:UpdateDL} -ArgumentList $d,$u
			}
		}
	} else {
		Write-Warning "CSV File not found, unable to process distribution list update"
	}

	# Close out the Exchange PSSession
	Remove-PSSession $Session
	Write-Host "Exited Exchange Session"
	Write-Host "Completed Task"
	Write-Host "======================================="

}
