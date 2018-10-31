<#
.SYNOPSIS
	Removes mailbox for a user
.DESCRIPTION
	Script to quickly unprovision a mailbox for an existing AD User
.PARAMETER User
	The user to whom the mailbox will be provisioned
.PARAMETER Domain
	Domain name (e.g. domain.local)
.PARAMETER DomainController
	Preferred Domain Controller
.PARAMETER Delete
	Switch to delete mailbox and user
.PARAMETER Force
	Switch to force delete, requires -delete
.PARAMETER WinCredential
	Take an argument of (Get-Credential) output
.PARAMETER ExchangeServer
	Hostname or fqdn of Exchange server
.EXAMPLE
C:\PS> Remove-UserMailbox -User person1
< Unprovisions mailbox for user "person1" >
#>

[CmdletBinding()]
Param (
	[Parameter(Mandatory=$true)]
	[string]$User, # The domain user

	[Parameter(Mandatory=$true)]
	[string]$Domain, # AD Domain for User search space

	[Parameter(Mandatory=$false)]
	[string]$DomainController = $env:LogonServer -replace '\\','', # AD Domain Controller that's preferred

	[Parameter(Mandatory=$false)]
	[switch]$Delete, # Delete mailbox and user

	[Parameter(Mandatory=$false)]
	[switch]$Force, # Purges from mailbox database

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
		return
	}

	# Let's make sure the user actually exists first
	if (!(Get-ADUser -LDAPFilter "(sAMAccountName=$User)" -Server $DomainController)) {
		$NopeError = [string]"Can't find user with the account '$User'"
		Write-Warning $NopeError
		return
	} else {
		Write-Output "Found a record for user '$User'"
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
		return
	}

	Write-Output "Processing Exchange account for user, $User ($Alias)"

	$out = `
	Invoke-Command -Session $Session -ScriptBlock {
		# Accept the parameters
		param ($User,$Domain,$Delete,$Force)

		if ($Delete) {
			if ($Force) {
				Remove-Mailbox -Identity "$Domain\$User" -Permanent $true
			} else {
				Remove-Mailbox -Identity "$Domain\$User"
			}
		} else {
			Disable-Mailbox -Identity "$Domain\$User" -Confirm:$false
		}
	} -ArgumentList $User,$Domain,$Delete,$Force # Pass in the variables

	# Close out the Exchange PSSession
	Remove-PSSession $Session
	Write-Output "Exited Exchange Session"
	Write-Output "Completed Task"
	Write-Output "======================================="

}
