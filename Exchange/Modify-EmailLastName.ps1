<#
.SYNOPSIS
	Adds SMTP alias to user mailbox and changes primary
.DESCRIPTION
	Script to quickly add an SMTP alias to a user's mailbox and change primary
	for name changes, i.e. marriage, divorce, etc.
.PARAMETER UserName
	sAMAccountName of user
.PARAMETER Alias
	The alias to append to the user's mailbox
.PARAMETER Domain
	Domain name (e.g. domain.local)
.PARAMETER DomainController
	Preferred Domain Controller
.PARAMETER ExchangeServer
	Mando argument, need exchange server address
.PARAMETER WinCredential
	Take an argument of (Get-Credential) output
.EXAMPLE
	C:\PS> Modify-EmailLastName -UserName "jsmith1" -Alias "j.smith"
#>

[CmdletBinding()]
Param (
	[Parameter(Mandatory=$true)]
	[string]$UserName, # AD Username

	[Parameter(Mandatory=$true)]
	[string]$Alias, # Added alias

	[Parameter(Mandatory=$true)]
	[string]$Domain, # AD Domain for User search space

	[Parameter(Mandatory=$false)]
	[string]$DomainController, # AD Domain Controller that's preferred

	[Parameter(Mandatory=$true)]
	[string]$ExchangeServer, # Must know where the script is talking to

	[Parameter(Mandatory=$false)]
	[object]$WinCredential # Get AD Credential for auth to Exchange box
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

	# Function to pass to Exchange PSSession for adding the alias
	$AddEmailAlias = Function AddEmailAlias($User,$Alias,$Domain) {
		# @TODO: Verify not already an alias nor primary

		$mb = Get-Mailbox -Identity $User
		$mb.EmailAddresses += "$Alias@$Domain"
		$mb | Set-Mailbox -EmailAddressPolicyEnabled $false
		$mb | Set-Mailbox -EmailAddresses $mb.EmailAddresses
		$mb | Set-Mailbox -PrimarySmtpAddress "$Alias@$Domain"
		$mb | Set-Mailbox -Alias $Alias
	}

	# Function to check user's existence
	Function UserLookup($User) {
		if (!(Get-ADUser -LDAPFilter "(sAMAccountName=$User)" -Server $DomainController)) {
			return $false
		} else {
			return $true
		}
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

	if(UserLookup($UserName)) {
		Write-Output "User '$($UserName)' exists."
		Write-Output "Adding $($Alias) to $($UserName)"
		Invoke-Command -Session $Session -ScriptBlock ${function:AddEmailAlias} -ArgumentList $UserName,$Alias,$Domain
	} else {
		Write-Warning "Username not found, unable to process update"
	}

	# Close out the Exchange PSSession
	Remove-PSSession $Session

}
