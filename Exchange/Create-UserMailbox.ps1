<#
.SYNOPSIS
	Creates a new mailbox for a user with quota and retention policy
.DESCRIPTION
	Script to quickly provision a new mailbox for an existing AD User and set quota and retention
.PARAMETER User
	The user to whom the mailbox will be provisioned
.PARAMETER Alias
	The alias of the user, the stuff before the '@' symbol
.PARAMETER Domain
	Domain name (e.g. domain.local)
.PARAMETER DomainController
	Preferred Domain Controller
.PARAMETER OU
	Organizational Unit for the shared mailbox to be created
.PARAMETER RetPol
	Mando argument for retention policy
.PARAMETER quota
	Mando argument for quota
.PARAMETER Shared
	Switch to toggle if a shared mailbox
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
	[string]$User, # The domain user

	[Parameter(Mandatory=$true)]
	[string]$Alias, # The email alias

	[Parameter(Mandatory=$true)]
	[string]$Domain, # AD Domain for User search space

	[Parameter(Mandatory=$false)]
	[string]$DomainController, # AD Domain Controller that's preferred

	[Parameter(Mandatory=$false)]
	[string]$OU, # AD Domain OU that's preferred

	[Parameter(Mandatory=$false)]
	[object]$WinCredential, # Get AD Credential for auth to Exchange box

	[Parameter(Mandatory=$false)]
	[string]$RetPol = "Default Policy", # Specify the retention policy

	[Parameter(Mandatory=$false)]
	[string]$Quota = "2GB", # Mailbox size limit

	[Parameter(Mandatory=$false)]
	[switch]$Shared = $false, # Shared mailbox toggle

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

    # Get Default Domain Controller
    if (!($DomainController)) { $DomainController = $env:LogonServer -replace '\\','' }

	# Let's make sure the user actually exists first
	if (!($Shared)) {
		if (!(Get-ADUser -LDAPFilter "(sAMAccountName=$User)" -Server $DomainController)) {
			$NopeError = [string]"Can't find user with the account '$User'"
			Write-Warning $NopeError
			exit
		} else {
			Write-Output "Found a record for user '$User'"
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

  Write-Output "Processing Exchange account for user, $User ($Alias)"

	$ExchangeCommands = Function ExchangeCommands($User,$Domain,$Alias,$RetPol,$ExchangeParams,$Shared,$OU) {
		if ($Shared) {
			New-Mailbox -Name $User -Alias $Alias -Shared -OrganizationalUnit $OU
			Set-Mailbox $Alias -RetentionPolicy $RetPol @ExchangeParams
		} else {
			Enable-Mailbox -Identity "$Domain\$User" -Alias $Alias
			Set-Mailbox -Identity "$Domain\$User" -RetentionPolicy $RetPol @ExchangeParams
		}
	}

	if ($Quota -eq "2GB") {
		$ExchangeParams = `
		@{
			IssueWarningQuota = "1.9gb"
			ProhibitSendQuota = "2.0gb"
			ProhibitSendReceiveQuota = "2.1gb"
			UseDatabaseQuotaDefaults = $false
		}
	} else {
		$ExchangeParams = ''
	}

	if (($Quota -eq "2GB") -Or ($Quota -eq "Unlimited")) {
		Write-Output "Creating account and setting the mailbox size and retention policy"
		if ($Shared) {
			Invoke-Command -Session $Session -ScriptBlock ${function:ExchangeCommands} -ArgumentList $User,$Domain,$Alias,$RetPol,$ExchangeParams,$Shared,$OU

			Set-ADUser $Alias -Description "$Alias@$Domain Shared Mailbox" -Server $DomainController
		} else {
			$OU = ''
			Invoke-Command -Session $Session -ScriptBlock ${function:ExchangeCommands} -ArgumentList $User,$Domain,$Alias,$RetPol,$ExchangeParams,$Shared,$OU
		}
	} else {
		$QuotaError = [string]"I am unfamiliar with that mailbox quota size supplied"
		Write-Warning $QuotaError
		exit
	}

	# Close out the Exchange PSSession
	Remove-PSSession $Session

}
