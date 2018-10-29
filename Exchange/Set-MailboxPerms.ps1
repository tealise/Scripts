<#
.SYNOPSIS
	Modify mailbox permissions
.DESCRIPTION
	Script to quickly give permissions to a user or users for a mailbox
.PARAMETER User
	The user to whom the mailbox will be provisioned
.PARAMETER MailboxAlias
	The alias of the user, the stuff before the '@' symbol
.PARAMETER Bulk
	Switch parameter for running script against a CSV
.PARAMETER Filepath
	Filepath of CSV file for bulk execution
.PARAMETER ExchangeServer
	Mando argument, need exchange server address
.PARAMETER Domain
	Domain name (e.g. domain.local)
.PARAMETER DomainController
	Preferred Domain Controller
.PARAMETER WinCredential
	Take an argument of (Get-Credential) output
.PARAMETER ExchangeServer
	Hostname or fqdn of Exchange server
.EXAMPLE
	C:\PS> .\Set-MailboxPerms -User person1 -MailboxAlias custsupport -Domain domain.local -ExchangeServer msexch101
	< Provisions mailbox perms for user "person1" with "FullAccess, Send on Behalf, and SendAs" for custsupport mailbox >
#>

[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false)]
	[string]$User, # The domain user

	[Parameter(Mandatory=$false)]
	[string]$MailboxAlias, # The email alias

	[Parameter(Mandatory=$false)]
	[switch]$Bulk, # Switch if bulk change

	[Parameter(Mandatory=$false)]
	[string]$Filepath, # Filepath for bulk execution

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
	Function MailboxLookup($Mailbox) {
		if ($Mailbox) {
			if (!(Get-ADUser -LDAPFilter "(mail=$Mailbox)" -Server $DomainController)) {
				$NopeError = [string]"Can't find mailbox with the address '$Mailbox'"
				Write-Warning $NopeError
				return $false
			} else {
				Write-Host "Found a record for mailbox '$Mailbox'"
				return $true
			}
		}
	}

	# If a MailboxAlias was supplied during script execution, verify existence
	if ($MailboxAlias) {
		if (!(MailboxLookup("$MailboxAlias@$Domain"))) { exit }
	}

	# If a filepath was given, verify existence
	if ($Filepath) {
		if (!(Test-Path $Filepath)) {
			Write-Warning "File '$Filepath' is not found."
			exit
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

	Write-Host "Connected to Microsoft Exchange"

	# This is the magic. Creates function to be passed into the Exchange PSSession
	#		which will handle the mailbox permissions
	$ExchangeCommands = Function ExchangeCommands($User,$box) {
		$boxinfo = Get-Recipient $box | Select-Object name
		$boxperson = $boxinfo.name

		Add-MailboxPermission $box -User $User -AccessRights FullAccess
		Set-Mailbox $box -GrantSendOnBehalfTo $User
		Get-Mailbox $box | Select-Object -ExpandProperty GrantSendOnBehalfTo
		Add-ADPermission -Identity "$boxperson" -User $User -AccessRights ExtendedRight -ExtendedRights "Send As"
	}

	# If -bulk specified, read host for csv path and filename
	#		then foreach line in the csv, call the magic function above with the UserID and MailboxAlias
	if ($Bulk) {
		if ($Filepath) {
			$csv = Import-CSV $Filepath
		} else {
			$workpath = Read-Host "Enter file path of CSV file (e.g. C:\Temp)"
			$filename = Read-Host "Enter filename (e.g. userpeople.csv)"
			$csv = Import-CSV $workpath\$filename
		}

		$csv | ForEach-Object -Process{
			$users = $_.UserID
			$name = @{Add="$users"}
			$MailboxName = $_.Mailbox

			if ($MailboxName -like "*@*") {
				$box = $MailboxName
			} else {
				$box = "$MailboxName@$Domain"
			}

			# Verifying mailbox exists before talking to Exchange
			if (MailboxLookup($box)) {
				Write-Host "Processing Exchange permissions for $users => ($MailboxName)"
				$out = Invoke-Command -Session $Session -ScriptBlock ${function:ExchangeCommands} -ArgumentList $_.UserID,$box
			}
		}

	} elseif ($User -And $MailboxAlias) {	# If during the script call, user and MailboxAlias were specified, process function
		Write-Host "Command received, determining how to execute... Please hold."

		if ($MailboxAlias -like "*@*") {
			$box = $MailboxAlias
		} else {
			$box = "$MailboxAlias@$Domain"
		}

		# Verifying mailbox exists before talking to Exchange
		if (MailboxLookup($box)) {
			Write-Host "Processing Exchange permissions for $User => ($MailboxAlias)"
			$out = Invoke-Command -Session $Session -ScriptBlock ${function:ExchangeCommands} -ArgumentList $User,$box
		}

	} else { # If no parameters given for userid, MailboxAlias, or bulk, assume single person/mailbox and ask for details.
		Write-Host "Guessing what you want from me"
		$User = Read-Host "Enter a UserID"
		$MailboxAlias = Read-host "Enter a Mailbox Alias"

		# Verifying mailbox exists before talking to Exchange
		if (MailboxLookup($box)) {
			Write-Host "Processing Exchange permissions for $User => ($MailboxAlias)"
			$out = Invoke-Command -Session $Session -ScriptBlock ${function:ExchangeCommands} -ArgumentList $User,$box
		}
	}

	# Close out the Exchange PSSession
	Remove-PSSession $Session
	Write-Output "Exited Exchange Session"
	Write-Host "Completed Task"
	Write-Host "======================================="

}
