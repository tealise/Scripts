<#
.SYNOPSIS
	Creates a new mailbox for a user with quota and retention policy
.DESCRIPTION
	Script to quickly provision a new mailbox for an existing AD User and set quota and retention
.PARAMETER DisplayName
	The display name to which the DL will be provisioned
.PARAMETER Alias
	The alias of the DL, the stuff before the '@' symbol
.PARAMETER Manager
  The email address of distribution group manager
.PARAMETER UserList
	Filepath of userlist.csv file (e.g. C:\Temp\userlist.csv)
	UserList.csv must contain two columns, DL and UserID
.PARAMETER Domain
	Domain name (e.g. domain.local)
.PARAMETER DomainController
	Preferred Domain Controller
.PARAMETER OU
	Organizational Unit for the shared mailbox to be created
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
	[string]$DisplayName, # The display name of distribution list

	[Parameter(Mandatory=$true)]
	[string]$Alias, # The email alias

	[Parameter(Mandatory=$false)]
	[string]$Manager = 'default@domain.com', # Default Distribution Manager

	[Parameter(Mandatory=$false)]
	[string]$UserList, # full filepath and name for userlist

	[Parameter(Mandatory=$true)]
	[string]$Domain, # AD Domain for User search space

	[Parameter(Mandatory=$false)]
	[string]$DomainController, # AD Domain Controller that's preferred

	[Parameter(Mandatory=$false)]
	[string]$OU, # AD Domain OU that's preferred

  [Parameter(Mandatory=$false)]
  [switch]$CreateOnly = $false,

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

    # Get Default Domain Controller
    if (!($DomainController)) { $DomainController = $env:LogonServer -replace '\\','' }

	# Let's check if it exists first
	if (!(Get-ADUser -LDAPFilter "(mailNickname=$Alias)" -Server $DomainController)) {
		$NopeError = [string]"Can't find DL with the alias '$Alias', defaulting to create"
		$Create = $true
		Write-Warning $NopeError
	} else {
		$Create = $false
		Write-Output "Found a record for distribution list '$DisplayName'"
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

  Write-Output "Processing Exchange account for DL, $DisplayName ($Alias)"

	$CreateDL = Function CreateDL($DisplayName,$Domain,$Alias,$OU) {
		$emailaddr = "$Alias@$Domain"
		New-DistributionGroup -Name $DisplayName -DisplayName $DisplayName -Alias $Alias -PrimarySmtpAddress $emailaddr -OrganizationalUnit $OU
		Set-DistributionGroup -Identity $DisplayName -ManagedBy $Manager -BypassSecurityGroupManagerCheck
	}

	$UpdateDL = Function UpdateDL($DL,$User) {
		Add-DistributionGroupMember -Identity $DL -Member $User
	}

	if ($Create) {
		Invoke-Command -Session $Session -ScriptBlock ${function:CreateDL} -ArgumentList $DisplayName,$Domain,$Alias,$OU
	} else {
        if (!($userlist)) { Write-Warning "Nothing to do since no file imported" }
    }

    if (!($CreateOnly)) {
	    if (Test-Path -Path $UserList) {
		    Write-Output "Updating distribution group members"
		    Import-CSV $userlist | ForEach {
			    Write-Output "Adding $($_.UserID) to $($_.DL)"
			    Invoke-Command -Session $Session -ScriptBlock ${function:UpdateDL} -ArgumentList $_.DL,$_.UserID
		    }
	    } else {
		    Write-Warning "CSV File not found, unable to process distribution list update`nPlease create a CSV with headers DL and UserID (order doesn't matter)"
	    }
    }

	# Close out the Exchange PSSession
	Remove-PSSession $Session
	Write-Output "Exited Exchange Session"
	Write-Output "Completed Task"
	Write-Output "======================================="

}
