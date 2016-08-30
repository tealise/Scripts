<#
.SYNOPSIS
    Automates the process on gathering BitLocker recovery password and TPM owner password.

.DESCRIPTION
    This script will lookup multiple attribute in Active Directory and display the correlating
    values that hold sensitive BitLocker information.  Additionally, the TPM Owner Password
    can be exported to a .tpm file, which can be used to make changes to the correlating machine.

.NOTES
    File Name      : winCrypt.ps1
    Author         : Joshua Nasiatka (bitcraftlabs.net)
    Prerequisite   : PowerShell V2 over Vista and upper
    Version History: 5/7/2015 (original release)
    Current Version: 2/5/2016 (not branded for my job edition)

.LINK
    Adapted from:
    http://jackstromberg.com/2015/02/exporting-tpm-owner-key-and-bitlocker-recovery-password-from-active-directory-via-powershell/

    Made into a user-friendly GUI
#>

clear
Write-Host "Bitlocker/TPM Key Retrieval Tool"
write-Host "Joshua Nasiatka (Feb 2016)`n"
Write-Host "loading application..."

function display ($title) {
[void][System.Reflection.Assembly]::LoadWithPartialName( "System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName( "Microsoft.VisualBasic")
[void][System.Reflection.Assembly]::LoadWithPartialName( "System.Drawing")

$form = New-Object "System.Windows.Forms.Form";
$form.Width = 395;
$form.Height = 385;
$form.Text = "Bitlocker/TPM Recovery Tool";
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen;
$form.FormBorderStyle = 'Fixed3D'
$form.ControlBox = $true
$form.MinimizeBox = $false
$form.MaximizeBox = $false
$form.TopMost = $true
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('unlock.ico')


############## Display Logo #############
$img = [System.Drawing.Image]::Fromfile('encrypt_back.png')
$logo = new-object Windows.Forms.PictureBox
$logo.Width =  $img.Size.Width
$logo.Height =  $img.Size.Height
$logo.Image = $img

############## Define Name ##############
$compname = New-Object "System.Windows.Forms.Label";
$compname.Left = 10;
$compname.Top = 65;
$compname.Height = 30;
$compname.Text = "Computer AD Name to Lookup: ";

############Define text box1 for input
$compnameField = New-Object "System.Windows.Forms.TextBox";
$compnameField.Left = 150;
$compnameField.Top = 68;
$compnameField.width = 210;

############Define output
$outputConsole = New-Object "System.Windows.Forms.TextBox";
$outputConsole.Text = "[console-out]"
$outputConsole.Left = 10;
$outputConsole.Top = 98;
$outputConsole.Multiline = $true;
$outputConsole.height = 205;
$outputConsole.width = 350;

#############define RETRIEVE button
$btnRetrieve = New-Object "System.Windows.Forms.Button";
$btnRetrieve.Left = 140;
$btnRetrieve.Top = 308;
$btnRetrieve.Width = 100;
$btnRetrieve.Text = "Retrieve";

### ACTION RETRIEVE ###
$dothething = [System.EventHandler] {

$credential = Get-Credential #New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName , $Password

# Get input on which machine to lookup
$computer = $compnameField.Text; #Read-Host 'Enter in computer name'

# Import our Active Directory PowerShell commands
Import-Module ActiveDirectory

# Check if the Computer Object exists in AD
$computerObject = Get-ADComputer -Filter {cn -eq $computer} -Property msTPM-OwnerInformation, msTPM-TpmInformationForComputer -Credential $credential
if($computerObject -eq $null){
    $outputConsole.Text = "Error: Computer object not found or bad credentials."
}

# Windows Vista and 7 stores the TPM owner password in the msTPM-OwnerInformation attribute, check that first.
# If the key hasn't been stored there, check the msTPM-TpmInformationForComputer object to see if it was backed up on a Win 8 or greater machine

if($computerObject.'msTPM-OwnerInformation' -eq $null){
    #Check if the computer object has had the TPM info backed up to AD
    if($computerObject.'msTPM-TpmInformationForComputer' -ne $null){
        # Grab the TPM Owner Password from the msTPM-InformationObject
        $TPMObject = Get-ADObject -Identity $computerObject.'msTPM-TpmInformationForComputer' -Properties msTPM-OwnerInformation  -Credential $credential
        $TPMRecoveryKey = $TPMObject.'msTPM-OwnerInformation'
    }else{
        $TPMRecoveryKey = '<not set>'
    }
}else{
    # Windows 7 and older OS TPM Owner Password
    $TPMRecoveryKey = $computerObject.'msTPM-OwnerInformation'
}

# Check if the computer object has had a BitLocker Recovery Password backed up to AD
$BitLockerObject = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $computerObject.DistinguishedName -Properties 'msFVE-RecoveryPassword' -Credential $credential | Select-Object -Last 1
if($BitLockerObject.'msFVE-RecoveryPassword'){
    $BitLockerRecoveryKey = $BitLockerObject.'msFVE-RecoveryPassword'
}else{
    $BitLockerRecoveryKey = '<not set>'
}

#Print out our findings
$theoutput = New-Object System.Text.StringBuilder
$theoutput.AppendLine('TPM Owner Recovery Key: ')
$theoutput.AppendLine($TPMRecoveryKey)
$theoutput.AppendLine()
$theoutput.AppendLine('BitLocker Recovery Password: ')
$theoutput.AppendLine($BitLockerRecoveryKey)
$outputConsole.Text = $theoutput

# Export TPM Owner Password File

 $thenewoutput = "C:\RecoveredKeys\TPM-" + $compnameField.Text + ".tpm"
 $TPMOwnerFile = '<?xml version="1.0" encoding="UTF-8"?><ownerAuth>' + $TPMRecoveryKey + '</ownerAuth>'
 $TPMOwnerFile | Out-File -force -filePath $thenewoutput
 $BitlockerKey += "--Bitlocker Recovery Key-- `r`n"
 $BitlockerKey += "Computer Name: " + $compnameField.Text + " `r`n"
 $BitlockerKey += "Recovery Key: " + $BitLockerRecoveryKey + " `r`n"
 $thenewoutput2 = "C:\RecoveredKeys\Bitlocker-" + $compnameField.Text + ".txt"
 $BitlockerKey | Out-File -force -filePath $thenewoutput2

}

### ADD THAT ACTION ###
$btnRetrieve.Add_Click($dothething);

#############Add controls to all the above objects defined
$form.Controls.Add($compname);
$form.Controls.Add($compnameField);
$form.Controls.Add($outputConsole);
$form.Controls.Add($btnRetrieve);
$form.Controls.Add($logo);
$ret = $form.ShowDialog();

} display
