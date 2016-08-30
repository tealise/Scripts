# AD-Bitlocker-Recovery
#### Joshua Nasiatka (Feb 2016)

Access Bitlocker keys and create TPM unlock files with ease

Here’s to a program to quickly retrieve and export Bitlocker Recovery Keys and TPM Owner Information (*.tpm) files to regain access to encrypted machines. Now it is possible for techs and other IT staff (in a Bitlocker Recovery accessible OU) to be able to type in a computer name and get its respective Bitlocker Key and TPM Owner Information (msTPM-OwnerInformation extended attribute).

The files will be automatically exported to <code>C:\RecoveredKeys\\</code>. To maintain system integrity, make sure to delete these exported files after immediate use. You don't need those files chilling on your system.

Lastly, if you'd like custom branding for the application, modify the <code>back.png</code> file matching the height x width specs of that one and then recompile using a .bat to .exe converter.

### Prerequisites
-	Must be run as an AD Account whose is in an OU with access to Bitlocker Recovery
-	Remote Server Administration Tools must be installed, more specifically:
  - Role Administrations Tools
  - AD DS and AD LDS Tools
  - Active Directory Module for Windows PowerShell

### Notes
-	Recovered Keys are saved in the directory <code>C:\RecoveredKeys\\</code>
-	These keys are saved as:
  - <code>Bitlocker-< computer-name >.txt</code>
  - <code>TPM-< computer-name >.tpm</code>

### Changelog
[*Version 1.0*](https://github.com/joshuanasiatka/Winadmin-Tools/blob/master/bitlocker-recovery.exe) - Created the GUI application
