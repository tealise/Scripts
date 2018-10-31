# #############################################################################
# ACTIVE DIRECTORY - IP PHONE EXTENSION
# NAME: Copy-ADUserPhoneToIPPhone.ps1
#
# AUTHOR: Joshua Nasiatka
# DATE:   2016/09/23
#
# COMMENT:  This script will copy the last four digits of "Main" phone
#           to "IP Phone" for every user account in Active Directory
#
# VERSION HISTORY
# 1.0 2016.09.23 Initial Version.
#
# FEATURES
# -Get-ADUser Phone Number, Last 4 digits
# -Push extension to "IP Phone"
# #############################################################################

 Import-Module ActiveDirectory

 echo "Beginning TelephoneNumber -> IPPhone Property Copy"

 Get-ADUser -Filter {Enabled -eq $true} -SearchBase “dc=domain,dc=local” -Properties TelephoneNumber | Where-Object {$_.TelephoneNumber -ne $null} | ForEach {
    $phoneNumber = $_.TelephoneNumber.Substring($_.TelephoneNumber.Length - 4, 4)
    Set-ADUser -Identity $_.SAMAccountName -Replace @{ipPhone=$phoneNumber}
 }

 echo "The properties have been succesfully applied to the AD Objects"
