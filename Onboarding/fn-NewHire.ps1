function Elevate {
  # Get AD Credential w/ permission to access print servers (e.g. Domain Admin account)
  if($cred = $host.ui.PromptForCredential('Admin Credentials Required', 'Please enter your systems admin credentials in order to poll Active Directory and Exchange.',
  '', "")){} else {
      Write-Warning "Need systems admin credentials in order to proceed.`r`nPlease re-run script and enter the appropriate credentials."
      exit
  }

  $username = $cred.username
  $password = $cred.GetNetworkCredential().password
  $CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
  $domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain,$UserName,$Password)
  if ($domain.name -eq $null) {
      Write-Warning "Authentication failed - please verify your username and password."
      Elevate
  } else {
      Write-Output "Successfully authenticated with domain" ($domain).name
  }
  return $cred
}

function LocationPicker($locations) {
  Clear-Host
  Write-Output @"
############################################################
#                    LOCATION PICKER                       #
############################################################
"@
  [int]$i=1
  foreach ($location in $locations) {
    Write-Output "[$i]" $location.Name "`t" $Address
    $i++
  }
  $loc = Read-Host "Please select a location"
  return $locations[$loc-1]
}

function RolePicker($roles) {
  ### SORT ROLES & LOCATIONS ###
  $role_list = $roles | ?{$_.Location -ne "" -and $_.Department -ne ""} | Sort Role

  ### DEPARTMENT ###
  $departments = $roles.Department | Sort | Select -Uniq
  if ($departments.Count -le 1) {$departments = @($departments)}

  Clear-Host
  Write-Output @"
############################################################
#                  DEPARTMENT PICKER                       #
############################################################
"@
  [int]$i=1
  foreach ($department in $departments) {
    Write-Output "[$i]" $department
    $i++
  }
  $dept = Read-Host "`r`nPlease select a department"
  $dept_name = $departments[$dept-1]

  ### LOCATIONS ###
  $location_options = (($role_list | ?{$_.Department -eq $dept_name}) | Sort Location -uniq) | ?{$_.Location -ne ""}

  Clear-Host
  Write-Output @"
############################################################
#                     LOCATION PICKER                      #
############################################################
"@
  [int]$i=1
  foreach ($location_option in $location_options) {
    Write-Output "[$i]" $location_option.Location
    $i++
  }
  $loc = Read-Host "`r`nPlease select a location"
  $loc_name = ($location_options[$loc-1]).Location

  ### ROLES ###
  $filtered_roles = $role_list | ?{ ($_.Location -eq $loc_name)}

  Clear-Host
  Write-Output @"
############################################################
#                       ROLE PICKER                        #
############################################################
"@
  [int]$j=1
  foreach ($role_option in $filtered_roles) {
    Write-Output "[$j]" $role_option.Role
    $j++
  }
  $role = Read-Host "`r`nPlease select a role"
  $role_name = $filtered_roles[$role-1]

  Write-Output $role
  return $role_name
}

function GenerateUsername {
  # Usernames in our environment use a A###### format (letter, company ID, and sequential employee number, e.g. E100123)
  # @TODO: Add GenerateUsername function that uses a first intial lastname format
  #        - flastname
  #        - fmlastname (if middle initial provided and flastname not available)
  #        - flastname2 (if middle initial not provided and flastname not available, e.g. jsmith1)
  $last_username = $(Get-ADUser -Filter {sAMAccountName -like "E100*"} -Properties sAMAccountName | Sort sAMAccountName -desc | Select -first 1).sAMAccountName
  $last_username_number = [decimal](($last_username -split 'e')[1])
  $next_available_username = 'E' + [string]([decimal](($last_username -split 'e')[1])+1)

  if ($last_username_number -ge 100950) {
    $warnings += "Approaching end of username space. Seak new naming scheme."
  }

  return $next_available_username
}

function GetUserName {
  $default = $(GenerateUsername)
  $username = Read-Host "Enter a username (blank for $default)"
  if (-not $username) {return $default} else {
    $does_it_exist = Get-ADUser -Filter * | Where {$_.sAMAccountName -eq $username} | Select -First 1
    if ($does_it_exist) {
      Write-Output "The supplied username '$username' already exists in AD."
      GetUserName
    } else {
      return $username
    }
  }
}

function ValidateExtension {
  # Check if extension is in valid DID region based on those configured in PBX system
  $extension = Read-Host  "Enter Extension (if applicable)     "
  if ((($extension -ge 1100) -and ($extension -le 1299)) -or (($extension -ge 3500) -and ($extension -le 3699)) -or (-not $extension)) {
    return $extension
  } else {
    Write-Warning "Invalid extension. Choose available between 1100-1299 and 3500-3699."
    ValidateExtension
  }
}

function ShowSummary {
  # Clear-Host
  Write-Output @"
############################################################
#              NEW HIRE ONBOARDING SUMMARY                 #
############################################################

FIRST NAME:`t$first_name$(" "*(16-$first_name.Length))LAST NAME:`t$last_name
USERNAME:`t$next_available_username$(" "*(16-$next_available_username.Length))EMAIL PREFIX:`t$($first_name[0]).$last_name
PHONE:`t`t$(if($extension){"x"+$extension}else{"N/A"})

LOCATION:`t$($location.Address), $($location.Town)
DEPARTMENT:`t$($role.Department)
JOB TITLE:`t$($role.Role)

AD TEMPLATE:`t$($role.Template)
OU:`t`t$($role.OU)

USERDOCS:
 - \\SERVER01\UserDocs\$next_available_username\My Documents
 - \\SERVER01\UserDocs\$next_available_username\Favorites

############################################################

The onboarding completed with the following warnings...`r`n
WARNINGS:
"@
ListArray($warnings)
Write-Output "`r`nERRORS:"
ListArray($errors)
Write-Output @"
############################################################
"@
}

function ListArray($arr) {
  if (-not ([string]::IsNullOrEmpty($arr))) {
    foreach($item in $arr) {
      if (-not ([string]::IsNullOrEmpty($item))){
        Write-Output " -" $item
      }
    }
  } else {Write-Output " - No items"}
}

############################################################
##               USERDOCS FOLDER CREATION                 ##
############################################################
# Apply ACL Patch to Folder Param
# Folder to have only Creator Owner, System, Builtin\Admin, Domain Admins, and Home Folder User
# Change "DOMAIN" to actual Active Directory domain

Function FixFolder ($udocs, $f) {
    # Create ACL object from scratch
    $acl = New-Object System.Security.AccessControl.DirectorySecurity
    $acl.SetAccessRuleProtection($true,$true) # remove inheritance

    try {
        $user = New-Object System.Security.Principal.NTAccount ("DOMAIN", "$f") # DOMAIN is your NetBIOS Domain Name
        $Rule = New-Object System.Security.AccessControl.FileSystemAccessRule($user, "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")
        $acl.AddAccessRule($Rule)
        $acl.SetOwner($user)

        # We also add System account and local administrators (Replace these french account names if needed)
        $OtherAccounts = @("CREATOR OWNER", "NT AUTHORITY\SYSTEM", "BUILTIN\Administrators", "DOMAIN\Domain Admins")

        ForEach ($Account in $OtherAccounts) {
            $ACLAccount = New-Object System.Security.Principal.NTAccount($Account)
            $Rule = New-Object System.Security.AccessControl.FileSystemAccessRule($ACLAccount, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
            $acl.AddAccessRule($Rule)
        }

        Set-Acl -AclObject $acl -Path "$udocs\$f"
        Get-ChildItem -Path "$udocs\$f" -Force | Set-Acl -AclObject $acl
    } catch {
        $errors += "ERROR: Unable to apply ACL to directory $f :  $($f.Exception.Message)"
    }
}

# Verify ACL, Skip Folder if ACL is OK
Function TestFolder ($udocs,$f) {
    try {
        $curr = (Get-ACL -Path "$udocs\$f").Access | Select-Object -ExpandProperty IdentityReference | Select-Object -ExpandProperty Value | select -Unique | Sort -Descending
        $des = @("NT AUTHORITY\SYSTEM","CREATOR OWNER","BUILTIN\Administrators","DOMAIN\Domain Admins","DOMAIN\$f")
        if ((Compare-Object $curr $des).Length -ne 0) {
            FixFolder $udocs $f
        }
    } catch {
        $errors += "ERROR: Unable to read ACL from directory '$f', please investigate!"
    }
}

Function CreateUserDocs ($udocs, $f) {
  Write-Output "Creating the folder $udocs\$f"
  New-Item -ItemType Directory -Force -Path "$udocs\$f" | Out-Null
  Write-Output "Setting the ACL"
  TestFolder $udocs $f
}
