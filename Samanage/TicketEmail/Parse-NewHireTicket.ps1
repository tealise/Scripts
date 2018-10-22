# Path to .msg files
$msgDir = ".\Input" # directory containing *.msg files from Samanage 
$export = ".\NewHires_" + (Get-Date -Format "yyyyMMdd") + ".csv"
# Array to store results
$msgArray = New-Object System.Collections.Generic.List[object]
$AllNewHires = @()

# Loop throuch each .msg file
$outlook = New-Object -comobject outlook.application
Get-ChildItem "$msgDir" -Filter *.msg |
    ForEach-Object {
        # Open .msg file
        $msg = $outlook.Session.OpenSharedItem($_.FullName)
        # Add .msg file Subject and Body to array
        $msgArray.Add([pscustomobject]@{Subject=$msg.Subject;Body=$msg.Body;})
        $msg.Close(0) # Close doesn't always work, see KB2633737 -- restart ISE/PowerShell
    }
$outlook.Exit

# Loop though / parse each message
ForEach ($message in $msgArray) {
    $subject    = $message.subject
    $body       = ($message.body).Trim().Split("`n")
    $name       = $body[($body|select-string -pattern 'Full Name'|select -expand 'LineNumber')].Trim()
    $name_split = $name.split(" ")
    if($name_split.Count -ge 3){ $mi = $name_split[1] } else { $mi = "" }

    $new_hire   = [PSCustomObject]@{
      "FirstName"   = $name_split[0]
      "MI"          = $mi
      "LastName"    = $name_split[$name_split.Count - 1]
      "DisplayName" = $name
      "Location"    = $body[($body|select-string -pattern 'Site'|select -expand 'LineNumber')].Trim()
      "Department"  = $body[($body|select-string -pattern 'Department'|select -expand 'LineNumber')].Trim()
      "Title"       = $body[($body|select-string -pattern 'Job Title'|select -expand 'LineNumber')].Trim()
      "Supervisor"  = $body[($body|select-string -pattern 'Supervisor'|select -expand 'LineNumber')].Trim()
      "StartDate"   = $body[($body|select-string -pattern 'Start Date'|select -expand 'LineNumber')].Trim()
      "Birthdate"   = $body[($body|select-string -pattern 'Birthday'|select -expand 'LineNumber')].Trim()
    }

    $AllNewHires += $new_hire
}

$AllNewHires
$AllNewHires | Export-CSV -Path $export -nti
