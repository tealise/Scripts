# #############################################################################
# MISC UTILS - APPLE TO WINDOWS LINK STANDARDS CONVERSION
# NAME: Convert-WeblocToUrl.ps1
#
# AUTHOR: Joshua Nasiatka
# DATE:   2018/06/07
#
# COMMENT:  Standards converter .webloc (xml\plist) to .url (ini)
#
# VERSION HISTORY
# 1.0 2018.06.07 Initial Version.
#
# EXAMPLE:
# .\Convert-WeblocToUrl.ps1 -InDir C:\temp\applelinks -OutDir C:\temp\winlinks
#
##############################################################################

[CmdletBinding()]
Param (
    [Parameter(Mandatory=$true)]
    [string]$InDir, # Filepath of Master List File

    [Parameter(Mandatory=$true)]
    [string]$OutDir # Filepath of the daily dump of PO Box
)

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$files = Get-ChildItem $InDir -Filter *.webloc

$files | Foreach {
    [xml]$fc = Get-Content $_.FullName
    $lnk = $fc.plist.dict.string
    " "
    $fc2 = @"
[code]
[InternetShortcut]
URL=$($lnk)
[/code]
"@
    $fc2
    $fc2 | Out-File "$OutDir\$($_).url" -Encoding ASCII
}
