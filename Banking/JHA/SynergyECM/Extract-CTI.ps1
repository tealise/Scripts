[xml]$xmlData = Get-Content "cti.xml"
$cabinets = $xmlData.SynergyDocDbStructure.Cabinet | Sort Name
$types    = $xmlData.SynergyDocDbStructure.Type | Sort Name
$indexes  = $xmlData.SynergyDocDbStructure.Index | Sort Name

$output = @"
=============================================
| EXPORT OF SYNERGY CTI STRUCTURE
| available Cabinets, Types, Indexes
|
| as of $(Get-Date -Format "yyyy.MM.dd")
=============================================

====== SYNERGY CABINETS ======
$($cabinets.Name | Format-Table -AutoSize | Out-String)

======= SYNERGY TYPES ========
$($types.Name | Format-Table -AutoSize | Out-String)

====== SYNERGY INDEXES =======
$(($indexes | Select Name, DataType, Precision | Format-Table -AutoSize | Out-String))
"@

$output

$output | Out-File -FilePath "cti.txt" -Encoding ascii
