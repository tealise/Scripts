# #############################################################################
# J. NASIATKA - FINSCRIPTS - ACH PARSE TOOL
# NAME: Parse-ACH.ps1
#
# AUTHOR: Joshua Nasiatka
# DATE:   2018/09/07
# EMAIL:  dev@joshuanasiatka.com
#
# COMMENT:  This script can parse NACHA-formatted ACH files and either convert
#           to csv or filter contents for specific record details and export
#           to ACH.
#
# VERSION HISTORY
# 1.0 2018.09.07 Initial Version.
#
# #############################################################################

############################### CONFIG SETTINGS #################################
$input_path = "Input"
$output_path = "Output"
$output_name_scheme = "ACH_EXTRACT_"
$csv = $false
$ach = $true
$combine = $true
$rdfi = '33714510'
$filter = @{
  'record_type' = '6' # 6 (Entry Detail PPD), 7 (ADDENDA CCD+), 8 (BATCH CONTROL), 9 (FILE CONTROL)
  'trans_code'  = '22'
  'odfi'        = '65456789'
  # 'rdfi'        = '32123456'
  # 'indv_name'   = 'COMPANY NAME'
}
$record_type = $filter.record_type
#################################################################################

$ach_files = (Get-ChildItem -Path $input_path).FullName
$combined_ach = @()
$combined_csv = @()

############################
$ErrorActionPreference = "Stop"

function Compare-Hashtable {
<#
.SYNOPSIS
Compare two Hashtable and returns an array of differences.
.DESCRIPTION
The Compare-Hashtable function computes differences between two Hashtables. Results are returned as
an array of objects with the properties: "key" (the name of the key that caused a difference),
"side" (one of "<=", "!=" or "=>"), "lvalue" an "rvalue" (resp. the left and right value
associated with the key).
.PARAMETER left
The left hand side Hashtable to compare.
.PARAMETER right
The right hand side Hashtable to compare.
.EXAMPLE
Returns a difference for ("3 <="), c (3 "!=" 4) and e ("=>" 5).
Compare-Hashtable @{ a = 1; b = 2; c = 3 } @{ b = 2; c = 4; e = 5}
.EXAMPLE
Returns a difference for a ("3 <="), c (3 "!=" 4), e ("=>" 5) and g (6 "<=").
$left = @{ a = 1; b = 2; c = 3; f = $Null; g = 6 }
$right = @{ b = 2; c = 4; e = 5; f = $Null; g = $Null }
Compare-Hashtable $left $right
#>
[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Hashtable]$Left,

        [Parameter(Mandatory = $true)]
        [Hashtable]$Right
	)

	function New-Result($Key, $LValue, $Side, $RValue) {
		New-Object -Type PSObject -Property @{
					key    = $Key
					lvalue = $LValue
					rvalue = $RValue
					side   = $Side
			}
	}
	[Object[]]$Results = $Left.Keys | % {
		if ($Left.ContainsKey($_) -and !$Right.ContainsKey($_)) {
			New-Result $_ $Left[$_] "<=" $Null
		} else {
			$LValue, $RValue = $Left[$_], $Right[$_]
			if ($LValue -ne $RValue) {
				New-Result $_ $LValue "!=" $RValue
			}
		}
	}
	$Results += $Right.Keys | % {
		if (!$Left.ContainsKey($_) -and $Right.ContainsKey($_)) {
			New-Result $_ $Null "=>" $Right[$_]
		}
	}
	$Results
}
############################
Function ConvertHashtableTo-Object {
    [CmdletBinding()]
    Param([Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True)]
        [hashtable]$ht
    )
    PROCESS {
        $results = @()

        $ht | %{
            $result = New-Object psobject;
            foreach ($key in $_.keys) {
                $result | Add-Member -MemberType NoteProperty -Name $key -Value $_[$key]
             }
             $results += $result;
         }
        return $results
    }
}
############################

Function SearchFilter ($file, $filter) {
  $ach_data = Get-Content -Path $file
  $ach_filtered = @()
  $csv_data = @()
  foreach ($line in $ach_data) {
    if ($line[0] -eq $filter.record_type) {
      $line_details = ReadACHLine $line $filter.record_type
      # $filter_results = Compare-Hashtable $filter $line_details | ?{$($_).lvalue}
      if ((Compare-Hashtable $filter $line_details | ?{$_.side -eq '!='}).count -eq 0) {
        $ach_filtered += $line
        $csv_data += $line_details | ConvertHashtableTo-Object
      }
    }
  }
  return @{
    'ach_filtered' = $ach_filtered
    'csv_data'     = $csv_data
  }
}

Function ReadACHLine ($line, $record_type) {
  if ($record_type -eq '6') {
    $record_details  = @{
      'record_type'  = $line.substring(0,1).trim()
      'trans_code'   = $line.substring(1,2).trim()
      'rdfi'         = $(try{[int]$line.substring(3,9).trim()}catch{$line.substring(3,9).trim()})
      'check_digit'  = $line[11]
      'odfi'         = $(try{[int]$line.substring(12,17).trim()}catch{$line.substring(12,17).trim()})
      'amount'       = $(try{($line.substring(29,10).trim())/100}catch{$line.substring(29,10).trim()})
      'indv_id'      = $line.substring(39,15).trim()
      'indv_name'    = $line.substring(54,22).trim()
      'discr_data'   = $line.substring(76,2).trim()
      'addenda_bool' = $line.substring(78,1).trim()
      'trace_number' = $line.substring(79,15).trim()
    }
  }

  return $record_details
}

Function ExportACH ($records, $output_path, $output_name_scheme) {
  $records | Out-File -FilePath "$output_path\$($output_name_scheme + $(date).ToString('yyyyMMdd_hhmmss')).ach" -encoding ascii
}

Function ExportCSV ($records, $output_path, $output_name_scheme) {
  $records.GetEnumerator() | Sort-Object Name | Export-CSV "$output_path\$($output_name_scheme + $(date).ToString('yyyyMMdd_hhmmss')).csv" -nti
}

$file_count = $ach_files.count
$index = 1
$time_begin = (Get-Date).Millisecond
Foreach ($file in $ach_files) {
  Write-Host "Reading file [$index of $file_count]: $file"
  $records = SearchFilter $file $filter #record_type
  if (($records.ach_filtered).Count -ge 1) {
    if ($combine) {
      $combined_ach += $records.ach_filtered
      $combined_csv += $records.csv_data
    } else {
      ExportACH $($records.ach_filtered) $output_path $output_name_scheme
      ExportCSV $($records.csv_data) $output_path $output_name_scheme
    }
  }
  $index++
}

if ($combine) {
  ExportACH $combined_ach $output_path $output_name_scheme
  ExportCSV $combined_csv $output_path $output_name_scheme
}

$time_end = (Get-Date).Millisecond
Write-Host "Parse job completed in $($time_end - $time_start) milliseconds."
$combined_ach
