<#
    .SYNOPSIS
        Merge-CSV is used to Combine two .csv files on a common
        parameter/key field. 
    .DESCRIPTION
        Merge-CSV is used to Combine two .csv files on a common
        parameter/key field. 
    .PARAMETER CSVPath1
        Path to first .csv file to merge
    .PARAMETER CSVPath2
        Path to second .csv file to merge with the first
    .PARAMETER MergeKey1
        Column name in first csv to use as merge key
    .PARAMETER MergeKey2
        Column name in second csv to use as merge key
    .PARAMETER OutPath
        Path to output merged csv
    .INPUTS
        Takes two .csv files 
    .OUTPUTS
        Will output a .csv
    .EXAMPLE  
        Merge-CSV -CSVPath1 ".\file1.csv" -CSVPath2 ".\file2.csv" -CSVMergeKey1 "Addresses" -CSVMergeKey2 "IPs" -OutPath ".\Merged.csv"
#>

function Merge-CSV {

    [CmdletBinding()]
    param (
        # Path to first .csv file
        [parameter(Mandatory=$True)]
        [string]$CSVPath1,

        # Path to second .csv file
        [parameter(Mandatory=$True)]
        [string]$CSVPath2,

        # Column header to merge on in $CSVPath1
        [parameter(Mandatory=$True)]
        [string]$CSVMergeKey1,
        
        # Column header to merge on in $CSVPath2
        [parameter(Mandatory=$True)]
        [string]$CSVMergeKey2,
        
        # Path to output the merged .csv file
        [parameter(Mandatory=$True)]
        [string]$OutPath
    )

    # Import .csv files to be merged
    $CSVImport_1 = Import-Csv -Path $CSVPath1
    $CSVImport_2 = Import-Csv -Path $CSVPath2

    # Convert $CSVMergeKey2 to $CSVMergeKey1
    if ($CSVMergeKey1 -ne $CSVMergeKey2) {
        $CSVImport_2 = $CSVImport_2 |select @{N=$CSVMergeKey1; E={$_.$CSVMergeKey2}}, * -ExcludeProperty $CSVMergeKey2
    }
    
    # Build list of CSV field headers
    [System.Collections.ArrayList]$CSVFieldList_1 = $CSVImport_1 |`
        gm -MemberType NoteProperty |select -ExpandProperty Name
    [System.Collections.ArrayList]$CSVFieldList_2 = $CSVImport_2 |`
        gm -MemberType NoteProperty |select -ExpandProperty Name 
    
    $CSVFieldList_2.Remove($CSVMergeKey1)
    
    # Collect list of duplicate keys
    [System.Collections.ArrayList]$CSVFieldListDupes = foreach($key in $CSVFieldList_1) {
        if($key -in $CSVFieldList_2) { 
            $key 
        } 
    }
    
    #If duplicate fields exist that are not the Merge Key modify duplicates in 2nd csv
    if ($CSVFieldListDupes.Count -gt 0) {
        foreach($dupe in $CSVFieldListDupes) {
            $CSVImport_2 = $CSVImport_2 |select @{N=$dupe+"2"; E={$_.$dupe}}, * -ExcludeProperty $dupe
            $CSVFieldList_2 = $CSVFieldList_2.replace($dupe,$dupe+"2")
        }
    }

    # Collect unique merge values
    [System.Collections.ArrayList]$uniqueKeyValues = $CSVImport_1.$CSVMergeKey1 + $CSVImport_2.$CSVMergeKey1 |sort -Unique

    # Collect unique header values
    [System.Collections.ArrayList]$CSVFieldList_ALL = $CSVFieldList_1 + $CSVFieldList_2

    #region Build hashtable with all unique Merge values and empty fields
    [System.Collections.ArrayList]$AllHashList = @()

    foreach ($value in $uniqueKeyValues) {
        [System.Collections.Hashtable]$HashStore = @{}
        foreach ($field in $CSVFieldList_ALL) {
            $HashStore.Add("$field",'')
        }
        $HashStore.$CSVMergeKey1 = $value
        $AllHashList += $HashStore
    }

    # Remove MergeField from FieldList 
    $CSVFieldList_ALL.Remove($CSVMergeKey1)

    # Update hashtable with data from imported csv files
    foreach ($table in $AllHashList){
        foreach ($value in $CSVFieldList_ALL) {
            if ($table.$CSVMergeKey1 -in $CSVImport_1.$CSVMergeKey1 -and $value -in $CSVFieldList_1) {
                $table.$value = $($CSVImport_1 |where {$_.$CSVMergeKey1 -eq $table.$CSVMergeKey1} |`
                                select -ExpandProperty $value)
            }
            elseif ($table.$CSVMergeKey1 -in $CSVImport_2.$CSVMergeKey1 -and $value -in $CSVFieldList_2 -or $value -eq $CSVMergeKey1) {
                $table.$value = $($CSVImport_2 |where {$_.$CSVMergeKey1 -eq $table.$CSVMergeKey1} |`
                                select -ExpandProperty $value)
            }
            else {
                $table.$value = ''
            }
        }
    }

    # Perform readability formatting on hashtables prior to exporting
    foreach ($table in $AllHashList) {
        foreach ($key in $($table.keys)) {
            $table.$key = $table.$key |select -Unique
            $table.$key = $table.$key -join "`r`n"
        }
    }

    # Export final results
    $AllHashList |ForEach-Object {[PSCustomObject]$_ }|Export-Csv -Path $OutPath -NoTypeInformation
}
