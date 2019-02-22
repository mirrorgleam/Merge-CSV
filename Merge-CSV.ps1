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
        Merge-CSV -CSVPath1 '.\file1.csv' -CSVPath2 '.\file2.csv' -CSVMergeKey1 'Addresses' -CSVMergeKey2 'IPs' -OutPath '.\Merged.csv'
#>

function Merge-CSV {

    [CmdletBinding()]
    param (
        # Path to first .csv file
        [parameter( Mandatory = $True )]
        [string]$CSVPath1,

        # Path to second .csv file
        [parameter( Mandatory = $True )]
        [string]$CSVPath2,

        # Column header to merge on in $CSVPath1
        [parameter( Mandatory = $True )]
        [string]$CSVMergeKey1,
        
        # Column header to merge on in $CSVPath2
        [parameter( Mandatory = $True )]
        [string]$CSVMergeKey2,
        
        # Path to output the merged .csv file
        [parameter( Mandatory = $True )]
        [string]$OutPath
    )

    # Import .csv files to be merged
    Try {
        $CSVImport_1 = Import-Csv -Path $CSVPath1
    } 
    Catch {
    $ImportError = @"
"Unable to import $CSVPath1
Check path and try again"
"@
        Write-Host $ImportError -ForegroundColor Red
        Return
    }

    # Import .csv files to be merged
    Try {
        $CSVImport_2 = Import-Csv -Path $CSVPath2
    } 
    Catch {
    $ImportError = @"
"Unable to import $CSVPath2
Check paths and try again"
"@
        Write-Host $ImportError -ForegroundColor Red
        Return
    }

    # Test if $CSVMergeKey1 is in $CSVImport_1
    if ($CSVMergeKey1 -notin $($CSVImport_1 | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)) {
        $KeyError = @"
"The value you provided for CSVMergeKey1 `"$CSVMergeKey1`" 
was not found in the header of $($CSVPath1.Split('\')[-1])
Please check provided values and try again"
"@
        Write-host $KeyError -ForegroundColor Red
        Return
    }

    # Convert $CSVMergeKey2 to $CSVMergeKey1
    if ($CSVMergeKey2 -notin $($CSVImport_2 | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)) {
        $KeyError = @"
"The value you provided for CSVMergeKey2 `"$CSVMergeKey2`" 
was not found in the header of $($CSVPath2.Split('\')[-1])
Please check provided values and try again"
"@
        Write-host $KeyError -ForegroundColor Red
        Return
    }
    
    Function Combine-Objects {
        Param (
            [Parameter(mandatory=$true)]$Object1, 
            [Parameter(mandatory=$true)]$Object2
        )
    
        $arguments = [Pscustomobject]@()
 
        foreach ( $Property in $Object1.psobject.Properties){
            $arguments += @{$Property.Name = $Property.value}
        
        }
 
        foreach ( $Property in $Object2.psobject.Properties){
            $arguments += @{ $Property.Name= $Property.value}
        
        }
    
        $Object3 = [Pscustomobject]$arguments
    
        return $Object3
    }

    $CSVFieldList_1 = $CSVImport_1[0].psobject.Properties.name
    
    if ($CSVMergeKey2 -in $CSVFieldList_1) {$CSVMergeKey2 = $CSVMergeKey2 + "2"}
    
    $CSVFieldList_2 = $CSVImport_2[0].psobject.Properties.name
    
    foreach ( $Key in $CSVFieldList_1 ) {
        if ( $Key -in $CSVFieldList_2 ) { 
            $CSVImport_2 = $CSVImport_2 | Select-Object @{ N=$Key+"2" ; E={ $_.$Key } } , * -ExcludeProperty $Key
            $CSVFieldList_2[$CSVFieldList_2.IndexOf($Key)] = $Key + "2"
        } 
    }

    $CSVFieldList_ALL = $CSVFieldList_1 + $CSVFieldList_2

    $UniqueKeyValues = ($CSVImport_1.$csvMergeKey1 + $CSVImport_2.$csvMergeKey2) |sort |Get-Unique

    $objectList = @()

    foreach ($value in $UniqueKeyValues) {

        $customObject = New-Object psobject
        foreach ($item in $CSVFieldList_ALL) {Add-Member -InputObject $customObject -MemberType NoteProperty -Name $item -Value ""}

        $CSVEntry_1 = $CSVImport_1 |where {$_.$csvMergeKey1 -eq $value}
        $CSVEntry_2 = $CSVImport_2 |where {$_.$csvMergeKey2 -eq $value}

        if ($CSVEntry_1 -and $CSVEntry_2) {
            $combinedObjects = Combine-Objects -Object1 $CSVEntry_1 -Object2 $CSVEntry_2
            foreach ($field in $CSVFieldList_ALL) {
                $customObject.$field = $combinedObjects.$field
            }
         $objectList += $customObject
        }
    
        elseif ($CSVEntry_1) {
            foreach ($field in $CSVFieldList_ALL) {$customObject.$field = $CSVEntry_1.$field}
            $objectList += $customObject
        }
    
        elseif ($CSVEntry_2) {
            foreach ($field in $CSVFieldList_ALL) {$customObject.$field = $CSVEntry_2.$field}
            $objectList += $customObject
        } 
    }

    $objectList |Export-Csv -Path $OutPath -NoTypeInformation

}
