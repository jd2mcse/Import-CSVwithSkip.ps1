<#
.SYNOPSIS
    Quickly import large CSV files after skipping a set number of lines, or 
    the lines above a word found in the header. 
.DESCRIPTION
    Choosing a CSV export from web reports often has a number of lines before the header for 
    the data. In a very large file, skipping this information can be time consuming.
    This script uses System.IO.StreamReader to read the lines to skip, then 
    quickly gets the rest of the file with  the ReadToEnd() method.
.PARAMETER FilePath
    The full path to the delimited file
.PARAMETER Delimiter
    The delimiter for the CSV text. Default is comma, you may specify an alternate, like "`t" for tab
.PARAMETER FindStartWord
    A unique word in the header. Search will start in line above the found word, regardless of location in line
.PARAMETER MaxSearchLines
    Stop seaching for the FindStartWord after reading this number of lines. Default is 100, 
    this is to prevent search continuing indefinitely
.PARAMETER SkipLines
    An alternative to finding a word in the header, this allows you to directly set number of lines to skip before import.
.NOTES
    Alan Kaplan 2/11/23. This function requires .NET methods
.LINK
    www.akaplan.com
.EXAMPLE
    Import-CSVwithSkip -FilePath "C:\hugefile.csv" -skiplines 5
    Import hugefile.csv after skipping the first 5 lines. Line 6 is the header.
.EXAMPLE
    Import-CSVwithSkip -FilePath "C:\hugefile.txt" -delimiter "`t" -skiplines 8
    Import hugefile.txt using tab as delimiter and after skipping the first 8 lines. Line 9 is the header.
.EXAMPLE
    Import-CSVwithSkip -FilePath "C:\hugefile.csv" -FindStartWord "Zip"
    Find the word "Zip" in hugefile.csv. Skip whatever number of lines are above, so the line with Zip is used as the header.
.EXAMPLE
    Import-CSVwithSkip -FilePath "C:\hugefile.csv" -FindStartWord "Zip" -MaxSearchLines 25
    Find the word "Zip" in hugefile.csv, quitting if not found in first 25 lines. 
    Skip whatever number of lines are above, so the line with Zip is used as the header.
#>

Function Import-CSVwithSkip {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [string]
        $FilePath,

        [Parameter(Mandatory = $False)]
        [string]
        $delimiter = ',',

        [Parameter(
            Mandatory = $True,
            ParameterSetName = 'Set1'
        )]
        [string]
        $FindStartWord,

        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'Set1'
        )]
        [uint16]
        $MaxSearchLines = 100,

        #Set the number of lines to ignore. Must be greater than 0
        [Parameter(
            Mandatory = $True,
            ParameterSetName = 'Set2'
        )]
        [ValidateScript({$_ -gt 0})]
        [uint16]
        $skiplines
    )

    $FN = [System.IO.Path]::GetFileName($filePath) 
    Write-Host "Importing data from $FN" -ForegroundColor Green

    # Read in the data skipping header lines by line count
    if ($FindStartWord) {
        $reader = [System.IO.StreamReader]::new($filePath)   
        $line = $reader.ReadLine()
        $i = 0
        #MaxSearchLines bails search if not found
        while (($line -inotmatch $FindStartWord) -and ($i -le $MaxSearchLines)) {
            $i++
            Write-Progress "Searching line $i" "From $FN"
            $line = $reader.ReadLine()        
        }

        if ($i -gt $MaxSearchLines){
            Write-Progress "Search failed" -Completed
            Write-Warning "Failed to find `"$findStartWord`" within $MaxSearchLines lines, quitting."
            Start-Sleep -Seconds 3
            Exit
        }

        $skiplines = $i
        $reader.Close() ; $reader.Dispose()
    }

    # Read in the data skipping header lines by line count
    if ($skiplines) {
        $reader = [System.IO.StreamReader]::new($filePath)
        $i = 0
        while ($i -lt $skiplines  ) {
            $i++
            Write-Progress "Skipped line $i" "From $FN"
            $reader.ReadLine() | Out-Null
        }
    }
  
    Write-Progress "Reading text following $skiplines skipped lines in $FN"
    $data = $reader.ReadToEnd()
    $reader.Close() ;  $reader.Dispose()
    $data | Out-String | ConvertFrom-Csv  -Delimiter $delimiter
    
    Write-Host "Import done`n" -ForegroundColor Green
    Write-Progress "Done" -Completed
}


