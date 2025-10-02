<#

Mangano IT - Run Ookla Speedtest
Created by: Gabriel Nugent
Version: 1.1.2

This runbook is designed to be run on MIT-HMLN1-MGT01.

#>

## SCRIPT VARIABLES ##

$Path = 'C:\Packages\Speedtest\speedtest.exe'
$LogFolder = 'C:\Logs\MGT01-Speedtest'

# Build log file name
$Date = Get-Date -UFormat "%Y%m%dT%R" | ForEach-Object { $_ -replace ":", "" }
$FileName = "$LogFolder\MGT01-Speedtest-$Date.txt"

## RUN SPEEDTEST ##

try {
    Write-Warning "Running speed test..."
    $Result = & $Path --f 'json-pretty' --accept-license
    Write-Warning "Speed test successfully run!"
} catch {
    Write-Error "Unable to run speed test : $($_)"
}

## SEND OUTPUT TO FLOW AND LOGS FOLDER ##

Write-Output $Result
$Result | Out-File -FilePath $FileName