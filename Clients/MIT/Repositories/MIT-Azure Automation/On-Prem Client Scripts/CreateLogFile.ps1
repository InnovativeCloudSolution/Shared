<#

Mangano IT - Local Device - Create Log File
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][string]$Log,
    [Parameter(Mandatory=$true)][string]$FilePath,
    [Parameter(Mandatory=$true)][string]$FileName
)

if (!(Test-Path -PathType container $FilePath)) { New-Item -ItemType Directory -Path $FilePath }
$Log | Out-File -FilePath "$FilePath\$FileName" -Force -Confirm:$false