<#

Mangano IT - ConnectWise Manage - Create Connection Object
Created by: Gabriel Nugent
Version: 1.0.4

This runbook is designed to be used in conjunction with other PowerShell scripts.

#>

# Set CRM variables, connect to server
$Server = Get-AutomationVariable -Name 'CWManageUrl'
$Company = Get-AutomationVariable -Name 'CWManageCompanyId'
$PublicKey = Get-AutomationVariable -Name 'PublicKey'
$PrivateKey = Get-AutomationVariable -Name 'PrivateKey'
$ClientId = Get-AutomationVariable -Name 'clientId'

# Create an object with CW credentials 
$Connection = @{
	Server = $Server
	Company = $Company
	pubkey = $PublicKey
	privatekey = $PrivateKey
	clientId = $ClientId
}

# Send connection object back to script
Write-Output $Connection