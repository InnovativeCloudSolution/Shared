#Parameters to receive from flow
param(
    [parameter(Mandatory = $true)]
    [string]$Email='',
	[string]$DistributionList=''
)

Write-Output "The workflow will now attempt to add the user to required Exchange assets.<br>"
$tenantID = 'seasonsliving.com.au'
$appID = '83318188-4697-443f-9392-c8a7c7c1d172'

## Connect to Exchange Online
Connect-ExchangeOnline -CertificateThumbprint "C021FC21F5D11AD88B4EE91042441D5DB0C31ABB" -AppID $appID -Organization $tenantID
Get-Mailbox $Email

Add-DistributionGroupMember -Identity $DistributionList -Member $Email -BypassSecurityGroupManagerCheck -Confirm:$false
$Log += "The user $Email has been added to the $DistributionList DL.<br>"

Write-Output $Log
Disconnect-ExchangeOnline