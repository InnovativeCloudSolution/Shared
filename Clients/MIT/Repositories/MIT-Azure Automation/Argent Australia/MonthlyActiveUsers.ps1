$SearchOU = 'OU=Users,OU=Argent,DC=internal,DC=argentaust,DC=com,DC=au'

$UserList = Get-ADUser -SearchBase $SearchOU -Filter 'enabled -eq $true' | `
Select-Object -Property Name,SamAccountName,UserPrincipalName

Write-Output $UserList | Sort-Object Name | ConvertTo-Json