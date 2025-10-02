Clear-Host

# Set Group Variables
$ExpiryGroup = 'SG.Azure.Policy.NeverExpirePassword'
$MFAExcludeGroup = 'SG.Azure.Policy.AzureMFA.Exclude'
$TenantId = "746036af-6519-4420-ba95-65e211e271c6"

## Program details - please remember to update the version number when making changes
Write-Host "`nSeasons Living - Disable Password Expiry" -ForegroundColor Yellow
Write-Host "Version: " -ForegroundColor yellow -NoNewLine; Write-Host "1.2"
Write-Host "Created by: " -ForegroundColor yellow -NoNewLine; Write-Host "Damien Nikiforides"
Write-Host "Maintained by: " -ForegroundColor yellow -NoNewLine; Write-Host "Gabriel Nugent"

# Log the user into 365
$AdminLogin = Read-Host "`nAdmin login"
Connect-AzureAD -TenantId $TenantId -AccountId $AdminLogin
Clear-Host

# Set up the recursive function for getting nested groups.

Function Get-RecursiveAzureAdGroupMemberUsers{
    [cmdletbinding()]

    param(
        [parameter(Mandatory=$True,ValueFromPipeline=$true)]
        $AzureGroup
    )

    Begin{
        if (-not(Get-AzureADCurrentSessionInfo)) { Connect-AzureAD }
    }

    Process {
        Write-Verbose -Message "Enumerating $($AzureGroup.DisplayName)"
        $Members = Get-AzureADGroupMember -ObjectId $AzureGroup.ObjectId -All $true
        $UserMembers = $Members | Where-Object{$_.ObjectType -eq 'User'}
        if ($Members | Where-Object {$_.ObjectType -eq 'Group'}) {
            $UserMembers += $Members | Where-Object{$_.ObjectType -eq 'Group'} | ForEach-Object{ Get-RecursiveAzureAdGroupMemberUsers -AzureGroup $_}
        }
    }
    end {
        Return $UserMembers
    }
}

Write-Host "Gathering members of" $ExpiryGroup
$data = Get-AzureADGroup -SearchString $ExpiryGroup | Get-RecursiveAzureAdGroupMemberUsers | Select-Object UserPrincipalName

#Import the CSV

#$data = Import-Csv 'userslist.csv'

$data | Format-Table | Out-String | ForEach-Object {Write-Host $_}

#Set the expiry settings

$ChangeExpiry = Read-Host "Should the above users be set to never expire? (Y/N)"

if ($ChangeExpiry -eq 'y')
{ 
    Clear-Host
    Write-Host "Setting the users to never expire, please wait..."
    $data | ForEach-Object { Set-AzureADUser -ObjectId $_.UserPrincipalName -PasswordPolicies DisablePasswordExpiration}
    Clear-Host
    Write-Host "Waiting for 60 seconds before checking the expiry state..."
    Start-Sleep -s 60
    Clear-Host
    Write-Host "See the below expiry state of the users after applying the disable setting:"
    #Get-MSOLUser -UserPrincipalName mits.test@seasonsliving.com.au | Select UserPrincipalName, PasswordNeverExpires
    $data | ForEach-Object { Get-AzureADUser -ObjectId $_.UserPrincipalName | Select-Object UserPrincipalName, PasswordPolicies} `
    | Format-Table | Out-String | ForEach-Object {Write-Host $_}
}
else
{ 
    Clear-Host
    Write-Host "Below is the current expiry state of your users:"
    Write-host "This might take a while..."

    $data | ForEach-Object { Get-AzureADUser -ObjectId $_.UserPrincipalName | Select-Object UserPrincipalName, PasswordPolicies} `
    | Format-Table | Out-String | ForEach-Object {Write-Host $_}
}

#Display the current state.

#$data | ForEach-Object {Get-MSOLUser -UserPrincipalName $_.UserPrincipalName}
#$data | ForEach-Object {Write-Host $_.UserPrincipalName}
#Write-Host "Test"
#Get-MSOLUser -UserPrincipalName mits.test@seasonsliving.com.au | Select UserPrincipalName, PasswordNeverExpire


 
 #add users to MFA Exclude group
 
$MFAExcludeAdd = Read-Host "Do you want me to try adding them to the MFA exclude group $MFAExcludeGroup (Y/N)" 
if ($MFAExcludeAdd -eq 'y') { 
    Clear-Host
    Write-Host "Adding accounts to the SG.Azure.Policy.AzureMFA.Exclude group. If they are already a member, an error will appear."
    $data | ForEach-Object { 
        $ObjectId = (Get-AzureADUser -ObjectId $_.UserPrincipalName).ObjectId
        if ($ObjectId) {
            Write-Host "Adding user" $_.UserPrincipalName
            try { Get-AzureADGroup -SearchString $MFAExcludeGroup | Add-AzureADGroupMember -RefObjectId $ObjectId }
            catch { Write-Host $_.UserPrincipalName "is already a member of the given group." }
        }
    }
    Write-Host "All users have been added to the requested group."
    start-sleep -s 2
}
else { 
    Clear-Host
    Write-Host "Please add the users manually if required."
    start-sleep -s 2
    Clear-Host
}

Write-Host "These are the current members of $MFAExcludeGroup"
Get-AzureADGroup -SearchString $MFAExcludeGroup | Get-AzureADGroupMember | Select-Object UserPrincipalName | Format-Table `
| Out-String | ForEach-Object {Write-Host $_}

$export = Read-Host "Do you want to export MFA exclude and Expiry state for all members to a CSV? (Y/N)"
if ($export -eq 'y') { 
    #$data | ForEach-Object {Get-MSOLUser -UserPrincipalName $_.UserPrincipalName | Select UserPrincipalName, PasswordNeverExpires} | Export-Csv .\Export\PasswordExpiryState.csv
    Write-Host "Wait a while..."
    $CurrentDate = Get-Date
    $CurrentDate = $CurrentDate.ToString('yyyy_MM_dd_hh-mm-ss')
    Get-AzureADUser -All | Select-Object UserPrincipalName, PasswordPolicies | Export-Csv .\Export\$CurrentDate.PasswordExpiryState.csv
    Get-AzureADGroup -SearchString $MFAExcludeGroup | Get-AzureADGroupMember | Select-Object UserPrincipalName | `
    Export-Csv .\Export\$CurrentDate.AzureMFA.Exclude.csv
    Write-Host "CSV exported, closing program..."
    start-sleep -s 2
    Clear-Host
}

else { 
    Clear-Host
    Write-Host "Closing program..."
    start-sleep -s 2
    Clear-Host
}

#Add-AzureADGroupMember -ObjectId bcc05e65-b230-48e4-93f7-7e9817c0c7b1

#Get-AzureADUser -All $true | Set-AzureADUser -PasswordPolicies DisablePasswordExpiration
#
#Get-AzureADUser -ObjectId BI_RSW01@seasonsliving.com.au  | Select-Object DisplayName,UserPrincipalName,Department
#Get-MSOLUser -UserPrincipalName BI_RSW02@seasonsliving.com.au | Select UserPrincipalName, PasswordNeverExpires
#Get-MSOLUser -all | Select UserPrincipalName, PasswordNeverExpires | Export-Csv C:\passwordexpire-after.csv