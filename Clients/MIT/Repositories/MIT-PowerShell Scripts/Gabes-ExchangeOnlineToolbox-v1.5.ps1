## MODULES ##
Import-Module ExchangeOnlineManagement

## Clears PowerShell terminal
Clear-Host

## Program details - please remember to update the version number when making changes
Write-Host "`nGabe's Exchange Online Toolbox" -ForegroundColor Yellow
Write-Host "Version: " -ForegroundColor yellow -NoNewLine; Write-Host "1.5"
Write-Host "Created by: " -ForegroundColor yellow -NoNewLine; Write-Host "Gabriel Nugent"
Start-Sleep -Seconds 0.5

## Get UPN of admin performing removal task
$AdminAccountUpn = $(Write-Host "`nAdmin login: " -ForegroundColor yellow -NoNewLine; Read-Host)
if ($AdminAccountUpn -eq "gabe") { $AdminAccountUpn = "adm.gabriel.nugent@manganoit.com.au" }
Write-Host "`nConnecting to Exchange Online..."

## Connect to required services
Connect-ExchangeOnline -UserPrincipalName $AdminAccountUpn -ShowBanner:$false -ShowProgress $true

## List attached shared mailboxes
function Get-ListOfSharedMailboxes($RunAsChildProcess, $EmailAddress) {
    $MailboxList = @()

    ## Get email address of user
    if ($false -eq $RunAsChildProcess) {
        Clear-Host
        $EmailAddress = $(Write-Host "`nPlease provide the user's email address: " -ForegroundColor yellow -NoNewLine; Read-Host)
    }

    ## Fetch list of shared mailboxes that the user has access to
    Write-Host "`nFetching list of shared mailboxes for"$EmailAddress"..."
    try {
        $MailboxList = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Get-EXOMailboxPermission |Select-Object Identity,User,AccessRights | `
        Where-Object {($_.user -like $EmailAddress)}
    }
    catch {
        Write-Host "ERROR: Most likely missing correct permissions to administrate Exchange Online." -ForegroundColor Red
        Write-Host "Returning to main menu..."
        Start-Sleep -Seconds 2
        Write-Output $false
        return
    }

    if ($null -eq $MailboxList) {
        Write-Host "ERROR: Email given has no access to any shared mailboxes, or does not exist." -ForegroundColor Red
        Write-Host "Returning to main menu..."
        Start-Sleep -Seconds 2
        Write-Output $false
        return
    }

    Write-Host "`nShared mailboxes that" $EmailAddress "has access to:`n" -ForegroundColor Yellow
    [int]$CurrentListNumber = 1
    foreach ($ListItem in $MailboxList) {
        Write-Host $CurrentListNumber":" $ListItem.Identity "("$ListItem.AccessRights")"
        $CurrentListNumber++
    }

    ## Feed list to variable if the function is being called to be used
    if ($true -eq $RunAsChildProcess) {
        Write-Output $MailboxList
        return
    }

    ## Prompt the user to do it all again or to return to the menu
    else {
        $UserInput_PostGet = $(Write-Host "`nWould you like to check another user's shared mailboxes? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

        ## Wait for answer
        while ("y","n" -notcontains $UserInput_PostGet) {
            $UserInput_PostGet = $(Write-Host "`nWould you like to check another user's shared mailboxes? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
        }

        if ($UserInput_PostGet -eq "y") {
            Get-ListOfSharedMailboxes -separateFunction $false
            return
        }

        else {
            Write-Host "`nReturning to main menu..."
            Clear-Host
        }
    }
}

function Edit-ListOfMailboxes {
    param (
        $InitialList
    )

    ## Ask user for list of mailboxes to keep
    $MailboxesToKeepList = $(Write-Host "`nSelect the mailboxes you require (by number, separated by ;, or write all): " -ForegroundColor yellow -NoNewLine; Read-Host)

    if ($MailboxesToKeepList -eq "all") {
        Write-Output $InitialList
        return
    }

    $MailboxesToKeepArray = $MailboxesToKeepList.Split(";")

    ## Create new list of mailboxes sans ones to keep
    Write-Host "`nBuilding modified list of mailboxes...`n"
    $ModifiedMailboxList = @()
    foreach ($ListItem in $InitialList) {
        foreach ($Index in $MailboxesToKeepArray) {
            #Write-Host "Checking if" $i.Identity "matches" $j"..."
            if ($ListItem -eq $InitialList[$Index-1]) {
                #Write-Host "Match located for" $i.Identity"." -ForegroundColor yellow
                $ModifiedMailboxList += $ListItem
                break
            }
        }
    }
    Write-Output $ModifiedMailboxList
    return
}

## Shared mailbox copy
function Copy-SharedMailboxes {
    Clear-Host
    $EmailAddress = $(Write-Host "`nPlease provide the email address of the user whose permissions are to be copied: " -ForegroundColor yellow -NoNewLine; Read-Host)
    $InitialList = Get-ListOfSharedMailboxes $true $EmailAddress
    if ($false -eq $InitialList) { return }
    $TrimmedList = Edit-ListOfMailboxes $InitialList

    Write-Host "`nList of mailboxes to copy to new account:"
    $TrimmedList | Format-Table

    $NewEmailAddress = $(Write-Host "Please provide the address to copy mailbox permissions to: " -ForegroundColor yellow -NoNewLine; Read-Host)
    $UserInput = $(Write-Host "`nWould you like to copy the above accounts from" $EmailAddress "to" $NewEmailAddress"? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

    ## Wait for answer
    While ("y","n" -notcontains $UserInput) {
        $UserInput = $(Write-Host "`nWould you like to copy the above accounts from" $EmailAddress "to" $NewEmailAddress"? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
    }

    ## If the user answers yes, proceed with rights migration
    If ($UserInput -eq 'y') {
        foreach ($ListItem in $TrimmedList) {
            Add-MailboxPermission -Identity $ListItem.Identity -User $NewEmailAddress -AccessRights $ListItem.AccessRights -InheritanceType All -Confirm:$false
        }

        Write-Host "`nMailboxes attached to" $NewEmailAddress":"
        $PostCopyList = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Get-EXOMailboxPermission |Select-Object Identity,User,AccessRights | `
        Where-Object {($_.user -like $NewEmailAddress)}

        $PostCopyList | Format-Table

        ## Prompt the user to do it all again or to return to the menu
        $UserInput_PostCopy = $(Write-Host "`nWould you like to copy another user's shared mailboxes? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

        ## Wait for answer
        while ("y","n" -notcontains $UserInput_PostCopy) {
            $UserInput_PostCopy = $(Write-Host "`nWould you like to copy another user's shared mailboxes? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
        }

        if ($UserInput_PostCopy -eq "y") {
            Copy-SharedMailboxes
            return
        }

        else {}
    }

    else {
        Write-Host "`nCancelling copy."
    }

    Write-Host "`nReturning to main menu..."
    return
}

## Shared mailbox permission migration
function Move-SharedMailboxes {
    Clear-Host
    $EmailAddress = $(Write-Host "`nPlease provide the email address of the user whose permissions are to be migrated: " -ForegroundColor yellow -NoNewLine; Read-Host)
    $InitialList = Get-ListOfSharedMailboxes $true $EmailAddress
    if ($false -eq $InitialList) { return }
    $TrimmedList = Edit-ListOfMailboxes $InitialList

    Write-Host "`nList of mailboxes to migrate to new account:"
    $TrimmedList | Format-Table

    $NewEmailAddress = $(Write-Host "Please provide the address to migrate mailbox permissions to: " -ForegroundColor yellow -NoNewLine; Read-Host)
    $UserInput = $(Write-Host "`nWould you like to move the above accounts from" $EmailAddress "to" $NewEmailAddress"? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

    ## Wait for answer
    While ("y","n" -notcontains $UserInput) {
        $UserInput = $(Write-Host "`nWould you like to move the above accounts from" $EmailAddress "to" $NewEmailAddress"? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
    }

    ## If the user answers yes, proceed with rights migration
    If ($UserInput -eq 'y') {
        foreach ($ListItem in $TrimmedList) {
            Remove-MailboxPermission -Identity $ListItem.Identity -User $EmailAddress -AccessRights $ListItem.AccessRights -InheritanceType All -Confirm:$false
            Add-MailboxPermission -Identity $ListItem.Identity -User $NewEmailAddress -AccessRights $ListItem.AccessRights -InheritanceType All -Confirm:$false
        }

        Write-Host "`nMailboxes attached to" $EmailAddress":"
        $PostMoveList = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Get-EXOMailboxPermission |Select-Object Identity,User,AccessRights | `
        Where-Object {($_.user -like $EmailAddress)}

        $PostMoveList | Format-Table

        Write-Host "Mailboxes attached to" $NewEmailAddress":"
        $PostMoveList2 = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Get-EXOMailboxPermission |Select-Object Identity,User,AccessRights | `
        Where-Object {($_.user -like $NewEmailAddress)}

        $PostMoveList2 | Format-Table

        ## Prompt the user to do it all again or to return to the menu
        $UserInput_PostMove = $(Write-Host "`nWould you like to move another user's shared mailboxes? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

        ## Wait for answer
        while ("y","n" -notcontains $UserInput_PostMove) {
            $UserInput_PostMove = $(Write-Host "`nWould you like to move another user's shared mailboxes? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
        }

        if ($UserInput_PostMove -eq "y") {
            Move-SharedMailboxes
            return
        }

        else {}
    }

    else {
        Write-Host "`nCancelling migration."
    }

    Write-Host "`nReturning to main menu..."
    return
}

## Shared mailbox removal
function Remove-SharedMailboxes {
    Clear-Host
    $EmailAddress = $(Write-Host "`nPlease provide the email address of the user whose permissions are to be removed: " -ForegroundColor yellow -NoNewLine; Read-Host)
    $InitialList = Get-ListOfSharedMailboxes $true $EmailAddress
    if ($false -eq $InitialList) { return }
    $TrimmedList = Edit-ListOfMailboxes $InitialList

    Write-Host "`nList of mailboxes to remove:"
    $TrimmedList | Format-Table

    $UserInput = $(Write-Host "Would you like to remove the above accounts from" $EmailAddress"? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

    ## Wait for answer
    While ("y","n" -notcontains $UserInput) {
        $UserInput = $(Write-Host "Would you like to remove the above accounts from" $EmailAddress"? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
    }

    ## If the user answers yes, proceed with rights migration
    If ($UserInput -eq 'y') {
        foreach ($ListItem in $TrimmedList) {
            Remove-MailboxPermission -Identity $ListItem.Identity -User $EmailAddress -AccessRights $ListItem.AccessRights -InheritanceType All -Confirm:$false
        }

        Write-Host "`nMailboxes attached to" $EmailAddress":"
        $PostRemoveList = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Get-EXOMailboxPermission |Select-Object Identity,User,AccessRights | `
        Where-Object {($_.user -like $EmailAddress)}

        $PostRemoveList | Format-Table

        ## Prompt the user to do it all again or to return to the menu
        $UserInput_PostRemove = $(Write-Host "`nWould you like to remove another user's shared mailboxes? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

        ## Wait for answer
        while ("y","n" -notcontains $UserInput_PostRemove) {
            $UserInput_PostRemove = $(Write-Host "`nWould you like to remove another user's shared mailboxes? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
        }

        if ($UserInput_PostRemove -eq "y") {
            Remove-SharedMailboxes
            return
        }

        else {}
    }

    else {
        Write-Host "`nCancelling removal."
    }

    Write-Host "`nReturning to main menu..."
    return
}

## List distribution groups the user is a member of
function Get-ListOfDistGroups($RunAsChildProcess, $EmailAddress) {
    $MailboxList = @()

    ## Get email address of user
    if ($false -eq $RunAsChildProcess) {
        Clear-Host
        $EmailAddress = $(Write-Host "`nPlease provide the user's email address: " -ForegroundColor yellow -NoNewLine; Read-Host)
    }

    ## Fetch list of shared mailboxes that the user has access to
    Write-Host "`nFetching list of distribution groups that"$EmailAddress" is a member of..."
    try {
        $MailboxList = Get-DistributionGroup | Where-Object { 
            (Get-DistributionGroupMember $_.Name | ForEach-Object {$_.PrimarySmtpAddress}) -contains "$EmailAddress"
        }
    }
    catch {
        Write-Host "ERROR: Most likely missing correct permissions to administrate Exchange Online." -ForegroundColor Red
        Write-Host "Returning to main menu..."
        Start-Sleep -Seconds 2
        Write-Output $false
        return
    }

    if ($null -eq $MailboxList) {
        Write-Host "ERROR: Email given is not a member of any distribution groups, or does not exist." -ForegroundColor Red
        Write-Host "Returning to main menu..."
        Start-Sleep -Seconds 2
        Write-Output $false
        return
    }

    Write-Host "`nDistribution groups that" $EmailAddress "is a member of:`n" -ForegroundColor Yellow
    [int]$CurrentListNumber = 1
    foreach ($ListItem in $MailboxList) {
        Write-Host $CurrentListNumber":" $ListItem.Identity
        $CurrentListNumber++
    }

    ## Feed list to variable if the function is being called to be used
    if ($true -eq $RunAsChildProcess) {
        Write-Output $MailboxList
        return
    }

    ## Prompt the user to do it all again or to return to the menu
    else {
        $UserInput_PostGet = $(Write-Host "`nWould you like to check another user's distribution groups? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

        ## Wait for answer
        while ("y","n" -notcontains $UserInput_PostGet) {
            $UserInput_PostGet = $(Write-Host "`nWould you like to check another user's distribution groups? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
        }

        if ($UserInput_PostGet -eq "y") {
            Get-ListOfSharedMailboxes -separateFunction $false
            return
        }

        else {
            Write-Host "`nReturning to main menu..."
            Clear-Host
        }
    }
}

## Distribution group copy
function Copy-DistGroups {
    Clear-Host
    $EmailAddress = $(Write-Host "`nPlease provide the email address of the user whose permissions are to be copied: " -ForegroundColor yellow -NoNewLine; Read-Host)
    $InitialList = Get-ListOfDistGroups $true $EmailAddress
    if ($false -eq $InitialList) { return }
    $TrimmedList = Edit-ListOfMailboxes $InitialList

    Write-Host "`nList of distribution groups to copy to new account:"
    $TrimmedList | Format-Table

    $NewEmailAddress = $(Write-Host "Please provide the address to copy mailbox permissions to: " -ForegroundColor yellow -NoNewLine; Read-Host)
    $UserInput = $(Write-Host "`nWould you like to copy the above groups from" $EmailAddress "to" $NewEmailAddress"? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

    ## Wait for answer
    While ("y","n" -notcontains $UserInput) {
        $UserInput = $(Write-Host "`nWould you like to copy the above groups from" $EmailAddress "to" $NewEmailAddress"? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
    }

    ## If the user answers yes, proceed with rights migration
    If ($UserInput -eq 'y') {
        foreach ($ListItem in $TrimmedList) {
            Add-DistributionGroupMember -Identity $ListItem.Identity -Member $NewEmailAddress -Confirm:$false
        }

        # Sleep required for dist groups to register
        Start-Sleep -Seconds 15

        Write-Host "`nMailboxes attached to" $NewEmailAddress":"
        $PostCopyList = = Get-DistributionGroup | Where-Object { 
            (Get-DistributionGroupMember $_.Name | ForEach-Object {$_.PrimarySmtpAddress}) -contains "$EmailAddress"
        }
        $PostCopyList | Format-Table

        ## Prompt the user to do it all again or to return to the menu
        $UserInput_PostCopy = $(Write-Host "`nWould you like to copy another user's distribution groups? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

        ## Wait for answer
        while ("y","n" -notcontains $UserInput_PostCopy) {
            $UserInput_PostCopy = $(Write-Host "`nWould you like to copy another user's distribution groups? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
        }

        if ($UserInput_PostCopy -eq "y") {
            Copy-SharedMailboxes
            return
        }

        else {}
    }

    else {
        Write-Host "`nCancelling copy."
    }

    Write-Host "`nReturning to main menu..."
    return
}

## Distribution group migration
function Move-DistGroups {
    Clear-Host
    $EmailAddress = $(Write-Host "`nPlease provide the email address of the user whose permissions are to be migrated: " -ForegroundColor yellow -NoNewLine; Read-Host)
    $InitialList = Get-ListOfDistGroups $true $EmailAddress
    if ($false -eq $InitialList) { return }
    $TrimmedList = Edit-ListOfMailboxes $InitialList

    Write-Host "`nList of distribution groups to move to new account:"
    $TrimmedList | Format-Table

    $NewEmailAddress = $(Write-Host "Please provide the address to move mailbox permissions to: " -ForegroundColor yellow -NoNewLine; Read-Host)
    $UserInput = $(Write-Host "`nWould you like to move the above groups from" $EmailAddress "to" $NewEmailAddress"? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

    ## Wait for answer
    While ("y","n" -notcontains $UserInput) {
        $UserInput = $(Write-Host "`nWould you like to move the above groups from" $EmailAddress "to" $NewEmailAddress"? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
    }

    ## If the user answers yes, proceed with rights migration
    If ($UserInput -eq 'y') {
        foreach ($ListItem in $TrimmedList) {
            Remove-DistributionGroupMember -Identity $ListItem.Identity -Member $NewEmailAddress -Confirm:$false
            Add-DistributionGroupMember -Identity $ListItem.Identity -Member $NewEmailAddress -Confirm:$false
        }

        # Sleep required for dist groups to register
        Start-Sleep -Seconds 15

        Write-Host "`nMailboxes attached to" $NewEmailAddress":"
        $PostMoveList = = Get-DistributionGroup | Where-Object { 
            (Get-DistributionGroupMember $_.Name | ForEach-Object {$_.PrimarySmtpAddress}) -contains "$EmailAddress"
        }
        $PostMoveList | Format-Table

        ## Prompt the user to do it all again or to return to the menu
        $UserInput_PostMove = $(Write-Host "`nWould you like to move another user's shared mailboxes? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

        ## Wait for answer
        while ("y","n" -notcontains $UserInput_PostMove) {
            $UserInput_PostMove = $(Write-Host "`nWould you like to move another user's shared mailboxes? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
        }

        if ($UserInput_PostMove -eq "y") {
            Move-SharedMailboxes
            return
        }

        else {}
    }

    else {
        Write-Host "`nCancelling migration."
    }

    Write-Host "`nReturning to main menu..."
    return
}

## Distribution group migration
function Remove-DistGroups {
    Clear-Host
    $EmailAddress = $(Write-Host "`nPlease provide the email address of the user whose permissions are to be removed: " -ForegroundColor yellow -NoNewLine; Read-Host)
    $InitialList = Get-ListOfDistGroups $true $EmailAddress
    if ($false -eq $InitialList) { return }
    $TrimmedList = Edit-ListOfMailboxes $InitialList

    Write-Host "`nList of distribution groups to remove membership of:"
    $TrimmedList | Format-Table

    $UserInput = $(Write-Host "`nWould you like to remove $EmailAddress as a member of the listed groups? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

    ## Wait for answer
    While ("y","n" -notcontains $UserInput) {
        $UserInput = $(Write-Host "`nWould you like to remove $EmailAddress as a member of the listed groups? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
    }

    ## If the user answers yes, proceed with rights migration
    If ($UserInput -eq 'y') {
        foreach ($ListItem in $TrimmedList) {
            Remove-DistributionGroupMember -Identity $ListItem.Identity -Member $EmailAddress -Confirm:$false
        }

        # Sleep required for dist groups to register
        Start-Sleep -Seconds 15

        Write-Host "`nGroups that"$EmailAddress" is a part of:"
        $PostRemoveList = = Get-DistributionGroup | Where-Object { 
            (Get-DistributionGroupMember $_.Name | ForEach-Object {$_.PrimarySmtpAddress}) -contains "$EmailAddress"
        }
        $PostRemoveList | Format-Table

        ## Prompt the user to do it all again or to return to the menu
        $UserInput_PostMove = $(Write-Host "`nWould you like to remove another user's distribution group memberships? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

        ## Wait for answer
        while ("y","n" -notcontains $UserInput_PostMove) {
            $UserInput_PostMove = $(Write-Host "`nWould you like to remove another user's distribution group memberships? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
        }

        if ($UserInput_PostMove -eq "y") {
            Move-SharedMailboxes
            return
        }

        else {}
    }

    else {
        Write-Host "`nCancelling migration."
    }

    Write-Host "`nReturning to main menu..."
    return
}

## Find mailbox by SMTP address
function Find-MailboxBySmtpAddress {
    Clear-Host
    $EmailAddress = $(Write-Host "`nPlease provide the email address you'd like to locate: " -ForegroundColor yellow -NoNewLine; Read-Host)

    ## Searches through each mailbox for a matching one with the provided SMTP address
    Get-EXOMailbox -ResultSize Unlimited -Filter {EmailAddresses -like $EmailAddress} | Select-Object DisplayName,PrimarySmtpAddress, `
    @{Name="EmailAddresses";Expression={($_.EmailAddresses | Where-Object {$_ -clike "smtp*"} | ForEach-Object {$_ -replace "smtp:",""}) -join ","}} | Sort-Object DisplayName

    $UserInput_PostFind = $(Write-Host "`nWould you like to locate another email address? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)

    ## Wait for answer
    while ("y","n" -notcontains $UserInput_PostFind) {
        $UserInput_PostFind = $(Write-Host "`nWould you like to locate another email address? (Y/N): " -ForegroundColor yellow -NoNewLine; Read-Host)
    }

    if ($UserInput_PostFind -eq "y") {
        Find-MailboxBySmtpAddress
        return
    }

    else {
        Write-Host "`nReturning to main menu..."
        Clear-Host
    }
}

## Build landing menu
function Show-Menu {
    Clear-Host
    Write-Host "================== Gabe's Exchange Online Toolbox ==================" -ForegroundColor Yellow
    
    Write-Host "`n1. Shared mailboxes"
    Write-Host "2. Distribution groups"
    Write-Host "3. Other"
    Write-Host "Q. Quit"
}

function Show-SubMenu ($Type) {
    Clear-Host
    Write-Host "$Type" -ForegroundColor Yellow

    Write-Host "`n1. List"
    Write-Host "2. Copy"
    Write-Host "3. Move"
    Write-Host "4. Delete"
    Write-Host "B. Back"
}

function Show-OtherMenu {
    Clear-Host
    Write-Host "`n1. Find mailbox by email address (currently broken)"
    Write-Host "B. Back"
}

do {
    Show-Menu
    $MenuSelection = Read-Host "`nPlease make a selection"
    switch ($MenuSelection) {
        '1' {
            do {
                Show-SubMenu ("Shared Mailboxes")
                $MenuSelection_Shared = Read-Host "`nPlease make a selection"
                switch ($MenuSelection_Shared) {
                    '1' {
                        Get-ListOfSharedMailboxes $false
                    } '2' {
                        Copy-SharedMailboxes
                    } '3' {
                        Move-SharedMailboxes
                    } '4' {
                        Remove-SharedMailboxes
                    }
                }
            } until ($MenuSelection_Shared -eq 'b')
        } '2' {
            do {
                Show-SubMenu ("Distribution Groups")
                $MenuSelection_Dist = Read-Host "`nPlease make a selection"
                switch ($MenuSelection_Dist) {
                    '1' {
                        Get-ListOfDistGroups $false
                    } '2' {
                        Copy-DistGroups
                    } '3' {
                        Move-DistGroups
                    } '4' {
                        Remove-DistGroups
                    }
                }
            } until ($MenuSelection_Dist -eq 'b')
        } '3' {
            do {
                Show-OtherMenu
                $MenuSelection_Other = Read-Host "`nPlease make a selection"
                switch ($MenuSelection_Other) {
                    '1' {
                        Find-MailboxBySmtpAddress
                    }
                }
            } until ($MenuSelection_Other -eq 'b')
        }
    }
    pause
} until ($MenuSelection-eq 'q')

Write-Host "`nDisconnecting from Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false
Clear-Host
exit