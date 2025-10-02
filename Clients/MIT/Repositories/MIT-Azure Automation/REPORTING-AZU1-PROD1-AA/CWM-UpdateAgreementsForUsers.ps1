<#

Mangano IT - ConnectWise Manage - Update Agreements to Match Current User Counts
Created by: Gabriel Nugent
Version: 1.8

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
    $ApiSecrets = $null,
    [bool]$Test = $false,
    $TestAgreementId = 871,
    $TestAdditionId = 3589
)

## SCRIPT VARIABLES ##

$AgreementType = '1 Managed Services Blue'
[string]$Log = ''
$Date = Get-Date -Format 'dd/MM/yyyy'

# Get ConnectWise secrets if not provided
if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

# Establish standard ticket arguments
$DefaultArguments = @{
    Summary = "Update user count for MSA - $Date"
    CompanyId = 0
    BoardName = 'Recurring'
    StatusName = 'Pre-Process'
    TypeName = 'Request'
    SubtypeName = 'Client Management'
    ItemName = 'Review'
    PriorityName = 'P5 Routine Response'
    InitialDescriptionNote = @"
This ticket is for reviewing the current user count at a client, and updating it to match an agreed upon reference point (e.g. a license count).

Once automation has checked the client, it will update this ticket to say whether or not the license count has been updated.

If the client does not have an Azure environment, the true-up will need to be done manually by checking new and exit user tickets for the month.
"@
}

# Establish standard note arguments
$DefaultNoteArguments = @{
    TicketId = 0
    Text = ''
    ResolutionFlag = $true
    ApiSecrets = $ApiSecrets
}

# SKU part numbers
$SkuPartNumber = @{
    M365BB = 'O365_BUSINESS_ESSENTIALS'
    M365BP = 'SPB'
    M365BP_Alt = 'O365_BUSINESS_PREMIUM'
    M365E5 = 'SPE_E5'
    O365E1 = 'STANDARDPACK'
    O365E3 = 'ENTERPRISEPACK'
}

# Invalid display names
$InvalidNames = @(
    '*admin*'
    '*service account*'
    '*mangano it*'
    '*test *'
    '*power bi*'
    '*reception*'
    '*warehouse*'
)

# Hashtable of all MSA clients with license that represents their standard user
# If details are blank, there is no Azure tenancy to match against - intentional
$Clients = @(
    @{
        Name = 'All Terrain Warriors'
        TenantUrl = 'allterrainwarriors1.onmicrosoft.com'
        LicenseSku = $SkuPartNumber.M365BP
    },
    @{
        Name = 'Base Architecture'
        TenantUrl = 'basearchitecture.onmicrosoft.com'
        LicenseSku = $SkuPartNumber.M365BP
    },
    @{
        Name = 'Departure Point'
        TenantUrl = 'departurepoint.com.au'
        LicenseSku = $SkuPartNumber.M365BP
    },
    @{
        Name = 'Kern Group'
        TenantUrl = 'nathanwood.onmicrosoft.com'
        LicenseSku = $SkuPartNumber.M365BB + ';' + $SkuPartNumber.M365BP_Alt
    },
    @{
        Name = 'm3architecture Pty Ltd'
        TenantUrl = 'm3architecture.onmicrosoft.com'
        LicenseSku = $SkuPartNumber.M365BP
    },
    @{
        Name = 'Olympic Fire'
        TenantUrl = 'olympicfire1.onmicrosoft.com'
        LicenseSku = $SkuPartNumber.O365E3 + ';' + $SkuPartNumber.O365E1
    },
    @{
        Name = 'Seasons Living Australia Pty Ltd'
        TenantUrl = 'seasonsagedcare.com.au'
        LicenseSku = $SkuPartNumber.M365BP
    },
    @{
        Name = 'Sherrin Rentals'
        TenantUrl = 'sherrinequipment.onmicrosoft.com'
        LicenseSku = $SkuPartNumber.M365BP
    },
    @{
        Name = 'Redland Investment Corporation'
        TenantUrl = 'redlandinvestcorp.com.au'
        LicenseSku = $SkuPartNumber.M365BP
    },
    @{
        Name = 'Engineering Applications Pty Ltd'
        TenantUrl = 'engapp.onmicrosoft.com'
        LicenseSku = $SkuPartNumber.M365BP
    },
    @{
        Name = 'Elston'
        TenantUrl = 'elstongroup.onmicrosoft.com'
        LicenseSku = $SkuPartNumber.M365E5
    },
    @{
        Name = 'Pinata Farms Operations Pty Ltd'
        TenantUrl = 'pinata.com.au'
        LicenseSku = $SkuPartNumber.M365BB + ';' + $SkuPartNumber.M365BP
    },
    @{
        Name = 'The Eye Health Centre'
        TenantUrl = 'qldeyesurgery.com'
        LicenseSku = $SkuPartNumber.M365BP
    },
    @{
        Name = 'Argent Australia'
        TenantUrl = 'argentaust1.onmicrosoft.com'
        LicenseSku = $SkuPartNumber.M365BP
    },
    @{
        Name = 'OPEC Systems Pty Ltd'
        TenantUrl = 'opecsystems.onmicrosoft.com'
        LicenseSku = $SkuPartNumber.M365E5
    }
)

## GRAB AGREEMENTS ##

try {
    $Log += "Fetching all '$AgreementType' agreements...`n"
    $Agreements = .\CWM-FindAgreementsAndUsers.ps1 -AgreementTypes $AgreementType -ApiSecrets $ApiSecrets | ConvertFrom-Json
    $Log += "SUCCESS: Fetched all '$AgreementType' agreements.`n`n"
} catch {
    $Log += "ERROR: Unable to fetch all '$AgreementType' agreements.`nERROR DETAILS: " + $_
}

## GRAB USER DETAILS PER AGREEMENT ##

foreach ($Agreement in $Agreements) {
    # Clear/set temporary variables
    $Client = $null
    $Company = $Agreement.Company
    $CompanyId = $Agreement.CompanyId
    $LicensesAndUsers = $null
    $LicensedUsers = $null
    $AgreementId = $Agreement.Id
    $AdditionId = 0
    $Product = ''
    $UserList = ''
    $SkippedUserList = ''
    $UserCount = 0
    $UserCount_Other = 0
    $NewUserCount = 0
    $TicketArguments = $DefaultArguments
    $TicketArguments.CompanyId = $CompanyId
    $TicketNoteArguments = $DefaultNoteArguments

    # Get client details from hashtable
    $Client = $Clients | Where-Object { $_.Name -eq $Company }

    # Skip if client not found
    if ($null -eq $Client) {
        $Log += "INFO: No details found for $Company.`n`n"
        $UnmatchedCompanies += "$Company`n"
        continue
    }

    # Create ticket
    $TicketId = .\New-CWTicketAndInitialNote.ps1 @TicketArguments
    $TicketNoteArguments.TicketId = $TicketId

    # Add note for manual check if tenant URL not provided
    if ($Client.TenantUrl -eq '' -or $Client.LicenseSku -eq '') {
        $Log += "INFO: No tenant URL/license SKU provided for $Company. A manual check will need to be performed.`n`n"
        $TicketNoteArguments = $DefaultNoteArguments
        $TicketNoteArguments.Text = "No tenant URL/license SKU is present for $Company, which means there is no Azure environment to reference.`n`nA manual check will need to be performed."
        .\CWM-AddTicketNote.ps1 @TicketNoteArguments | Out-Null
        continue
    }

    # Look for additions that define standard users
    foreach ($Addition in $Agreement.UserAdditions) {
        if (($Addition.Product -eq 'AGMT-UserCentric' -or $Addition.Product -eq 'AGMT-UserStandard') -and $UserCount -eq 0) {
            $Product = $Addition.Product
            $AdditionId = $Addition.AdditionId
            $UserCount = [int]$Addition.Quantity
            $Log += "INFO: User count for $Company located - $UserCount.`n"
            $TicketNoteArguments = $DefaultNoteArguments
            $TicketNoteArguments.Text = "User count for $Company located - $UserCount."
            .\CWM-AddTicketNote.ps1 @TicketNoteArguments | Out-Null
            continue
        } elseif ($Company -eq 'OPEC Systems Pty Ltd' -and $Addition.Product -eq 'AGMT-UserBasic') {
            # Skip reduction for basic users at OPEC - they are on a different license
            continue
        } elseif ($Addition.Quantity -ne 0) {
            $UserCount_Other += [int]$Addition.Quantity
            $Log += "INFO: Non-standard user count for $Company located. Alternate user count updated - $UserCount_Other.`n"
            $TicketNoteArguments = $DefaultNoteArguments
            $TicketNoteArguments.Text = "Non-standard user count for $Company located. Alternate user count updated - $UserCount_Other."
            .\CWM-AddTicketNote.ps1 @TicketNoteArguments | Out-Null
        }
    }

    # Get license details
    foreach ($Sku in $Client.LicenseSku.Split(';')) {
        if ($Sku -ne '') {
            $LicensesAndUsers = .\AAD-FindLicensesAndUsers.ps1 -TenantUrl $Client.TenantUrl -SkuPartNumber $Sku | ConvertFrom-Json
            $LicensedUsers += $LicensesAndUsers.Users
        }
    }
    
    # Organise user list to increase count based on valid users
    :Outer foreach ($User in $LicensedUsers) {
        $DisplayName = $User.displayName
        $DisplayNameLower = $DisplayName.ToLower();
        $UserPrincipalName = $User.userPrincipalName
        foreach ($InvalidName in $InvalidNames) {
            if ($DisplayNameLower -like $InvalidName) {
                $SkippedUserList += "Display name: $DisplayName`n"
                $SkippedUserList += "User principal name: $UserPrincipalName`n`n"
                continue Outer
            }
        }
        $UserList += "Display name: $DisplayName`n"
        $UserList += "User principal name: $UserPrincipalName`n`n"
        $NewUserCount += 1
    }

    # Reduce user count by alt user count
    $NewUserCount -= $UserCount_Other

    # Add note with approved list of users
    $InternalNoteArguments = @{
        TicketId = $TicketId
        Text = "List of users considered valid for the count:`n`n$UserList"
        InternalFlag = $true
        ApiSecrets = $ApiSecrets
    }
    .\CWM-AddTicketNote.ps1 @InternalNoteArguments | Out-Null
    if ($SkippedUserList -ne '') {
        $InternalNoteArguments.Text = "List of users considered invalid for the count:`n`n$SkippedUserList"
        .\CWM-AddTicketNote.ps1 @InternalNoteArguments | Out-Null
    }

    # Add note with final confirmation of user count
    $TicketNoteArguments = $DefaultNoteArguments
    if ($UserCount -lt $NewUserCount) {
        $TicketNoteArguments.Text = "Current user count is $UserCount.`n`n"
        $TicketNoteArguments.Text += "The new user count for addition $Product should be $NewUserCount (licensed real users reduced by $UserCount_Other).`n`n"
        $TicketNoteArguments.Text += "Please review the agreements and update as necessary."
        .\CWM-AddTicketNote.ps1 @TicketNoteArguments | Out-Null

        ## AUTOMATED AGREEMENT UPDATING ##

        <#
        # Add sleep so that ticket notes appear in the right order
        Start-Sleep -Seconds 30

        # Set agreement details to test agreement and addition
        if ($Test) {
            $AgreementId = $TestAgreementId
            $AdditionId = $TestAdditionId
        }

        # Update addition with new user count
        $Arguments = @{
            AgreementId = $AgreementId
            AdditionId = $AdditionId
            Value = $NewUserCount
            ApiSecrets = $ApiSecrets
        }
        $UpdateAddition = .\CWM-UpdateAddition.ps1 @Arguments | ConvertFrom-Json
        $Log += $UpdateAddition.Log + "`n`n"

        # Add note if successful
        if ($UpdateAddition.Result) {
            $TicketNoteArguments = $DefaultNoteArguments
            $TicketNoteArguments.Text = "The agreement has been updated to $NewUserCount."
            .\CWM-AddTicketNote.ps1 @TicketNoteArguments | Out-Null
        } #>
    } elseif ($UserCount -gt $NewUserCount) {
        $TicketNoteArguments.Text = "Current user count is $UserCount.`n`n"
        $TicketNoteArguments.Text += "The new user count for addition $Product should be $NewUserCount (licensed real users reduced by $UserCount_Other).`n`n"
        $TicketNoteArguments.Text += "Please review this agreement manually before applying this reduction."
        .\CWM-AddTicketNote.ps1 @TicketNoteArguments | Out-Null
    } elseif ($UserCount -eq $NewUserCount) {
        $TicketNoteArguments.Text = "Current user count is $UserCount.`n`n"
        $TicketNoteArguments.Text += "Automation has confirmed that the count for addition $Product should be $NewUserCount (licensed real users reduced by $UserCount_Other).`n`n"
        $TicketNoteArguments.Text += "The agreement will not be updated."
        .\CWM-AddTicketNote.ps1 @TicketNoteArguments | Out-Null
    }
}

## SEND DETAILS TO FLOW ##

$Output = @{
    UnmatchedCompanies = $UnmatchedCompanies
    Log = $Log
}

Write-Output $Output | ConvertTo-Json -Depth 100