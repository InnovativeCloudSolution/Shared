<#	
	.NOTES
	===========================================================================
     Version:       1.0.5
	 Updated on:   	8/14/2018
	 Created by:   	/u/TheLazyAdministrator
     Contributors:  /u/ascIVV, /u/jmn_lab, /u/nothingpersonalbro
	===========================================================================

        AzureAD  Module is required
            Install-Module -Name AzureAD
            https://www.powershellgallery.com/packages/azuread/
        ReportHTML Moduile is required
            Install-Module -Name ReportHTML
            https://www.powershellgallery.com/packages/ReportHTML/

        UPDATES
        1.0.5
            /u/ascIVV: Added the following:
                - Admin Tab
                    - Privileged Role Administrators
                    - Exchange Administrators
                    - User Account Administrators
                    - Tech Account Restricted Exchange Admin Role
                    - SharePoint Administrators
                    - Skype Administrators
                    - CRM Service Administrators
                    - Power BI Administrators
                    - Service Support Administrators
                    - Billing Administrators
            /u/TheLazyAdministrator
                - Cleaned up formatting
                - Error Handling for $Null obj
                - Console status
                - Windows Defender ATP SKU
        

	.DESCRIPTION
		Generate an interactive HTML report on your Office 365 tenant. Report on Users, Tenant information, Groups, Policies, Contacts, Mail Users, Licenses and more!
    
    .Link
        Original: http://thelazyadministrator.com/2018/06/22/create-an-interactive-html-report-for-office-365-with-powershell/
#>
#########################################
#                                       #
#            VARIABLES                  #
#                                       #
#########################################

$GraphclientId = "e05b08c8-78aa-4bc2-9d3c-f99bca9bf5f6"
$GraphclientSecret = "6a._9At5Yh3~-Bg1-j1iaHqKRR_6TsiEOm"
$SPclientId = "5e3bb8a5-5c48-4715-8e0f-0333df9651bc"
$SPclientSecret = "pN18Q~eatzFLx1R_AL1zelEfxL5rbHa_E1EVqb8n"
$SPScope = "https://qrlcomau.sharepoint.com/.default"
$ExoclientId = "df81cdc6-ef98-420d-84fd-51bfe0e7992c"
$ExoclientSecret = "vy48Q~XSn4YdN-.eTVxLVLC55LNC~8CVHq8kYb63"
$ExoScope = "https://outlook.office365.com/.default"
$tenantId = "qrlcomau.onmicrosoft.com"
$graphBase = "https://graph.microsoft.com/v1.0"
$graphBaseBeta = "https://graph.microsoft.com/beta"
$CompanyLogo = "https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/Mangano IT (Logo-Horizontal).png"
$RightLogo = "https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/Mangano IT (Logo-Horizontal).png"
$ReportSavePath = "C:\Scripts\Repositories\PowerShell Scripts\MIT\DEV\ASIO\Reports\365\"
$LicenseFilter = "9000"

$Table = New-Object 'System.Collections.Generic.List[System.Object]'
$LicenseTable = New-Object 'System.Collections.Generic.List[System.Object]'
$UserTable = New-Object 'System.Collections.Generic.List[System.Object]'
$SharedMailboxTable = New-Object 'System.Collections.Generic.List[System.Object]'
$GroupTypetable = New-Object 'System.Collections.Generic.List[System.Object]'
$IsLicensedUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ContactTable = New-Object 'System.Collections.Generic.List[System.Object]'
$MailUser = New-Object 'System.Collections.Generic.List[System.Object]'
$ContactMailUserTable = New-Object 'System.Collections.Generic.List[System.Object]'
$RoomTable = New-Object 'System.Collections.Generic.List[System.Object]'
$EquipTable = New-Object 'System.Collections.Generic.List[System.Object]'
$StrongPasswordTable = New-Object 'System.Collections.Generic.List[System.Object]'
$CompanyInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
$DomainTable = New-Object 'System.Collections.Generic.List[System.Object]'

$Sku = @{
    "O365_BUSINESS_ESSENTIALS"                                     = "Office 365 Business Essentials"
    "O365_BUSINESS_PREMIUM"                                        = "Office 365 Business Premium"
    "DESKLESSPACK"                                                 = "Office 365 (Plan K1)"
    "DESKLESSWOFFPACK"                                             = "Office 365 (Plan K2)"
    "LITEPACK"                                                     = "Office 365 (Plan P1)"
    "EXCHANGESTANDARD"                                             = "Office 365 Exchange Online Only"
    "STANDARDPACK"                                                 = "Enterprise Plan E1"
    "STANDARDWOFFPACK"                                             = "Office 365 (Plan E2)"
    "ENTERPRISEPACK"                                               = "Enterprise Plan E3"
    "ENTERPRISEPACKLRG"                                            = "Enterprise Plan E3"
    "ENTERPRISEWITHSCAL"                                           = "Enterprise Plan E4"
    "STANDARDPACK_STUDENT"                                         = "Office 365 (Plan A1) for Students"
    "STANDARDWOFFPACKPACK_STUDENT"                                 = "Office 365 (Plan A2) for Students"
    "ENTERPRISEPACK_STUDENT"                                       = "Office 365 (Plan A3) for Students"
    "ENTERPRISEWITHSCAL_STUDENT"                                   = "Office 365 (Plan A4) for Students"
    "STANDARDPACK_FACULTY"                                         = "Office 365 (Plan A1) for Faculty"
    "STANDARDWOFFPACKPACK_FACULTY"                                 = "Office 365 (Plan A2) for Faculty"
    "ENTERPRISEPACK_FACULTY"                                       = "Office 365 (Plan A3) for Faculty"
    "ENTERPRISEWITHSCAL_FACULTY"                                   = "Office 365 (Plan A4) for Faculty"
    "ENTERPRISEPACK_B_PILOT"                                       = "Office 365 (Enterprise Preview)"
    "STANDARD_B_PILOT"                                             = "Office 365 (Small Business Preview)"
    "VISIOCLIENT"                                                  = "Visio Pro Online"
    "POWER_BI_ADDON"                                               = "Office 365 Power BI Addon"
    "POWER_BI_INDIVIDUAL_USE"                                      = "Power BI Individual User"
    "POWER_BI_STANDALONE"                                          = "Power BI Stand Alone"
    "POWER_BI_STANDARD"                                            = "Power-BI Standard"
    "PROJECTESSENTIALS"                                            = "Project Lite"
    "PROJECTCLIENT"                                                = "Project Professional"
    "PROJECTONLINE_PLAN_1"                                         = "Project Online"
    "PROJECTONLINE_PLAN_2"                                         = "Project Online and PRO"
    "ProjectPremium"                                               = "Project Online Premium"
    "ECAL_SERVICES"                                                = "ECAL"
    "EMS"                                                          = "Enterprise Mobility Suite"
    "RIGHTSMANAGEMENT_ADHOC"                                       = "Windows Azure Rights Management"
    "MCOMEETADV"                                                   = "PSTN conferencing"
    "SHAREPOINTSTORAGE"                                            = "SharePoint storage"
    "PLANNERSTANDALONE"                                            = "Planner Standalone"
    "CRMIUR"                                                       = "CMRIUR"
    "BI_AZURE_P1"                                                  = "Power BI Reporting and Analytics"
    "INTUNE_A"                                                     = "Windows Intune Plan A"
    "PROJECTWORKMANAGEMENT"                                        = "Office 365 Planner Preview"
    "ATP_ENTERPRISE"                                               = "Exchange Online Advanced Threat Protection"
    "EQUIVIO_ANALYTICS"                                            = "Office 365 Advanced eDiscovery"
    "AAD_BASIC"                                                    = "Azure Active Directory Basic"
    "RMS_S_ENTERPRISE"                                             = "Azure Active Directory Rights Management"
    "AAD_PREMIUM"                                                  = "Azure Active Directory Premium"
    "MFA_PREMIUM"                                                  = "Azure Multi-Factor Authentication"
    "STANDARDPACK_GOV"                                             = "Microsoft Office 365 (Plan G1) for Government"
    "STANDARDWOFFPACK_GOV"                                         = "Microsoft Office 365 (Plan G2) for Government"
    "ENTERPRISEPACK_GOV"                                           = "Microsoft Office 365 (Plan G3) for Government"
    "ENTERPRISEWITHSCAL_GOV"                                       = "Microsoft Office 365 (Plan G4) for Government"
    "DESKLESSPACK_GOV"                                             = "Microsoft Office 365 (Plan K1) for Government"
    "ESKLESSWOFFPACK_GOV"                                          = "Microsoft Office 365 (Plan K2) for Government"
    "EXCHANGESTANDARD_GOV"                                         = "Microsoft Office 365 Exchange Online (Plan 1) only for Government"
    "EXCHANGEENTERPRISE_GOV"                                       = "Microsoft Office 365 Exchange Online (Plan 2) only for Government"
    "SHAREPOINTDESKLESS_GOV"                                       = "SharePoint Online Kiosk"
    "EXCHANGE_S_DESKLESS_GOV"                                      = "Exchange Kiosk"
    "RMS_S_ENTERPRISE_GOV"                                         = "Windows Azure Active Directory Rights Management"
    "OFFICESUBSCRIPTION_GOV"                                       = "Office ProPlus"
    "MCOSTANDARD_GOV"                                              = "Lync Plan 2G"
    "SHAREPOINTWAC_GOV"                                            = "Office Online for Government"
    "SHAREPOINTENTERPRISE_GOV"                                     = "SharePoint Plan 2G"
    "EXCHANGE_S_ENTERPRISE_GOV"                                    = "Exchange Plan 2G"
    "EXCHANGE_S_ARCHIVE_ADDON_GOV"                                 = "Exchange Online Archiving"
    "EXCHANGE_S_DESKLESS"                                          = "Exchange Online Kiosk"
    "SHAREPOINTDESKLESS"                                           = "SharePoint Online Kiosk"
    "SHAREPOINTWAC"                                                = "Office Online"
    "YAMMER_ENTERPRISE"                                            = "Yammer for the Starship Enterprise"
    "EXCHANGE_L_STANDARD"                                          = "Exchange Online (Plan 1)"
    "MCOLITE"                                                      = "Lync Online (Plan 1)"
    "SHAREPOINTLITE"                                               = "SharePoint Online (Plan 1)"
    "OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ"                           = "Office ProPlus"
    "EXCHANGE_S_STANDARD_MIDMARKET"                                = "Exchange Online (Plan 1)"
    "MCOSTANDARD_MIDMARKET"                                        = "Lync Online (Plan 1)"
    "SHAREPOINTENTERPRISE_MIDMARKET"                               = "SharePoint Online (Plan 1)"
    "OFFICESUBSCRIPTION"                                           = "Office ProPlus"
    "YAMMER_MIDSIZE"                                               = "Yammer"
    "DYN365_ENTERPRISE_PLAN1"                                      = "Dynamics 365 Customer Engagement Plan Enterprise Edition"
    "ENTERPRISEPREMIUM_NOPSTNCONF"                                 = "Enterprise E5 (without Audio Conferencing)"
    "ENTERPRISEPREMIUM"                                            = "Enterprise E5 (with Audio Conferencing)"
    "MCOSTANDARD"                                                  = "Skype for Business Online Standalone Plan 2"
    "PROJECT_MADEIRA_PREVIEW_IW_SKU"                               = "Dynamics 365 for Financials for IWs"
    "STANDARDWOFFPACK_IW_STUDENT"                                  = "Office 365 Education for Students"
    "STANDARDWOFFPACK_IW_FACULTY"                                  = "Office 365 Education for Faculty"
    "EOP_ENTERPRISE_FACULTY"                                       = "Exchange Online Protection for Faculty"
    "EXCHANGESTANDARD_STUDENT"                                     = "Exchange Online (Plan 1) for Students"
    "OFFICESUBSCRIPTION_STUDENT"                                   = "Office ProPlus Student Benefit"
    "STANDARDWOFFPACK_FACULTY"                                     = "Office 365 Education E1 for Faculty"
    "STANDARDWOFFPACK_STUDENT"                                     = "Microsoft Office 365 (Plan A2) for Students"
    "DYN365_FINANCIALS_BUSINESS_SKU"                               = "Dynamics 365 for Financials Business Edition"
    "DYN365_FINANCIALS_TEAM_MEMBERS_SKU"                           = "Dynamics 365 for Team Members Business Edition"
    "FLOW_FREE"                                                    = "Microsoft Flow Free"
    "POWER_BI_PRO"                                                 = "Power BI Pro"
    "O365_BUSINESS"                                                = "Office 365 Business"
    "DYN365_ENTERPRISE_SALES"                                      = "Dynamics Office 365 Enterprise Sales"
    "RIGHTSMANAGEMENT"                                             = "Rights Management"
    "PROJECTPROFESSIONAL"                                          = "Project Professional"
    "VISIOONLINE_PLAN1"                                            = "Visio Online Plan 1"
    "EXCHANGEENTERPRISE"                                           = "Exchange Online Plan 2"
    "DYN365_ENTERPRISE_P1_IW"                                      = "Dynamics 365 P1 Trial for Information Workers"
    "DYN365_ENTERPRISE_TEAM_MEMBERS"                               = "Dynamics 365 For Team Members Enterprise Edition"
    "CRMSTANDARD"                                                  = "Microsoft Dynamics CRM Online Professional"
    "EXCHANGEARCHIVE_ADDON"                                        = "Exchange Online Archiving For Exchange Online"
    "EXCHANGEDESKLESS"                                             = "Exchange Online Kiosk"
    "SPZA_IW"                                                      = "App Connect"
    "WINDOWS_STORE"                                                = "Windows Store for Business"
    "MCOEV"                                                        = "Microsoft Phone System"
    "VIDEO_INTEROP"                                                = "Polycom Skype Meeting Video Interop for Skype for Business"
    "SPE_E5"                                                       = "Microsoft 365 E5"
    "SPE_E3"                                                       = "Microsoft 365 E3"
    "ATA"                                                          = "Advanced Threat Analytics"
    "MCOPSTN2"                                                     = "Domestic and International Calling Plan"
    "FLOW_P1"                                                      = "Microsoft Flow Plan 1"
    "FLOW_P2"                                                      = "Microsoft Flow Plan 2"
    "WIN_DEF_ATP"                                                  = "Windows Defender ATP"
    "POWER_BI_PRO_CE"                                              = "Power BI Pro (Capacity-Based)"
    "Microsoft_365_Copilot"                                        = "Microsoft 365 Copilot"
    "Microsoft_365_Business_Premium_Donation_(Non_Profit_Pricing)" = "Microsoft 365 Business Premium Donation (Non-Profit)"

}

function Get-AccessToken {
    param (
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)][string]$ClientSecret,
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $false)][string]$Scope = "https://graph.microsoft.com/.default"
    )

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = $Scope
        grant_type    = "client_credentials"
    }

    try {
        $response = Invoke-RestMethod -Method POST -Uri $tokenEndpoint -Body $body -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    }
    catch {
        Write-Error "Failed to get access token for scope [$Scope]: $($_.Exception.Message)"
        return $null
    }
}

function Invoke-ApiCall {
    param (
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        $Body = $null,
        [hashtable]$Headers = @{},
        [int]$RetryCount = 3
    )

    $isSharePoint = $Uri -match "sharepoint\.com"

    $defaultHeaders = @{
        Authorization = "Bearer $AccessToken"
        Accept        = if ($isSharePoint) { "application/json;odata=verbose" } else { "application/json" }
        "Content-Type" = if ($isSharePoint) { "application/json;odata=verbose" } else { "application/json" }
    }

    $combinedHeaders = $defaultHeaders + $Headers

    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            if ($Method -eq "GET") {
                return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $combinedHeaders
            }
            else {
                $jsonBody = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 }
                return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $combinedHeaders -Body $jsonBody
            }
        }
        catch {
            if ($i -eq $RetryCount) { throw }
            Start-Sleep -Seconds (3 * $i)
        }
    }
}

function Get-RoleMembers {
    param (
        [string]$roleDisplayName,
        [string]$Uri,
        [string]$AccessToken
    )

    $safeName = ($roleDisplayName -replace '\s', '') -replace '[^a-zA-Z0-9]', ''
    $table = New-Object System.Collections.Generic.List[Object]

    $roleList = Invoke-ApiCall -Method "GET" -Uri "$Uri/directoryRoles" -AccessToken $AccessToken
    $role = $roleList.value | Where-Object { $_.displayName -eq $roleDisplayName }

    if ($role) {
        $members = Invoke-ApiCall -Method "GET" -Uri "$Uri/directoryRoles/$($role.id)/members" -AccessToken $AccessToken

        foreach ($member in $members.value) {
            switch ($member.'@odata.type') {
                "#microsoft.graph.user" {
                    $user = Invoke-ApiCall -Method "GET" -Uri "$Uri/users/$($member.id)" -AccessToken $AccessToken
                    $methods = Invoke-ApiCall -Method "GET" -Uri "$Uri/users/$($member.id)/authentication/methods" -AccessToken $AccessToken
                    $mfaStatus = if ($methods.value.Count -gt 0) { "Enabled" } else { "Disabled" }
                    $email = if ($user.mail) { $user.mail } else { $user.userPrincipalName }
                    $licensed = $user.assignedLicenses.Count -gt 0

                    $obj = [PSCustomObject]@{
                        'Name'           = $user.displayName
                        'MFA Status'     = $mfaStatus
                        'Is Licensed'    = $licensed
                        'E-Mail Address' = $email
                    }
                }

                "#microsoft.graph.group" {
                    $group = Invoke-ApiCall -Method "GET" -Uri "$Uri/groups/$($member.id)" -AccessToken $AccessToken
                    $obj = [PSCustomObject]@{
                        'Name'           = $group.displayName
                        'MFA Status'     = "N/A (Group)"
                        'Is Licensed'    = "N/A"
                        'E-Mail Address' = $group.mail
                    }
                }

                "#microsoft.graph.servicePrincipal" {
                    $sp = Invoke-ApiCall -Method "GET" -Uri "$Uri/servicePrincipals/$($member.id)" -AccessToken $AccessToken
                    $obj = [PSCustomObject]@{
                        'Name'           = $sp.displayName
                        'MFA Status'     = "N/A (Service Principal)"
                        'Is Licensed'    = "N/A"
                        'E-Mail Address' = $sp.appId
                    }
                }

                default {
                    continue
                }
            }

            $table.Add($obj)
        }
    }

    return @{ ($safeName + "Table") = $table }
}

$graphToken = Get-AccessToken -ClientId $GraphClientId -ClientSecret $GraphClientSecret -TenantId $TenantId
$sptoken = Get-AccessToken -ClientId $GraphClientId -ClientSecret $GraphClientSecret -TenantId $TenantId -Scope $SPScope
$exoToken = Get-AccessToken -ClientId $ExoClientId -ClientSecret $ExoClientSecret -TenantId $TenantId -Scope $ExoScope
Connect-ExchangeOnline -AccessToken $exoToken -Organization $TenantId -ShowBanner:$false

$AllUsers = @()
$nextLink = "$graphBaseBeta/users?`$top=50&`$select=id,displayName,userPrincipalName,assignedLicenses,passwordPolicies,accountEnabled,proxyAddresses,createdDateTime,userType,creationType,externalUserState,signInActivity"
while ($nextLink) {
    $response = Invoke-ApiCall -Method "GET" -Uri $nextLink -AccessToken $graphToken
    $AllUsers += $response.value
    $nextLink = $response.'@odata.nextLink'
}

$Groups = @()
$nextLink = "$graphBase/groups`?$top=999"
while ($nextLink) {
    $resp = Invoke-ApiCall -Method "GET" -Uri $nextLink -AccessToken $graphToken
    $Groups += $resp.value
    $nextLink = $resp.'@odata.nextLink'
}

$AllSites = @()
$nextLink = "$graphBase/sites/getAllSites"
while ($nextLink) {
    $resp = Invoke-ApiCall -Method "GET" -Uri $nextLink -AccessToken $graphToken
    $AllSites += $resp.value
    $nextLink = $resp.'@odata.nextLink'
}
$AllSites = $AllSites | Where-Object { $_.webUrl -notmatch "/personal/" }

Write-Host "Gathering Company Info..." -ForegroundColor Yellow
$company = (Invoke-ApiCall -Method "GET" -Uri "$graphBase/organization" -AccessToken $graphToken).value[0]
$DirSyncStatus = "Not Available"
$LastDirSync = "Not Available"
$PasswordSync = "Not Available"
$LastPasswordSync = "Not Available"
if ($null -ne $company.DirSyncEnabled) {
    $DirSyncStatus = if ($company.DirSyncEnabled -eq $true) { "Enabled" } else { "Disabled" }
}
if ($DirSyncStatus -eq "Enabled") {
    $LastDirSync = if ($company.CompanyLastDirSyncTime) { $company.CompanyLastDirSyncTime } else { "Not Available" }
    try {
        $syncStatus = Invoke-ApiCall -Method "GET" -Uri "$graphBase/onPremisesDirectorySynchronization" -AccessToken $graphToken
        $sync = $syncStatus.value[0]
        $PasswordSync = if ($sync.value.passwordSyncConfiguration) { "Enabled" } else { "Disabled" }
        $LastPasswordSync = if ($sync.value.lastSuccessfulPasswordSyncDateTime) { $sync.value.lastSuccessfulPasswordSyncDateTime } else { "Not Available" }
    }
    catch { }
}
if (-not ($CompanyInfoTable -is [System.Collections.Generic.List[Object]])) {
    $CompanyInfoTable = New-Object 'System.Collections.Generic.List[Object]'
}
$CompanyInfoTable.Add([PSCustomObject]@{
        'Name'                = $company.displayName
        'Technical E-mail'    = ($company.TechnicalNotificationMails -join ", ")
        'Directory Sync'      = $DirSyncStatus
        'Password Sync'       = $PasswordSync
        'Last Password Sync'  = $LastPasswordSync
        'Last Directory Sync' = $LastDirSync
    })

Write-Host "Retrieving Admin Roles..." -ForegroundColor Yellow
$allRoles = Invoke-ApiCall -Method "GET" -Uri "$graphBase/directoryRoles" -AccessToken $graphToken
$adminRoles = $allRoles.value | Where-Object { $_.displayName -like "*Administrator*" }
$adminRoleTables = @{}
foreach ($role in $adminRoles) {
    $roleName = $role.displayName
    $result = Get-RoleMembers -roleDisplayName $roleName -Uri $graphBase -AccessToken $graphToken
    foreach ($key in $result.Keys) {
        if ($result[$key].Count -gt 0) {
            $adminRoleTables[$key] = $result[$key]
            Set-Variable -Name $key -Value $result[$key] -Scope Global
        }
    }
}

Write-Host "Getting Users with Strong Password Disabled..." -ForegroundColor Yellow
$LooseUsers = $AllUsers | Where-Object { $_.passwordPolicies -eq "DisableStrongPassword" }
foreach ($LooseUser in $LooseUsers) {
    $StrongPasswordTable.Add([PSCustomObject]@{
            'Name'                     = $LooseUser.displayName
            'UserPrincipalName'        = $LooseUser.userPrincipalName
            'Is Licensed'              = ($LooseUser.assignedLicenses.Count -gt 0)
            'Strong Password Required' = 'False'
        })
}
if ($StrongPasswordTable.Count -eq 0) {
    $StrongPasswordTable.Add([PSCustomObject]@{
            'Information' = 'Information: No Users were found with Strong Password Enforcement disabled'
        })
}


Write-Host "Getting Tenant Domains..." -ForegroundColor Yellow
$Domains = Invoke-ApiCall -Method "GET" -Uri "$graphBase/domains" -AccessToken $graphToken
foreach ($Domain in $Domains.value) {
    $DomainTable.Add([PSCustomObject]@{
            'Domain Name'         = $Domain.id
            'Verification Status' = $Domain.isVerified
            'Default'             = $Domain.isDefault
        })
}

Write-Host "Getting Groups..." -ForegroundColor Yellow
if (-not $table) { $table = New-Object System.Collections.Generic.List[Object] }
if (-not $GroupTypetable) { $GroupTypetable = New-Object System.Collections.Generic.List[Object] }
$GroupTypeCounts = @{
    "Office 365 Group"            = 0
    "Distribution List"           = 0
    "Security Group"              = 0
    "Mail Enabled Security Group" = 0
}
foreach ($Group in $Groups) {
    $type = switch ($true) {
        ($Group.groupTypes -contains "Unified") { "Office 365 Group"; break }
        ($Group.mailEnabled -and -not $Group.securityEnabled) { "Distribution List"; break }
        ($Group.mailEnabled -and $Group.securityEnabled) { "Mail Enabled Security Group"; break }
        (-not $Group.mailEnabled -and $Group.securityEnabled) { "Security Group"; break }
        default { "Other" }
    }
    if ($GroupTypeCounts.ContainsKey($type)) {
        $GroupTypeCounts[$type]++
    }
    $memberUri = "$graphBase/groups/$($Group.id)/members?$top=999"
    $memberData = Invoke-ApiCall -Method "GET" -Uri $memberUri -AccessToken $graphToken
    $userNames = if ($memberData.value) { ($memberData.value | Select-Object -ExpandProperty displayName) -join ", " } else { "" }
    $table.Add([PSCustomObject]@{
            'Name'           = $Group.displayName
            'Type'           = $type
            'Members'        = $userNames
            'E-mail Address' = $Group.mail
        })
}
foreach ($entry in $GroupTypeCounts.GetEnumerator()) {
    $GroupTypetable.Add([PSCustomObject]@{
            'Name'  = $entry.Key
            'Count' = $entry.Value
        })
}
if ($table.Count -eq 0) {
    $table.Add([PSCustomObject]@{
            'Information' = 'Information: No Groups were found in the tenant'
        })
}

Write-Host "Getting Licenses..." -ForegroundColor Yellow
$LicenseTable = New-Object System.Collections.Generic.List[Object]
$IsLicensedUsersTable = New-Object System.Collections.Generic.List[Object]
$Licenses = (Invoke-ApiCall -Method "GET" -Uri "$graphBase/subscribedSkus" -AccessToken $graphToken).value
foreach ($License in $Licenses) {
    $skuId = $License.skuId
    $skuPart = $License.skuPartNumber
    $friendlyName = if ($Sku.ContainsKey($skuPart)) { $Sku[$skuPart] } else { $skuPart }
    $Sku[$skuId] = $friendlyName
    $total = $License.prepaidUnits.enabled
    $assigned = $License.consumedUnits
    $unassigned = $total - $assigned
    if ($total -lt $LicenseFilter) {
        $LicenseTable.Add([PSCustomObject]@{
                'Name'                = $friendlyName
                'Total Amount'        = $total
                'Assigned Licenses'   = $assigned
                'Unassigned Licenses' = $unassigned
            })
    }
}
if ($LicenseTable.Count -eq 0) {
    $LicenseTable.Add([PSCustomObject]@{ 'Information' = 'Information: No Licenses were found in the tenant' })
}
$licensedUsers = ($AllUsers | Where-Object { $_.assignedLicenses.Count -gt 0 }).Count
$unlicensedUsers = ($AllUsers | Where-Object { $_.assignedLicenses.Count -eq 0 }).Count
$IsLicensedUsersTable.Add([PSCustomObject]@{ 'Name' = 'Users Licensed'; 'Count' = $licensedUsers })
$IsLicensedUsersTable.Add([PSCustomObject]@{ 'Name' = 'Users Not Licensed'; 'Count' = $unlicensedUsers })

Write-Host "Getting Users..." -ForegroundColor Yellow
if (-not $UserTable) { $UserTable = New-Object System.Collections.Generic.List[Object] }
foreach ($User in $AllUsers | Where-Object { $_.userType -ne "Guest" }) {
    $DisplayName = $User.displayName
    $UPN = $User.userPrincipalName
    $CreatedOn = $User.createdDateTime
    $Enabled = if ($User.accountEnabled) { "Enabled" } else { "" }
    $LastLogon = if ($User.signInActivity) { $User.signInActivity.lastSignInDateTime } else { "" }
    $TotalSize = ""
    $MailboxPercent = ""
    $ArchiveSize = ""
    $ArchivePercent = ""
    $ArchiveQuota = ""
    $QuotaBytes = ""
    $TotalBytes = ""
    $ArchiveBytes = ""
    $ArchiveQuotaBytes = ""
    try {
        $MailboxStats = Get-MailboxStatistics -Identity $UPN -ErrorAction Stop | Select-Object *
        if ($MailboxStats.TotalItemSize -match '\(([\d,]+) bytes\)') {
            $TotalBytes = [double]::Parse($matches[1].Replace(',', ''))
            $TotalSize = [math]::Round($TotalBytes / 1GB, 2)
        }
        if ($MailboxStats.SystemMessageSizeShutoffQuota -match '\(([\d,]+) bytes\)') {
            $QuotaBytes = [double]::Parse($matches[1].Replace(',', ''))
            if ($QuotaBytes -gt 0 -and $TotalBytes -gt 0) {
                $MailboxPercent = [math]::Round(($TotalBytes / $QuotaBytes) * 100, 2)
            }
        }
    }
    catch {}
    try {
        $ArchiveStats = Get-MailboxStatistics -Identity $UPN -Archive -ErrorAction Stop | Select-Object *
        if ($ArchiveStats.TotalItemSize -match '\(([\d,]+) bytes\)') {
            $ArchiveBytes = [double]::Parse($matches[1].Replace(',', ''))
            $ArchiveSize = [math]::Round($ArchiveBytes / 1GB, 2)
        }
        if ($ArchiveStats.BackupMessageSizeShutoffQuota -match '\(([\d,]+) bytes\)') {
            $ArchiveQuotaBytes = [double]::Parse($matches[1].Replace(',', ''))
            $ArchiveQuota = [math]::Round($ArchiveQuotaBytes / 1GB, 2)
            if ($ArchiveQuotaBytes -gt 0 -and $ArchiveBytes -gt 0) {
                $ArchivePercent = [math]::Round(($ArchiveBytes / $ArchiveQuotaBytes) * 100, 2)
            }
        }
    }
    catch {}
    $AssignedSkus = @()
    foreach ($SkuId in $User.assignedLicenses.skuId) {
        $AssignedSkus += if ($Sku.ContainsKey($SkuId)) { $Sku[$SkuId] } else { $SkuId }
    }
    $LicenseString = $AssignedSkus -join ", "
    $ProxyAddresses = if ($User.proxyAddresses) {
    ($User.proxyAddresses | Where-Object { $_ -notmatch '^X500:' -and $_ -notmatch '^/' } |
        ForEach-Object { ($_ -split ":")[-1] }) -join ", "
    }
    else { "" }
    $UserTable.Add([PSCustomObject]@{
            'Display Name'            = $DisplayName
            'User Principal Name'     = $UPN
            'Email Addresses'         = $ProxyAddresses
            'Created On'              = $CreatedOn
            'Account Status'          = $Enabled
            'Last Mailbox Login'      = $LastLogon
            'Mailbox Free Space (GB)' = $TotalSize
            'Mailbox Usage %'         = if ($MailboxPercent) { "$MailboxPercent%" } else { "" }
            'Archive Size (GB)'       = $ArchiveSize
            'Archive Size %'          = if ($ArchivePercent) { "$ArchivePercent%" } else { "" }
            'License'                 = $LicenseString
        })
}

Write-Host "Getting Guest Users..." -ForegroundColor Yellow
if (-not $GuestTable) { $GuestTable = New-Object System.Collections.Generic.List[Object] }
if (-not $GuestSPTable) { $GuestSPTable = New-Object System.Collections.Generic.List[Object] }
$SitePermissionsMap = @{}
foreach ($site in $AllSites) {
    try {
        $permissions = Invoke-ApiCall -Method "GET" -Uri "$($site.webUrl)/_api/web/roleassignments?`$expand=Member,RoleDefinitionBindings" -AccessToken $spToken
        $SitePermissionsMap[$site.webUrl] = $permissions.d.results
    }
    catch {
        if ($_.Exception.Response.StatusCode.value__ -eq 423) {
            Write-Host "Skipping locked site: $($site.webUrl)" -ForegroundColor DarkYellow
        }
    }
}
foreach ($Guest in $AllUsers | Where-Object { $_.userType -eq "Guest" }) {
    $CreatedOn = $Guest.createdDateTime
    $Age = (New-TimeSpan -Start $CreatedOn -End (Get-Date)).Days
    $Email = if ($Guest.proxyAddresses) {
        ($Guest.proxyAddresses | Where-Object { $_ -notmatch '^X500:' -and $_ -notmatch '^/' } | ForEach-Object { ($_ -split ":")[-1] }) -join ", "
    } else { "" }
    $MembershipGroups = @()
    try {
        $userGroups = Invoke-ApiCall -Method "GET" -Uri "$graphBase/users/$($Guest.id)/memberOf" -AccessToken $graphToken
        $MembershipGroups = $userGroups.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' } | Select-Object -ExpandProperty displayName
    }
    catch { }
    $MembershipSites = @()
    foreach ($kvp in $SitePermissionsMap.GetEnumerator()) {
        $siteUrl = $kvp.Key
        $perms = $kvp.Value
        foreach ($perm in $perms) {
            if ($perm.Member.PrincipalType -eq 1 -and $perm.Member.Title -eq $Guest.displayName) {
                $siteInfo = $AllSites | Where-Object { $_.webUrl -eq $siteUrl }
                if ($siteInfo) {
                    $MembershipSites += [PSCustomObject]@{
                        siteName = $siteInfo.displayName
                        siteUrl  = $siteInfo.webUrl
                    }
                }
            }
        }
    }
    $groupList = if ($MembershipGroups) { $MembershipGroups -join ", " } else { "" }
    $siteList = if ($MembershipSites) { $MembershipSites | ForEach-Object { $_.siteName } -join ", " } else { "" }
    $GuestTable.Add([PSCustomObject]@{
        'Name'                = $Guest.displayName
        'User Principal Name' = $Guest.userPrincipalName
        'E-mail Addresses'    = $Email
        'Created On'          = $CreatedOn
        'Account Age(days)'   = $Age
        'Creation Type'       = $Guest.creationType
        'Invitation Accepted' = $Guest.externalUserState
        'Group Membership'    = $groupList
        'Site Membership'     = $siteList
    })
    foreach ($site in $MembershipSites) {
        $GuestSPTable.Add([PSCustomObject]@{
            'Name'                = $Guest.displayName
            'User Principal Name' = $Guest.userPrincipalName
            'E-mail Addresses'    = $Email
            'Created On'          = $CreatedOn
            'Site name'           = $site.siteName
            'Site URL'            = $site.siteUrl
        })
    }
}

Write-Host "Getting Shared Mailboxes..." -ForegroundColor Yellow
$SharedMailboxes = Get-Recipient -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }
foreach ($SharedMailbox in $SharedMailboxes) {
    $Name = $SharedMailbox.Name
    $PrimEmail = $SharedMailbox.PrimarySmtpAddress.ToString()
    $ProxyList = @()
    foreach ($address in $SharedMailbox.EmailAddresses) {
        $cleanAddress = ($address -split ":")[-1]
        if ($cleanAddress -ne $PrimEmail) {
            $ProxyList += $cleanAddress
        }
    }
    $SharedMailboxTable.Add([PSCustomObject]@{
            'Name'             = $Name
            'Primary E-Mail'   = $PrimEmail
            'E-mail Addresses' = ($ProxyList -join ", ").TrimEnd(", ")
        })
}
if ($SharedMailboxTable.Count -eq 0) {
    $SharedMailboxTable.Add([PSCustomObject]@{
            'Information' = 'Information: No Shared Mailboxes were found in the tenant'
        })
}

Write-Host "Getting Contacts..." -ForegroundColor Yellow
$Contacts = Get-MailContact
foreach ($Contact in $Contacts) {
    $ContactName = $Contact.DisplayName
    $ContactPrimEmail = $Contact.PrimarySmtpAddress
    $ContactTable.Add([PSCustomObject]@{
            'Name'           = $ContactName
            'E-mail Address' = $ContactPrimEmail
        })
}
if ($ContactTable.Count -eq 0) {
    $ContactTable.Add([PSCustomObject]@{
            'Information' = 'Information: No Contacts were found in the tenant'
        })
}

Write-Host "Getting Mail Users..." -ForegroundColor Yellow
$MailUsers = Get-MailUser
foreach ($MailUser in $MailUsers) {
    $MailPrimEmail = $MailUser.PrimarySmtpAddress
    $MailName = $MailUser.DisplayName
    $MailArray = @()
    foreach ($address in $MailUser.EmailAddresses) {
        $split = ($address -split ":")[-1]
        if ($split -ne $MailPrimEmail) {
            $MailArray += $split
        }
    }
    $ContactMailUserTable.Add([PSCustomObject]@{
            'Name'             = $MailName
            'Primary E-Mail'   = $MailPrimEmail
            'E-mail Addresses' = ($MailArray -join ", ").TrimEnd(", ")
        })
}
if ($ContactMailUserTable.Count -eq 0) {
    $ContactMailUserTable.Add([PSCustomObject]@{
            'Information' = 'Information: No Mail Users were found in the tenant'
        })
}

Write-Host "Getting Room Mailboxes..." -ForegroundColor Yellow
$Rooms = Get-Mailbox -ResultSize Unlimited -Filter '(RecipientTypeDetails -eq "RoomMailbox")'
foreach ($Room in $Rooms) {
    $RoomName = $Room.DisplayName
    $RoomPrimEmail = $Room.PrimarySmtpAddress.ToString()
    $RoomArray = @()
    foreach ($address in $Room.EmailAddresses) {
        $clean = ($address -split ":")[-1]
        if ($clean -ne $RoomPrimEmail) {
            $RoomArray += $clean
        }
    }
    $RoomTable.Add([PSCustomObject]@{
            'Name'             = $RoomName
            'Primary E-Mail'   = $RoomPrimEmail
            'E-mail Addresses' = ($RoomArray -join ", ").TrimEnd(", ")
        })
}
if ($RoomTable.Count -eq 0) {
    $RoomTable.Add([PSCustomObject]@{
            'Information' = 'Information: No Room Mailboxes were found in the tenant'
        })
}

Write-Host "Getting Equipment Mailboxes..." -ForegroundColor Yellow
$EquipMailboxes = Get-Mailbox -ResultSize Unlimited -Filter '(RecipientTypeDetails -eq "EquipmentMailbox")'
foreach ($EquipMailbox in $EquipMailboxes) {
    $EquipName = $EquipMailbox.DisplayName
    $EquipPrimEmail = $EquipMailbox.PrimarySmtpAddress.ToString()
    $EquipArray = @()
    foreach ($address in $EquipMailbox.EmailAddresses) {
        $clean = ($address -split ":")[-1]
        if ($clean -ne $EquipPrimEmail) {
            $EquipArray += $clean
        }
    }
    $EquipTable.Add([PSCustomObject]@{
            'Name'             = $EquipName
            'Primary E-Mail'   = $EquipPrimEmail
            'E-mail Addresses' = ($EquipArray -join ", ").TrimEnd(", ")
        })
}
if ($EquipTable.Count -eq 0) {
    $EquipTable.Add([PSCustomObject]@{
            'Information' = 'Information: No Equipment Mailboxes were found in the tenant'
        })
}

Write-Host "Getting Enterprise Applications..." -ForegroundColor Yellow
if (-not $EnterpriseAppsTable) { $EnterpriseAppsTable = New-Object System.Collections.Generic.List[Object] }
$allApps = Invoke-ApiCall -Method GET -Uri "$graphBase/applications" -AccessToken $graphToken
foreach ($app in $allApps.value) {
    $name = $app.displayName
    $homepage = $app.homepage
    $created = if ($app.createdDateTime) { (Get-Date $app.createdDateTime).ToString("yyyy-MM-dd") } else { "N/A" }
    $certs = $app.keyCredentials | Where-Object { $_.type -eq "AsymmetricX509Cert" }
    $activeCert = $certs | Sort-Object -Property endDateTime -Descending | Select-Object -First 1
    if ($activeCert) {
        $expiry = Get-Date $activeCert.endDateTime
        $expiryStatus = if ($expiry -lt (Get-Date)) { "Expired" } else { "Valid" }
        $expiryDate = $expiry.ToString("yyyy-MM-dd")
    }
    else {
        $expiryStatus = "None"
        $expiryDate = ""
    }
    $EnterpriseAppsTable.Add([PSCustomObject]@{
            'Name'               = $name
            'Home Page'          = $homepage
            'Created'            = $created
            'Cert Expiry Status' = $expiryStatus
            'Active Cert Expiry' = $expiryDate
        })
}
if ($EnterpriseAppsTable.Count -eq 0) {
    $EnterpriseAppsTable.Add([PSCustomObject]@{ 'Information' = 'Information: No Enterprise Applications found in tenant' })
}

Write-Host "Generating HTML Report..." -ForegroundColor Yellow

$tabarray = @('Dashboard', 'Admins', 'Licenses', 'Users', 'Groups', 'Shared Mailboxes', 'Guests', 'Enterprise Applications', 'MFA Audit', 'Security Audit', 'Contacts', 'Resources')

Write-Host "Generating pie chart: Total Licenses" -ForegroundColor Cyan
$PieObject2 = Get-HTMLPieChartObject
$PieObject2.Title = "Office 365 Total Licenses"
$PieObject2.Size.Height = 500
$PieObject2.Size.width = 500
$PieObject2.ChartStyle.ChartType = 'doughnut'
$PieObject2.ChartStyle.ColorSchemeName = 'Random'
$PieObject2.DataDefinition.DataNameColumnName = 'Name'
$PieObject2.DataDefinition.DataValueColumnName = 'Total Amount'

Write-Host "Generating pie chart: Assigned Licenses" -ForegroundColor Cyan
$PieObject3 = Get-HTMLPieChartObject
$PieObject3.Title = "Office 365 Assigned Licenses"
$PieObject3.Size.Height = 500
$PieObject3.Size.width = 500
$PieObject3.ChartStyle.ChartType = 'doughnut'
$PieObject3.ChartStyle.ColorSchemeName = 'Random'
$PieObject3.DataDefinition.DataNameColumnName = 'Name'
$PieObject3.DataDefinition.DataValueColumnName = 'Assigned Licenses'

Write-Host "Generating pie chart: Unassigned Licenses" -ForegroundColor Cyan
$PieObject4 = Get-HTMLPieChartObject
$PieObject4.Title = "Office 365 Unassigned Licenses"
$PieObject4.Size.Height = 500
$PieObject4.Size.width = 500
$PieObject4.ChartStyle.ChartType = 'doughnut'
$PieObject4.ChartStyle.ColorSchemeName = 'Random'
$PieObject4.DataDefinition.DataNameColumnName = 'Name'
$PieObject4.DataDefinition.DataValueColumnName = 'Unassigned Licenses'

Write-Host "Generating pie chart: Group Type Breakdown" -ForegroundColor Cyan
$PieObjectGroupType = Get-HTMLPieChartObject
$PieObjectGroupType.Title = "Office 365 Groups"
$PieObjectGroupType.Size.Height = 500
$PieObjectGroupType.Size.width = 500
$PieObjectGroupType.ChartStyle.ChartType = 'doughnut'
$PieObjectGroupType.ChartStyle.ColorSchemeName = 'Random'
$PieObjectGroupType.DataDefinition.DataNameColumnName = 'Name'
$PieObjectGroupType.DataDefinition.DataValueColumnName = 'Count'

Write-Host "Generating pie chart: License Status Breakdown" -ForegroundColor Cyan
$PieObjectULicense = Get-HTMLPieChartObject
$PieObjectULicense.Title = "License Status"
$PieObjectULicense.Size.Height = 500
$PieObjectULicense.Size.width = 500
$PieObjectULicense.ChartStyle.ChartType = 'doughnut'
$PieObjectULicense.ChartStyle.ColorSchemeName = 'Random'
$PieObjectULicense.DataDefinition.DataNameColumnName = 'Name'
$PieObjectULicense.DataDefinition.DataValueColumnName = 'Count'

$rpt = New-Object 'System.Collections.Generic.List[System.Object]'
$rpt += get-htmlopenpage -TitleText 'Office 365 Tenant Report' -LeftLogoString $CompanyLogo 

Write-Host "Creating page: Dashboard" -ForegroundColor Cyan
$rpt += Get-HTMLTabHeader -TabNames $tabarray 
$rpt += get-htmltabcontentopen -TabName $tabarray[0] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HtmlContentOpen -HeaderText "Office 365 Dashboard"
$rpt += Get-HTMLContentOpen -HeaderText "Company Information"
$rpt += Get-HtmlContentTable $CompanyInfoTable 
$rpt += Get-HTMLContentClose

$rpt += get-HtmlColumn1of2
$rpt += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText 'Global Administrators'
$rpt += get-htmlcontentdatatable $GlobalAdministratorTable -HideFooter
$rpt += Get-HtmlContentClose
$rpt += get-htmlColumnClose
$rpt += get-htmlColumn2of2
$rpt += Get-HtmlContentOpen -HeaderText 'Users With Strong Password Enforcement Disabled'
$rpt += get-htmlcontentdatatable $StrongPasswordTable -HideFooter 
$rpt += Get-HTMLContentClose
$rpt += get-htmlColumnClose

$rpt += Get-HTMLContentOpen -HeaderText "Domains"
$rpt += Get-HtmlContentTable $DomainTable 
$rpt += Get-HTMLContentClose

$rpt += Get-HtmlContentClose 
$rpt += get-htmltabcontentclose

Write-Host "Creating page: Admins" -ForegroundColor Cyan
$rpt += get-htmltabcontentopen -TabName $tabarray[1] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HtmlContentOpen -HeaderText "Role Assignments"
$RoleTablesQueue = @()

foreach ($roleEntry in $adminRoleTables.GetEnumerator() | Where-Object { $_.Value.Count -gt 0 }) {
    if ($roleEntry.Key -eq "GlobalAdministratorTable") { continue }
    $roleLabel = ($roleEntry.Key -replace 'Table$', '')
    $roleLabel = $roleLabel -creplace '([a-z])([A-Z])', '$1 $2'
    $words = $roleLabel.Trim().Split(" ")
    $words[-1] = if ($words[-1] -notmatch "s$") { "$($words[-1])s" } else { $words[-1] }
    $formattedRole = ($words -join " ")
    $RoleTablesQueue += [PSCustomObject]@{
        RoleName = $formattedRole
        Data     = $roleEntry.Value
    }
}

for ($i = 0; $i -lt $RoleTablesQueue.Count; $i += 2) {
    $left = $RoleTablesQueue[$i]

    $rpt += Get-HtmlColumn1of2
    $rpt += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText $left.RoleName
    $rpt += Get-HTMLContentDataTable $left.Data -HideFooter
    $rpt += Get-HtmlContentClose
    $rpt += Get-HtmlColumnClose

    if ($i + 1 -lt $RoleTablesQueue.Count) {
        $right = $RoleTablesQueue[$i + 1]
        $rpt += Get-HtmlColumn2of2
        $rpt += Get-HtmlContentOpen -HeaderText $right.RoleName
        $rpt += Get-HTMLContentDataTable $right.Data -HideFooter
        $rpt += Get-HtmlContentClose
        $rpt += Get-HtmlColumnClose
    }
}
$rpt += Get-HtmlContentClose
$rpt += get-htmltabcontentclose

Write-Host "Creating page: Licenses" -ForegroundColor Cyan
$rpt += get-htmltabcontentopen -TabName $tabarray[2] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Licenses"
$rpt += get-htmlcontentdatatable $LicenseTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Licensing Charts"
$rpt += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2
$rpt += Get-HTMLPieChart -ChartObject $PieObject2 -DataSet $licensetable
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2
$rpt += Get-HTMLPieChart -ChartObject $PieObject3 -DataSet $licensetable
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLContentclose
$rpt += get-htmltabcontentclose

Write-Host "Creating page: Users" -ForegroundColor Cyan
$rpt += get-htmltabcontentopen -TabName $tabarray[3] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Users"
$rpt += get-htmlcontentdatatable $UserTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLContentOpen -HeaderText "Licensed & Unlicensed Users Chart"
$rpt += Get-HTMLPieChart -ChartObject $PieObjectULicense -DataSet $IsLicensedUsersTable
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose

Write-Host "Creating page: Groups" -ForegroundColor Cyan
$rpt += get-htmltabcontentopen -TabName $tabarray[4] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Groups"
$rpt += get-htmlcontentdatatable $Table -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Groups Chart"
$rpt += Get-HTMLPieChart -ChartObject $PieObjectGroupType -DataSet $GroupTypetable
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose

Write-Host "Creating page: Shared Mailboxes" -ForegroundColor Cyan
$rpt += get-htmltabcontentopen -TabName $tabarray[5] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Shared Mailboxes"
$rpt += get-htmlcontentdatatable $SharedMailboxTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose

Write-Host "Creating page: Guests" -ForegroundColor Cyan
$rpt += get-htmltabcontentopen -TabName $tabarray[6] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Guest Users"
$rpt += Get-HTMLContentTable $GuestTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 SharePoint Groups"
$rpt += Get-HTMLContentTable $GuestSharePointTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose

Write-Host "Creating page: Enterprise Applications" -ForegroundColor Cyan
$rpt += get-htmltabcontentopen -TabName $tabarray[7] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Enterprise Applications"
$rpt += get-htmlcontentdatatable $EnterpriseAppsTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose

Write-Host "Creating page: MFA Audit" -ForegroundColor Cyan
$rpt += get-htmltabcontentopen -TabName $tabarray[8] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "MFA Enforcement Report"
$rpt += get-htmlcontentdatatable $MfaAuditTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose

Write-Host "Creating page: Security Audit" -ForegroundColor Cyan
$rpt += get-htmltabcontentopen -TabName $tabarray[9] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Security Defaults & Policies"
$rpt += get-htmlcontentdatatable $SecurityAuditTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose

Write-Host "Creating page: Contacts" -ForegroundColor Cyan
$rpt += get-htmltabcontentopen -TabName $tabarray[10] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Contacts"
$rpt += get-htmlcontentdatatable $ContactTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Mail Users"
$rpt += get-htmlcontentdatatable $ContactMailUserTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose

Write-Host "Creating page: Resources" -ForegroundColor Cyan
$rpt += get-htmltabcontentopen -TabName $tabarray[11] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Room Mailboxes"
$rpt += get-htmlcontentdatatable $RoomTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Equipment Mailboxes"
$rpt += get-htmlcontentdatatable $EquipTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose

$rpt += Get-HTMLClosePage

$Day = (Get-Date).Day
$Month = (Get-Date).Month
$Year = (Get-Date).Year
$ReportName = ("$Month-$Day-$Year-O365 Tenant Report")
Save-HTMLReport -ReportContent $rpt -ShowReport -ReportName $ReportName -ReportPath $ReportSavePath