param (
    [Parameter(Mandatory = $true)]
    [object]$WebhookData
)

function Get-WebhookData {
    param (
        [Parameter(Mandatory = $true)]
        [object]$WebhookData
    )

    if (-not $WebhookData) {
        throw "Invalid Webhook Data."
    }
    
    return @{
        RequestHeader = $WebhookData.RequestHeader
        WebhookName   = $WebhookData.WebhookName
        RequestBody   = (ConvertFrom-Json $WebhookData.RequestBody)
    }
}

function Invoke-Webhook {
    param (
        [Parameter(Mandatory = $true)]
        [string]$RequestType,
        [Parameter(Mandatory = $true)]
        [string]$JsonPayload,
        [bool]$Test = $False
    )

    Write-Output "[INFO]: Setting webhook URL"
    $webhookUrl = "https://1ce978e3-0165-4625-bc15-bdcdfaeb9c7e.webhook.ae.azure-automation.net/webhooks?token=xUFmMKhtDzQVjGtnqDE6Xd8zMkBi2Zabt9BnzGkCFro%3d"

    Write-Output "[INFO]: Sending webhook request to $webhookUrl"

    if (-not $Test) {
        try {
            $response = Invoke-RestMethod -Method Post -Uri $webhookUrl -Body $JsonPayload -ContentType 'application/json'
            Write-Output "[INFO]: Webhook response: $($response | ConvertTo-Json -Compress)"
        }
        catch {
            Write-Output "[ERROR]: Failed to send webhook request $_"
        }
    }
    else {
        Write-Output "[TEST]: Test mode enabled, webhook not sent"
    }
}

function Convert-OnboardingFields($sections, $log, $ticket, $formName) {
    $result = @{
        # Identifiers
        "cwpsa_ticket"                             = $ticket
        [string]"action"                           = "post"
        [string]"request_type"                     = ""
        [string]"status_operation"                 = ""
        [string]"status_preapproved"               = "No"
        [string]"status_approved"                  = "No"
        [string]"status_sspr"                      = "No"
        [string]"status_user_created"              = "No"
        [string]"status_license_assigned"          = "No"
        [string]"status_notification_scheduled"    = "No"

        # User Details
        [string]"user_firstname"                   = ""
        [string]"user_lastname"                    = ""
        [string]"user_middlename"                  = ""
        [string]"user_fullname"                    = ""
        [string]"user_username"                    = ""
        [string]"user_upn"                         = ""
        [string]"user_mailnickname"                = ""
        [string]"user_externalemailaddress"        = ""
        [string]"user_displayname"                 = ""
        [string]"user_primary_smtp"                = ""
        [string]"user_mobile_personal"             = ""
        [string]"user_mobile_business"             = ""
        [string]"user_telephone"                   = ""
        [string]"user_homephone"                   = ""
        [string]"user_fax"                         = ""
        [string]"user_employee_id"                 = ""
        [string]"user_employee_type"               = ""
        [string]"user_password_reset_required"     = "Yes"
        [string]"user_startdate"                   = ""
        [string]"user_enddate"                     = ""
        [string]"user_account_status"              = ""
        [string]"user_account_expiry"              = ""
        [string]"user_company"                     = ""
        [string]"user_department"                  = ""
        [string]"user_title"                       = ""
        [string]"user_manager_name"                = ""
        [string]"user_manager_upn"                 = ""
        [string]"user_office"                      = ""
        [string]"user_streetaddress"               = ""
        [string]"user_city"                        = ""
        [string]"user_state"                       = ""
        [string]"user_postalcode"                  = ""
        [string]"user_country"                     = ""
        [string]"user_description"                 = ""
        [string]"user_extensionattribute1"         = ""
        [string]"user_extensionattribute2"         = ""
        [string]"user_extensionattribute3"         = ""
        [string]"user_extensionattribute4"         = ""
        [string]"user_extensionattribute5"         = ""
        [string]"user_extensionattribute6"         = ""
        [string]"user_extensionattribute7"         = ""
        [string]"user_extensionattribute8"         = ""
        [string]"user_extensionattribute9"         = ""
        [string]"user_extensionattribute10"        = ""
        [string]"user_extensionattribute11"        = ""
        [string]"user_extensionattribute12"        = ""
        [string]"user_extensionattribute13"        = ""
        [string]"user_extensionattribute14"        = ""
        [string]"user_extensionattribute15"        = ""
        
        # Organisation Metadata
        [string]"organisation_company_append"      = "No"
        [string]"organisation_persona"             = ""
        [string]"organisation_site"                = ""

        # Clone
        [string]"clone_tag"                        = "No"
        [string]"clone_user"                       = ""

        # Microsoft & Exchange
        [string]"microsoft_domain"                 = ""
        [string]"exchange_email_domain"            = ""
        [string]"exchange_add_aliases"             = "No"
        [string]"exchange_sharedmailbox"           = ""
        [string]"exchange_distributionlist"        = ""
        [string]"exchange_mailboxdelegate"         = ""
        [string]"exchange_ooodelegate"             = ""
        [string]"exchange_forwardto"               = ""

        # Group Memberships
        [string]"group_license_sku"                = ""
        [string]"group_license"                    = ""
        [string]"group_security"                   = ""
        [string]"group_teams"                      = ""
        [string]"group_software"                   = ""
        [string]"group_sharepoint"                 = ""
        [string]"group_extra1"                     = ""
        [string]"group_extra2"                     = ""

        # Hybrid AD
        [string]"hybrid_ou"                        = ""
        [string]"hybrid_home_drive"                = ""
        [string]"hybrid_home_driveletter"          = ""

        # Teams
        [string]"teams_delegateowner"              = ""

        # Device Requirements - Mobile
        [string]"mobile_required"                  = "No"
        [string]"mobile_source"                    = ""
        [string]"mobile_vendor"                    = ""
        [string]"mobile_tag"                       = ""

        # Device Requirements - Mobile Number
        [string]"mobile_number_required"           = "No"
        [string]"mobile_number_source"             = ""
        [string]"mobile_number_tag"                = ""

        # Device Requirements - Endpoint
        [string]"endpoint_required"                = "No"
        [string]"endpoint_source"                  = ""
        [string]"endpoint_vendor"                  = ""
        [string]"endpoint_tag"                     = ""

        # Device Requirements - Tablet
        [string]"tablet_required"                  = "No"
        [string]"tablet_source"                    = ""
        [string]"tablet_vendor"                    = ""
        [string]"tablet_tag"                       = ""

        # Asset Management
        [string]"asset_delegate"                   = ""

        # Manual Tasks
        [string]"manual_task"                      = ""

        # Notification
        [string]"notification_sendto_submitter"    = "No"
        [string]"notification_password_recipient"  = ""
        [string]"notification_password_recipient1" = ""
        [string]"notification_password_recipient2" = ""

        # Client Specific Task
        [string]"cs_task"                          = "No"
        [string]"cs_task_notification"             = "No"
    }

    switch -Regex ($formName) {
        "User Onboarding" { $result["request_type"] = "user_onboarding"; break }
        "User Offboarding" { $result["request_type"] = "user_offboarding"; break }
    }

    $managerHandled = $false
    $groupFields = @("group_license", "group_security", "group_teams", "group_software", "group_sharepoint", "group_extra1", "group_extra2")

    foreach ($section in $sections) {
        foreach ($field in $section.fields) {
            $identifier = $field.identifier

            if (-not $managerHandled -and $identifier -eq "user_manager" -and $field.value -match "^(?<name>.+?) <(?<email>[^>]+)>$") {
                $result["user_manager_name"] = $Matches["name"]
                $result["user_manager_upn"] = $Matches["email"]
                $managerHandled = $true
                continue
            }

            if ($identifier -match '^status_(preapproved|sspr)$') {
                $result[$identifier] = "Yes"
                continue
            }

            if ($formName -match 'Offboarding|SharePoint User Permissions') {
                if ((($identifier -eq "user_offboardeduser") -or ($identifier -eq "user_identifier")) -and $field.choices[0].name -match "^(?<name>.+?) <(?<email>[^>]+)>$") {
                    $result["user_fullname"] = $Matches["name"]
                    $result["user_upn"] = $Matches["email"]
                    $result["user_primary_smtp"] = $Matches["email"]
                    continue
                }
            }

            if ($field.PSObject.Properties.Name -contains 'choices' -and $identifier -ne "user_upn") {
                $out = @()
                foreach ($choice in $field.choices) {
                    if ($choice.name) {
                        $raw = $choice.name -replace '&lt;', '<' -replace '&gt;', '>'

                        if ($groupFields -contains $identifier) {
                            $isManual = $raw -match '^(?<base>.+?) \[MANUAL\]$'
                            if ($isManual) { $raw = $Matches["base"] }

                            if ($raw -match '^(?<pre>[^\[]*?)\s*\[(?<inner>[^\]]+)\]$') {
                                $pre = ($Matches["pre"]).Trim()
                                $inner = $Matches["inner"]
                                $innerSplit = $inner -split '\|', 2
                                $main = $innerSplit[0]
                                $email = if ($innerSplit.Count -gt 1) { $innerSplit[1] } else { "" }

                                $mainParts = $main -split ':', 2
                                $partA = $mainParts[0]
                                $partB = if ($mainParts.Count -gt 1) { $mainParts[1] } else { "" }

                                $permList = @("Member", "Owner", "Send on Behalf", "Send As", "Full Access", "Read Permission")
                                $perm = if ($permList -contains $partB) { $partB } else { "" }

                                $sourceTag = if ($isManual) { "MANUAL" } else { "" }
                                $out += "$pre`:$partA`:$perm`:$sourceTag|$email"
                            }
                        }
                        else {
                            if ($raw -match '<(?<email>[^>]+)>') {
                                $out += $Matches["email"]
                            }
                        }
                    }
                }
                $result[$identifier] = ($out -join ',')
                continue
            }

            if ($field.PSObject.Properties.Name -contains 'value') {
                $cleaned = $field.value
                if ($cleaned -match '\[(?<inner>[^\]]+)\]') { $cleaned = $Matches["inner"] }
                elseif ($cleaned -match '<(?<email>[^>]+)>') { $cleaned = $Matches["email"] }
                
                if ($result.ContainsKey($identifier) -and $result[$identifier] -eq "No" -and $cleaned) {
                    $result[$identifier] = $cleaned
                }
                elseif ($result.ContainsKey($identifier)) {
                    $result[$identifier] = $cleaned
                }
                continue
            }

            if ($identifier -in @("user_firstname", "user_lastname")) {
                $fn = $result["user_firstname"]
                $ln = $result["user_lastname"]
                if ($fn -and $ln) { $result["user_fullname"] = "$fn $ln" }
            }
            
            else {
                if ($result.ContainsKey($identifier) -and $result[$identifier] -eq "No" -and $field.value) {
                    $result[$identifier] = $field.value
                }
                elseif ($result.ContainsKey($identifier)) {
                    $result[$identifier] = $field.value
                }
            }
        }
    }

    return $result
}

function ProcessMainFunction {
    Write-Output "[INFO]: Processing Webhook"

    try {
        $WebhookData = Get-WebhookData -WebhookData $WebhookData
        Write-Output "[INFO]: Webhook data parsed successfully"
    }
    catch {
        Write-Error "[ERROR]: Failed to parse WebhookData. Error: $_"
        exit 1
    }

    Write-Output "[INFO]: The Webhook Name: $($WebhookData.WebhookName)"
    Write-Output "[INFO]: The Webhook Header: $($WebhookData.RequestHeader)"
    Write-Output "[INFO]: Webhook Processed. Starting the Main Process"

    try {
        $DecodedRequestBody = $WebhookData.RequestBody
        Write-Output "[INFO]: Webhook RequestBody successfully extracted"
    }
    catch {
        Write-Error "[ERROR]: Failed to parse WebhookData.RequestBody. Error: $_"
        exit 1
    }

    Write-Output "[INFO]: Processing the following RequestBody:"
    Write-Output ($DecodedRequestBody | ConvertTo-Json -Depth 10)

    foreach ($submission in $DecodedRequestBody) {
        $sections = $submission.form.sections
        $formName = $submission.form.name
        $ticketId = $submission.ticket.entityId

        $result = Convert-OnboardingFields -sections $sections -log $null -ticket $ticketId -formName $formName
        $JsonPayload = $result | ConvertTo-Json -Depth 10
        Write-Output "[INFO]: Converted payload:"
        Write-Output $JsonPayload

        Write-Output "[INFO]: Invoking webhook with converted payload"
        Invoke-Webhook -RequestType $result.request_type -JsonPayload $JsonPayload
    }

    Write-Output "[INFO]: ProcessMainFunction execution completed"
}


ProcessMainFunction
