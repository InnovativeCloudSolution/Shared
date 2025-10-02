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

    switch ($RequestType.ToLower()) {
        "user_onboarding" {
            $webhookUrl = "https://au.webhook.myconnectwise.net/HbSHrnJJbexELm2IDb0OW85moFmd1VwQzpVBVzIqtNKFGdjrlzxOqNThFH-kbyccLctvZA=="
        }
        "user_offboarding" {
            $webhookUrl = "https://au.webhook.myconnectwise.net/HLnbqXwYbb9EeW_bXr0OCsxkoAHNgF8QlcNNAmYq5YHVGo3oS1kqoSFvEEdKYiAv0rMiyQ=="
        }
        default {
            Write-Output "[ERROR]: Invalid RequestType [$RequestType] — no webhook URL mapped"
            return
        }
    }

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

function Get-CosmosDbAuthHeader {
    param (
        [string]$Verb,
        [string]$ResourceType,
        [string]$ResourceLink,
        [string]$MasterKey,
        [string]$DateUtc
    )
    $key = [Convert]::FromBase64String($MasterKey)
    $payload = "$($Verb.ToLowerInvariant())`n$($ResourceType.ToLowerInvariant())`n$ResourceLink`n$($DateUtc.ToLowerInvariant())`n`n"
    $hmac = [System.Security.Cryptography.HMACSHA256]::new()
    $hmac.Key = $key
    $sig = [Convert]::ToBase64String($hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($payload)))
    $enc = [System.Web.HttpUtility]::UrlEncode($sig)
    "type=master&ver=1.0&sig=$enc"
}

function Invoke-CosmosDbRequest {
    param (
        [string]$CosmosEndpoint,
        [string]$Verb,
        [string]$Uri,
        [string]$ResourceType,
        [string]$ResourceLink,
        [string]$MasterKey,
        [hashtable]$AdditionalHeaders = @{},
        [string]$Body = $null
    )
    $date = [DateTime]::UtcNow.ToString("R")
    $auth = Get-CosmosDbAuthHeader -Verb $Verb -ResourceType $ResourceType -ResourceLink $ResourceLink -MasterKey $MasterKey -DateUtc $date
    $headers = @{
        Authorization  = $auth
        "x-ms-version" = "2017-02-22"
        "x-ms-date"    = $date
    }
    foreach ($k in $AdditionalHeaders.Keys) { $headers[$k] = $AdditionalHeaders[$k] }
    $full = "$CosmosEndpoint/$Uri"
    if ($Body) { Invoke-RestMethod -Method $Verb -Uri $full -Headers $headers -Body $Body }
    else { Invoke-RestMethod -Method $Verb -Uri $full -Headers $headers }
}

function Get-CosmosDbDocument {
    param (
        [string]$CosmosEndpoint,
        [string]$DatabaseName,
        [string]$ContainerName,
        [string]$MasterKey,
        [int]$TicketNumber
    )
    $rl = "dbs/$DatabaseName/colls/$ContainerName"
    $uri = "$rl/docs"
    $qry = @{ query = "SELECT * FROM c WHERE c['cwpsa_ticket'] = @ticket"; parameters = @(@{name = '@ticket'; value = $TicketNumber }) } | ConvertTo-Json -Depth 3
    $hdr = @{ "x-ms-documentdb-isquery" = "true"; "x-ms-documentdb-query-enablecrosspartition" = "true"; "Content-Type" = "application/query+json" }
    (Invoke-CosmosDbRequest -CosmosEndpoint $CosmosEndpoint -Verb POST -Uri $uri -ResourceType docs -ResourceLink $rl -MasterKey $MasterKey -AdditionalHeaders $hdr -Body $qry).Documents
}

function New-CosmosDbDocument {
    param (
        [string]$CosmosEndpoint,
        [string]$DatabaseName,
        [string]$ContainerName,
        [string]$MasterKey,
        [hashtable]$Document
    )
    if (Get-CosmosDbDocument -CosmosEndpoint $CosmosEndpoint -DatabaseName $DatabaseName -ContainerName $ContainerName -MasterKey $MasterKey -TicketNumber $Document.cwpsa_ticket) {
        throw "Document already exists for ticket [$($Document.cwpsa_ticket)]"
    }
    Add-CosmosDbDocument -CosmosEndpoint $CosmosEndpoint -DatabaseName $DatabaseName -ContainerName $ContainerName -MasterKey $MasterKey -Document $Document
}

function Add-CosmosDbDocument {
    param (
        [string]$CosmosEndpoint,
        [string]$DatabaseName,
        [string]$ContainerName,
        [string]$MasterKey,
        [hashtable]$Document
    )
    $rl = "dbs/$DatabaseName/colls/$ContainerName"
    $uri = "$rl/docs"
    $partitionKey = $Document["cwpsa_ticket"]
    $hdr = @{
        "x-ms-documentdb-is-upsert"    = "true"
        "x-ms-documentdb-partitionkey" = "[$partitionKey]"
        "Content-Type"                 = "application/json"
    }

    $body = $Document | ConvertTo-Json -Depth 10
    (Invoke-CosmosDbRequest -CosmosEndpoint $CosmosEndpoint -Verb POST -Uri $uri -ResourceType docs -ResourceLink $rl -MasterKey $MasterKey -AdditionalHeaders $hdr -Body $body).id
}

function Update-CosmosDbDocument {
    param (
        [string]$CosmosEndpoint,
        [string]$DatabaseName,
        [string]$ContainerName,
        [string]$MasterKey,
        [int]$TicketNumber,
        [hashtable]$InputDocument
    )
    $docs = Get-CosmosDbDocument -CosmosEndpoint $CosmosEndpoint -DatabaseName $DatabaseName -ContainerName $ContainerName -MasterKey $MasterKey -TicketNumber $TicketNumber
    if (-not $docs) { throw "No document found for ticket [$TicketNumber]" }
    $doc = $docs[0]
    foreach ($k in $InputDocument.Keys) { if ($InputDocument[$k]) { $doc[$k] = $InputDocument[$k] } }
    $id = $doc.id
    $rl = "dbs/$DatabaseName/colls/$ContainerName/docs/$id"
    $hdr = @{ "Content-Type" = "application/json" }
    $body = $doc | ConvertTo-Json -Depth 10
    (Invoke-CosmosDbRequest -CosmosEndpoint $CosmosEndpoint -Verb PUT -Uri $rl -ResourceType docs -ResourceLink $rl -MasterKey $MasterKey -AdditionalHeaders $hdr -Body $body).id
}

function Get-PersonaTemplate {
    param (
        [string]$PersonaName,
        [string]$CompanyIdentifier,
        [string]$CosmosEndpoint,
        [string]$CosmosKey
    )

    Add-Type -AssemblyName System.Web

    $DatabaseName = $CompanyIdentifier.ToLower()
    $ContainerName = "persona"
    $ResourceLink = "dbs/$DatabaseName/colls/$ContainerName"
    $DateUtc = [DateTime]::UtcNow.ToString("R")
    $AuthHeader = Get-CosmosDbAuthHeader -Verb "POST" -ResourceType "docs" -ResourceLink $ResourceLink -MasterKey $CosmosKey -DateUtc $DateUtc

    $Query = @{
        query      = "SELECT * FROM c WHERE c['persona'] = @persona"
        parameters = @(@{ name = "@persona"; value = $PersonaName })
    } | ConvertTo-Json -Depth 3

    $Headers = @{
        Authorization                                = $AuthHeader
        "x-ms-version"                               = "2018-12-31"
        "x-ms-date"                                  = $DateUtc
        "x-ms-documentdb-isquery"                    = "true"
        "x-ms-documentdb-query-enablecrosspartition" = "true"
        "Content-Type"                               = "application/query+json"
    }

    $CosmosEndpoint = $CosmosEndpoint.TrimEnd("/")
    $Uri = "$CosmosEndpoint/$ResourceLink/docs"

    try {
        $Response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $Query
        if ($Response -and $Response.Documents.Count -gt 0) {
            return $Response.Documents[0]
        }
    }
    catch {}

    return @{}
}

function Get-UserManagementDefault {
    param (
        [string]$RequestType,
        [string]$CompanyIdentifier,
        [string]$CosmosEndpoint,
        [string]$CosmosKey
    )

    Add-Type -AssemblyName System.Web

    $DatabaseName = $CompanyIdentifier.ToLower()
    switch ($RequestType.ToLower()) {
        "user_onboarding" { $ContainerName = "user_onboarding_default" }
        "user_offboarding" { $ContainerName = "user_offboarding_default" }
        default { return @{} }
    }

    $ResourceLink = "dbs/$DatabaseName/colls/$ContainerName"
    $DateUtc = [DateTime]::UtcNow.ToString("R")
    $AuthHeader = Get-CosmosDbAuthHeader -Verb "POST" -ResourceType "docs" -ResourceLink $ResourceLink -MasterKey $CosmosKey -DateUtc $DateUtc

    $Query = @{ query = "SELECT * FROM c" } | ConvertTo-Json -Depth 3

    $Headers = @{
        Authorization                                = $AuthHeader
        "x-ms-version"                               = "2018-12-31"
        "x-ms-date"                                  = $DateUtc
        "x-ms-documentdb-isquery"                    = "true"
        "x-ms-documentdb-query-enablecrosspartition" = "true"
        "Content-Type"                               = "application/query+json"
    }

    $CosmosEndpoint = $CosmosEndpoint.TrimEnd("/")
    $Uri = "$CosmosEndpoint/$ResourceLink/docs"

    try {
        $Response = Invoke-RestMethod -Method POST -Uri $Uri -Headers $Headers -Body $Query
        if ($Response -and $Response.Documents.Count -gt 0) {
            return $Response.Documents[0]
        }
    }
    catch {}

    return @{}
}

function Format-UserString {
    param (
        [string]$FormatString,
        [string]$FirstName,
        [string]$LastName
    )
    $FormatString `
        -replace "firstname", $FirstName `
        -replace "lastname", $LastName `
        -replace "f", $FirstName.Substring(0, 1) `
        -replace "l", $LastName.Substring(0, 1)
}

function Get-Fields {
    param (
        $Source,
        [hashtable]$Target
    )

    if ($Source.license -is [System.Collections.IEnumerable]) {
        $Target["group_license"] = Set-Field $Target["group_license"] ([string]::Join(",", ($Source.license | ForEach-Object { $_.name })))
        $Target["group_license_sku"] = Set-Field $Target["group_license_sku"] ([string]::Join(",", ($Source.license | Where-Object { $_.sku } | ForEach-Object { $_.sku })))
    }

    if ($Source.group -is [System.Collections.IEnumerable]) {
        $Target["group_security"] = Set-Field $Target["group_security"] ([string]::Join(",", ($Source.group | ForEach-Object { $_.name })))
    }

    if ($Target.ContainsKey("group_teams") -and $Target["group_teams"]) {
        $Target["group_security"] = Set-Field $Target["group_security"] $Target["group_teams"]
    }

    if ($Source.application -is [System.Collections.IEnumerable]) {
        $Target["group_software"] = Set-Field $Target["group_software"] ([string]::Join(",", ($Source.application | ForEach-Object { $_.name })))
    }

    if ($Source.sharepoint -is [System.Collections.IEnumerable]) {
        $Target["group_sharepoint"] = Set-Field $Target["group_sharepoint"] ([string]::Join(",", ($Source.sharepoint | ForEach-Object { $_.name })))
    }

    if ($Source.mailbox -is [System.Collections.IEnumerable]) {
        $Target["exchange_sharedmailbox"] = Set-Field $Target["exchange_sharedmailbox"] ([string]::Join(",", ($Source.mailbox | ForEach-Object { $_.name })))
    }

    if ($Source.distribution -is [System.Collections.IEnumerable]) {
        $Target["exchange_distributionlist"] = Set-Field $Target["exchange_distributionlist"] ([string]::Join(",", ($Source.distribution | ForEach-Object { $_.name })))
    }

    if ($Source.ou) { $Target["hybrid_ou"] = Set-Field $Target["hybrid_ou"] $Source.ou }
    if ($Source.home_drive) { $Target["hybrid_home_drive"] = Set-Field $Target["hybrid_home_drive"] $Source.home_drive }
    if ($Source.home_driveletter) { $Target["hybrid_home_driveletter"] = Set-Field $Target["hybrid_home_driveletter"] $Source.home_driveletter }
    if ($Source.fax) { $Target["user_fax"] = Set-Field $Target["user_fax"] $Source.fax }
}

function Set-Field {
    param (
        $Existing,
        $NewValue
    )
    if (-not $NewValue) { return $Existing }
    if (-not $Existing) { return $NewValue }
    return "$Existing,$NewValue"
}

function Update-OnboardingInput {
    param (
        [hashtable]$InputData,
        [string]$CosmosEndpoint,
        [string]$CosmosKey,
        [string]$CompanyIdentifier
    )

    $defaultPersona = Get-PersonaTemplate -PersonaName "DEFAULT" -CompanyIdentifier $CompanyIdentifier -CosmosEndpoint $CosmosEndpoint -CosmosKey $CosmosKey
    if ($defaultPersona.Count -gt 0) { Get-Fields $defaultPersona $InputData }

    $personaValue = $InputData["organisation_persona"]
    if ($personaValue) {
        $personaData = Get-PersonaTemplate -PersonaName $personaValue -CompanyIdentifier $CompanyIdentifier -CosmosEndpoint $CosmosEndpoint -CosmosKey $CosmosKey
        if ($personaData.Count -gt 0) { Get-Fields $personaData $InputData }
    }

    $defaults = Get-UserManagementDefault -RequestType "user_onboarding" -CompanyIdentifier $CompanyIdentifier -CosmosEndpoint $CosmosEndpoint -CosmosKey $CosmosKey
    if ($defaults) {
        $fname = ($InputData["user_firstname"] | ForEach-Object { $_ }) -join ""
        $lname = ($InputData["user_lastname"]  | ForEach-Object { $_ }) -join ""

        $InputData["user_username"] = Format-UserString -FormatString $defaults.un_format -FirstName $fname -LastName $lname
        $UpnPrefix = Format-UserString -FormatString $defaults.upn_format -FirstName $fname -LastName $lname
        $Domain = $InputData["microsoft_domain"]
        $InputData["user_upn"] = if ($UpnPrefix -and $Domain) { "$UpnPrefix@$Domain" } else { $UpnPrefix }
        $InputData["user_mailnickname"] = Format-UserString -FormatString $defaults.mn_format -FirstName $fname -LastName $lname
        $ti = [System.Globalization.CultureInfo]::InvariantCulture.TextInfo
        $InputData["user_displayname"] = $ti.ToTitleCase((Format-UserString -FormatString $defaults.dn_format -FirstName $fname -LastName $lname).ToLower())

        $CombinedManual = @()
        if ($defaults.manual_task) { $CombinedManual += $defaults.manual_task }
        if ($InputData.manual_task) { $CombinedManual += $InputData.manual_task }
        $InputData["manual_task"] = ($CombinedManual | Where-Object { $_ }) -join ","

        if (($InputData["organisation_company_append"] -as [string]).Trim().ToLower() -eq "yes") {
            $CompanyName = ($InputData["user_company"] -as [string]).Trim()
            if ($CompanyName) {
                $InputData["user_displayname"] = "$($InputData["user_displayname"]) ($CompanyName)"
            }
        }

        $MailNickname = ($InputData["user_mailnickname"] | ForEach-Object { $_ }) -join "" | ForEach-Object { $_.Trim() }
        $EmailDomain = ($InputData["exchange_email_domain"] | ForEach-Object { $_ }) -join "" | ForEach-Object { $_.Trim() }
        $MicrosoftDomain = ($InputData["microsoft_domain"] | ForEach-Object { $_ }) -join "" | ForEach-Object { $_.Trim() }

        if ($EmailDomain -and ($EmailDomain -ne $MicrosoftDomain)) {
            $InputData["user_primary_smtp"] = "$MailNickname@$EmailDomain"
        }
        else {
            $InputData["user_primary_smtp"] = "$MailNickname@$MicrosoftDomain"
        }
    }

    $recipient1 = ($InputData["notification_password_recipient1"] -as [string]).Trim()
    $recipient2 = ($InputData["notification_password_recipient2"] -as [string]).Trim()
    
    if ($recipient1) {
        $InputData["notification_password_recipient"] = $recipient1
    }
    elseif ($recipient2) {
        $InputData["notification_password_recipient"] = $recipient2
    }
    else {
        $InputData["notification_password_recipient"] = ""
    }

    $InputData.Remove("notification_password_recipient1") | Out-Null
    $InputData.Remove("notification_password_recipient2") | Out-Null
    $InputData.Remove("group_license_sku") | Out-Null

    return $InputData
}

function Update-OffboardingInput {
    param (
        [hashtable]$InputData,
        [string]$CosmosEndpoint,
        [string]$CosmosKey,
        [string]$CompanyIdentifier
    )

    $defaults = Get-UserManagementDefault -RequestType "user_offboarding" -CompanyIdentifier $CompanyIdentifier -CosmosEndpoint $CosmosEndpoint -CosmosKey $CosmosKey
    if ($defaults) {
        $CombinedManual = @()
        if ($defaults.manual_task) { $CombinedManual += $defaults.manual_task }
        if ($InputData.manual_task) { $CombinedManual += $InputData.manual_task }
        $InputData["manual_task"] = ($CombinedManual | Where-Object { $_ }) -join ","

        if ($defaults.hybrid_ou) {
            $InputData["hybrid_ou"] = $defaults.hybrid_ou
        }
    }

    return $InputData
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

    $payload = @{ }
    $DecodedRequestBody.PSObject.Properties | ForEach-Object {
        $payload[$_.Name] = $_.Value
    }

    $RequestType = ($payload.request_type).ToLower()
    $Action = ($payload.action).ToLower()
    $TicketNumber = [int]$payload.cwpsa_ticket

    if (-not $Action) {
        Write-Output "[WARN]: Missing 'action' in payload. Skipping."
        return
    }

    Set-StrictMode -Off

    $CWMClientId = Get-AutomationVariable -Name 'clientId'
    $CWMPublicKey = Get-AutomationVariable -Name 'PublicKey'
    $CWMPrivateKey = Get-AutomationVariable -Name 'PrivateKey'
    $CWMCompanyId = Get-AutomationVariable -Name 'CWManageCompanyId'
    $CWMUrl = Get-AutomationVariable -Name 'CWManageUrl'

    $Connection = @{
        Server     = $CWMUrl
        Company    = $CWMCompanyId
        pubkey     = $CWMPublicKey
        privatekey = $CWMPrivateKey
        clientId   = $CWMClientId
    }

    Connect-CWM @Connection
    $Ticket = Get-CWMTicket -id $TicketNumber

    $payload.Remove("action")       | Out-Null
    $payload.Remove("request_type") | Out-Null

    $CosmosEndpoint = Get-AutomationVariable -Name "MIT-AZU1-CosmosDB-Endpoint"
    $CosmosKey = Get-AutomationVariable -Name "MIT-AZU1-CosmosDB-Key"
    $DatabaseName = "submission"
    $ContainerName = $RequestType
    $CompanyIdentifier = $Ticket.company.identifier

    if ($RequestType -notin @("user_onboarding", "user_offboarding")) {
        Write-Output "[ERROR]: Invalid request_type [$RequestType]"
        return
    }

    try {
        Write-Output "Action = $Action"
        Write-Output "Ticket = $TicketNumber"
        Write-Output "Request Type = $RequestType"

        switch ($Action.ToLower()) {
            "get" {
                $results = Get-CosmosDbDocument -CosmosEndpoint $CosmosEndpoint -DatabaseName $DatabaseName -ContainerName $ContainerName -MasterKey $CosmosKey -TicketNumber $TicketNumber
                if ($results) {
                    Write-Output "Found $($results.Count) document(s) for ticket [$TicketNumber]"
                    $results | ConvertTo-Json -Depth 10
                }
                else {
                    Write-Output "No document found for ticket [$TicketNumber]"
                }
            }

            "post" {
                $payload["id"] = [guid]::NewGuid().ToString()
                $payload["cwpsa_ticket"] = $TicketNumber

                if ($RequestType -eq "user_onboarding") {
                    $enriched = Update-OnboardingInput -InputData $payload -CosmosEndpoint $CosmosEndpoint -CosmosKey $CosmosKey -CompanyIdentifier $CompanyIdentifier
                }
                elseif ($RequestType -eq "user_offboarding") {
                    $enriched = Update-OffboardingInput -InputData $payload -CosmosEndpoint $CosmosEndpoint -CosmosKey $CosmosKey -CompanyIdentifier $CompanyIdentifier
                }
                else {
                    throw "Unsupported request_type [$RequestType]"
                }

                $docId = New-CosmosDbDocument -CosmosEndpoint $CosmosEndpoint -DatabaseName $DatabaseName -ContainerName $ContainerName -MasterKey $CosmosKey -Document $enriched
                Write-Output "Created document with ID [$docId]"
            }

            "update" {
                $docId = Update-CosmosDbDocument -CosmosEndpoint $CosmosEndpoint -DatabaseName $DatabaseName -ContainerName $ContainerName -MasterKey $CosmosKey -TicketNumber $TicketNumber -InputDocument $payload
                Write-Output "Updated document with ID [$docId]"
            }

            default {
                Write-Output "Invalid action [$Action]. Use get, post, or update."
                return
            }
        }
    }
    catch {
        Write-Output "Unhandled exception: $($_.Exception.Message)"
    }

    $results = Get-CosmosDbDocument -CosmosEndpoint $CosmosEndpoint -DatabaseName $DatabaseName -ContainerName $ContainerName -MasterKey $CosmosKey -TicketNumber $TicketNumber
    if ($results) {
        Write-Output "Found $($results.Count) document(s) for ticket [$TicketNumber]"
        $jsonPayload = $results[0] | ConvertTo-Json -Depth 10
        Write-Output "[DEBUG]: Final webhook JSON payload:"
        Write-Output $jsonPayload

        Invoke-Webhook -RequestType $RequestType -JsonPayload $jsonPayload
    }
    else {
        Write-Output "[WARN]: No document found after action for ticket [$TicketNumber]"
    }
}

ProcessMainFunction