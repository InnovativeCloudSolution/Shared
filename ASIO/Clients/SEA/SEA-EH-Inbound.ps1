param (
    [Parameter(Mandatory = $false)]
    [object]$WebhookData
)

. .\DEV\UserManagement\Function-EH-Common.ps1
. .\DEV\UserManagement\Function-SEA-EH-Common.ps1

function Send-Webhook {
    param (
        [string]$EventType,
        [string]$EmployeeId
    )
    
    $webhookUrl = "https://1ce978e3-0165-4625-bc15-bdcdfaeb9c7e.webhook.ae.azure-automation.net/webhooks?token=1K7hCJlpvQNta%2fW%2bO2WjxAPAOHllJ6aElDxD4jMuAzY%3d"
    
    $webhookPayload = @{
        data  = @{
            id = $EmployeeId
        }
        event = $EventType
    }

    $jsonPayload = $webhookPayload | ConvertTo-Json

    try {
        Invoke-RestMethod -Uri $webhookUrl -Method Post -Body $jsonPayload -ContentType "application/json"
        Write-MessageLog "Webhook sent successfully for event: $EventType and employee ID: $EmployeeId."
    }
    catch {
        Write-MessageLog "Failed to send webhook: $_" -LogType "Error"
    }
}

function Start-InboundCheck {
    try {
        Write-MessageLog "Starting script execution."

        Write-MessageLog "The Webhook Header"
        Write-MessageLog $WebhookData.RequestHeader

        Write-MessageLog "The Webhook Header Message"
        Write-MessageLog $WebhookData.RequestHeader.Message

        Write-MessageLog 'The Webhook Name'
        Write-MessageLog $WebhookData.WebhookName

        Write-MessageLog 'The Webhook Request Body'
        Write-MessageLog $WebhookData.RequestBody

        Write-MessageLog "Getting PreCheck details."
        $PreCheckData = Get-PreCheck -WebhookData $WebhookData.RequestBody

        if ($PreCheckData) {
            Write-MessageLog "Webhook triggered action."

            $Result = $WebhookData.RequestBody | ConvertFrom-Json
            $EmployeeData = $Result.data
            $EmployeeId = $EmployeeData.id
            $Event = $Result.event

            if ($Event -eq "employee_created") {
                Write-MessageLog "Sending webhook for employee creation."
                Send-Webhook -EventType "employee_created" -EmployeeId $EmployeeId
            }
            elseif ($Event -eq "employee_onboarded") {
                Write-MessageLog "Sending webhook for employee onboarding."
                Send-Webhook -EventType "employee_onboarded" -EmployeeId $EmployeeId
            }
            else {
                Write-MessageLog "Unknown event: $Event" -LogType "Warning"
            }
        }
        else {
            Write-MessageLog "Scheduled run triggered."

            Write-MessageLog "Getting Secrets from Azure Key Vault."
            $AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'
            $EHSecrets = Get-EH-Secrets -AzKeyVaultName $AzKeyVaultName

            Write-MessageLog "Setting EH Secret Variables."
            $EHclient_Id = $EHSecrets.EHclient_Id
            $EHclient_secret = $EHSecrets.EHclient_secret
            $EHcode = $EHSecrets.EHcode
            $EHrefresh_token = $EHSecrets.EHrefresh_token
            $EHOrganizationId = $EHSecrets.EHOrganizationId
            $EHRedirectUri = $EHSecrets.EHRedirectUri

            $EHAuthorization = Get-EH-Authorization -EHclient_Id $EHclient_Id -EHclient_secret $EHclient_secret -EHcode $EHcode -EHrefresh_token $EHrefresh_token -EHRedirectUri $EHRedirectUri

            Write-MessageLog "Fetching all employees."
            $Employees = Get-EHAllEmployees -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization
            $CurrentTime = Get-Date
            $TimeWindow = $CurrentTime.AddHours(3)

            foreach ($Employee in $Employees) {
                $EmployeeId = $Employee.Id
                $EmployeeCustomfields = (Get-EH-EmployeeCustomfields -EmployeeId $EmployeeId -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization).items

                foreach ($EmployeeCustomfield in $EmployeeCustomfields) {
                    if ($EmployeeCustomfield.name -eq "Offboarding" -and $EmployeeCustomfield.value) {
                        try {
                            $ParsedDate = [datetime]::ParseExact($EmployeeCustomfield.value, "dd/MM/yy HHmm", $null)

                            if ($ParsedDate -ge $CurrentTime -and $ParsedDate -le $TimeWindow) {
                                Write-MessageLog "Employee: $($Employee.first_name) $($Employee.last_name)"
                                Write-MessageLog "Offboarding Date: $($EmployeeCustomfield.value)"
                                Write-MessageLog "This offboarding date falls within the next 3 hours."

                                Send-Webhook -EventType "employee_offboarding" -EmployeeId $EmployeeId
                            }
                        }
                        catch {
                            Write-MessageLog "Failed to parse offboarding date for Employee: $($Employee.first_name) $($Employee.last_name). Offboarding value: $($EmployeeCustomfield.value)" -LogType "Error"
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-MessageLog "An error occurred: $_" "ERROR"
    }

    Write-MessageLog "Script execution completed."
}

Start-InboundCheck