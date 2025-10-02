$ApiUrl = "https://graph.microsoft.com/v1.0/applications?`$select=id,appId,displayName,passwordCredentials,keyCredentials&`$top=100"
$Headers = @{ 'Content-Type' = 'application/json'; 'Authorization' = "Bearer $BearerToken"; 'ConsistencyLevel' = 'eventual' }
$ValidAppReg = @()
$ValidAppCerts = @()
$TicketDetails = @()

$CurrentDate = Get-Date
$StartDate = $CurrentDate.AddDays(-30)
$ExpirationThreshold = $CurrentDate.AddDays(45)

while ($null -ne $ApiUrl) {
    $ApiArguments = @{
        Uri             = $ApiUrl
        Method          = 'GET'
        Headers         = $Headers
        UseBasicParsing = $true
    }

    try {
        $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
        Write-Warning "SUCCESS: Fetched enterprise applications for $TenantUrl"

        foreach ($App in $ApiResponse.value) {
            if ($App.passwordCredentials -and $App.passwordCredentials.Count -gt 0) {
                $ValidAppReg += $App.displayName
                foreach ($Credential in $App.passwordCredentials) {
                    $EndDate = [datetime]$Credential.endDateTime
                    if ($EndDate -ge $StartDate -and $EndDate -le $ExpirationThreshold) {
                        $TicketDetails += @{
                            "DisplayName"  = $App.displayName
                            "Summary"      = "Azure App Registration secret expiring - $($App.displayName)"
                            "EndDateTime"  = $Credential.endDateTime
                            "InternalNote" = "Secret ID: '$($Credential.keyId)'\nExpiry date: '$($Credential.endDateTime)'\n\nApp registration location: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/$($App.appId)"
                        }
                    }
                }
            }

            if ($App.keyCredentials -and $App.keyCredentials.Count -gt 0) {
                $ValidAppCerts += $App.displayName
                foreach ($Credential in $App.keyCredentials) {
                    $EndDate = [datetime]$Credential.endDateTime
                    if ($EndDate -ge $StartDate -and $EndDate -le $ExpirationThreshold) {
                        $TicketDetails += @{
                            "DisplayName"  = $App.displayName
                            "Summary"      = "Azure App Registration certificate expiring - $($App.displayName)"
                            "EndDateTime"  = $Credential.endDateTime
                            "InternalNote" = "Certificate ID: '$($Credential.keyId)'\nExpiry date: '$($Credential.endDateTime)'\n\nApp registration location: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/$($App.appId)"
                        }
                    }
                }
            }
        }
        $ApiUrl = $ApiResponse.'@odata.nextLink'
    }
    catch {
        Write-Error "Failed to fetch service principals for $($TenantUrl): $_"
        break
    }
}

$TicketDetails | ForEach-Object {
    [PSCustomObject]@{
        DisplayName  = $_.DisplayName
        Summary      = $_.Summary
        EndDateTime  = $_.EndDateTime
        InternalNote = $_.InternalNote
    }
} | Export-Csv -Path "C:\temp\list.csv" -NoTypeInformation -Encoding UTF8
