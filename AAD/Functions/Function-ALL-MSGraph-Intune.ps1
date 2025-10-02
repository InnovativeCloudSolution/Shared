function Get-IntuneConnectorHealth {
    try {
        $url = "https://graph.microsoft.com/beta/deviceManagement/ndesConnectors"
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers @{ Authorization = "Bearer $($AccessToken)" }

        if ($response -and $response.value -and $response.value.Count -gt 0) {
            $nonActive = $response.value | Where-Object { $_.state -ne 'active' }
            return $nonActive
        } else {
            return @()
        }
    }
    catch {
        Write-ErrorLog -Message "Failed to query Intune Connector Health. Error: $($_ | Out-String)"
        return @()
    }
}

function Get-ApplePushNotificationCertificate {
    try {
        $url = "https://graph.microsoft.com/beta/deviceManagement/applePushNotificationCertificate"
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers @{ Authorization = "Bearer $($AccessToken)" }

        if ($response -ne $null) {
            $expiringSoon = $response | Where-Object { ([datetime]$_.expirationDateTime -lt (Get-Date).AddDays(61)) }
            return $expiringSoon
        } else {
            return @()
        }
    }
    catch {
        if ($_.Exception.Response.StatusCode -eq 404) {
            return @()
        } else {
            Write-ErrorLog -Message "Failed to check Apple Push Notification Certificate. Error: $($_ | Out-String)"
            return @()
        }
    }
}

function Get-AppleVppTokens {
    try {
        $url = "https://graph.microsoft.com/beta/deviceAppManagement/vppTokens"
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers @{ Authorization = "Bearer $($AccessToken)" }

        if ($response -and $response.value -and $response.value.Count -gt 0) {
            $invalidTokens = $response.value | Where-Object { $_.lastSyncStatus -eq 'failed' -or $_.state -ne 'valid' }
            return $invalidTokens
        } else {
            return @()
        }
    }
    catch {
        Write-ErrorLog -Message "Failed to check Apple VPP Tokens. Error: $_"
        return @()
    }
}

function Get-AppleDepTokens {
    try {
        $url = "https://graph.microsoft.com/beta/deviceManagement/depOnboardingSettings"
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers @{ Authorization = "Bearer $($AccessToken)" }

        if ($response -and $response.value -and $response.value.Count -gt 0) {
            $invalidTokens = $response.value | Where-Object { $_.lastSyncErrorCode -ne 0 -or ([datetime]$_.tokenExpirationDateTime -lt (Get-Date).AddDays(61)) }
            return $invalidTokens
        } else {
            return @()
        }
    }
    catch {
        Write-ErrorLog -Message "Failed to check Apple DEP Tokens. Error: $_"
        return @()
    }
}

function Get-ManagedGooglePlay {
    try {
        $url = "https://graph.microsoft.com/beta/deviceManagement/androidManagedStoreAccountEnterpriseSettings"
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers @{ Authorization = "Bearer $($AccessToken)" }

        $issues = $response | Where-Object {
            $_.bindStatus -ne 'boundAndValidated' -or
            $_.lastAppSyncStatus -ne 'success'
        }

        $unconfigured = $response | Where-Object {
            $_.enrollmentTarget -eq 'none' -and
            $_.deviceOwnerManagementEnabled -eq $false
        }

        if ($unconfigured.Count -gt 0) {
            return @()
        }

        if ($issues.Count -gt 0) {
            return $issues
        }
        return @()
    }
    catch {
        Write-ErrorLog -Message "Failed to check Managed Google Play settings. Error: $($_ | Out-String)"
        return @()
    }
}

function Get-Autopilot {
    try {
        $url = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotSettings"
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers @{ Authorization = "Bearer $($AccessToken)" }
        $issues = $response | Where-Object { $_.syncStatus -ne 'completed' -and $_.syncStatus -ne 'inProgress' }
        return $issues
    }
    catch {
        Write-ErrorLog -Message "Failed to check Autopilot. Error: $_"
        return @()
    }
}

function Get-MobileThreatDefenseConnectors {
    try {
        $url = "https://graph.microsoft.com/beta/deviceManagement/mobileThreatDefenseConnectors"
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers @{ Authorization = "Bearer $($AccessToken)" }

        if ($response -and $response.value -and $response.value.Count -gt 0) {
            $issues = $response.value | Where-Object { $_.partnerState -ne 'enabled' -and $_.partnerState -ne 'available' }
            return $issues
        } else {
            return @()
        }
    } catch {
        Write-ErrorLog -Message "Failed to check Mobile Threat Defense Connectors. Error: $($_ | Out-String)"
        return @()
    }
}
