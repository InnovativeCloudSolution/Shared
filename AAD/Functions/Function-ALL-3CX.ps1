function Authenticate3CX {

    try {
        $secrets = Get-MSGraph-Secrets -AzKeyVaultName $AzKeyVaultName -SecretNames @($Secret3CX)
        $password = $secrets[$Secret3CX]

        $headers = @{ "Content-Type" = "application/json" }

        $body = @{
            SecurityCode = ""
            Password     = "$password"
            Username     = "7000"
        } | ConvertTo-Json

        $response = Invoke-RestMethod "$3CXURL/webclient/api/Login/GetAccessToken" -Method "POST" -Headers $headers -Body $body
        $Global:ACCESS_TOKEN = $response.token.access_token
        $Global:headers = @{
            "Content-Type"  = "application/json"
            "Authorization" = "Bearer $ACCESS_TOKEN"
        }
    }
    catch {
        Write-ErrorLog "Failed to authenticate with 3CX: $_"
        throw
    }
}

function Authenticate3CXtest {

    try {

        $headers = @{ "Content-Type" = "application/json" }

        $body = @{
            SecurityCode = ""
            Password     = "9fcb56ae1eRW!a"
            Username     = "7000"
        } | ConvertTo-Json

        $response = Invoke-RestMethod "$3CXURL/webclient/api/Login/GetAccessToken" -Method "POST" -Headers $headers -Body $body
        $Global:ACCESS_TOKEN = $response.token.access_token
        $Global:headers = @{
            "Content-Type"  = "application/json"
            "Authorization" = "Bearer $ACCESS_TOKEN"
        }
    }
    catch {
        Write-ErrorLog "Failed to authenticate with 3CX: $_"
        throw
    }
}