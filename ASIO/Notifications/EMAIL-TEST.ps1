param (
    [string]$TenantId = "5792a6c1-f4fe-466b-b97c-10eaf4fb3122",
    [string]$ClientId = "e05b08c8-78aa-4bc2-9d3c-f99bca9bf5f6",
    [string]$ClientSecret = "6a._9At5Yh3~-Bg1-j1iaHqKRR_6TsiEOm",
    [string]$FromAddress = "support@manganoit.com.au",
    [string]$ToAddress = "juan.moredo@manganoit.com.au",
    [string]$Subject = "Test",
    [string]$HtmlFilePath = "C:\Scripts\Repositories\Private\DEV\ASIO\Notifications\Ticket Notifications\MIT-Ticket-Notification-Waiting.html"
)


function Get-AccessToken {
    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $tokenBody = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

    try {
        $response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $tokenBody -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    } catch {
        Write-Error "Failed to obtain access token: $($_.Exception.Message)"
        exit 1
    }
}

function Send-GraphEmail {
    param (
        [string]$Token,
        [string]$From,
        [string]$To,
        [string]$Subject,
        [string]$HtmlBody
    )

    $uri = "https://graph.microsoft.com/v1.0/users/$From/sendMail"

    $body = @{
        message = @{
            subject = $Subject
            body = @{
                contentType = "HTML"
                content     = $HtmlBody
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $To
                    }
                }
            )
        }
        saveToSentItems = $true
    }

    try {
        Invoke-RestMethod -Method Post -Uri $uri -Headers @{ Authorization = "Bearer $Token" } -Body ($body | ConvertTo-Json -Depth 10 -Compress) -ContentType "application/json"
        Write-Host "Email sent successfully to $To"
    } catch {
        Write-Error "Failed to send email: $($_.Exception.Message)"
    }
}

if (-Not (Test-Path $HtmlFilePath)) {
    Write-Error "HTML file not found at $HtmlFilePath"
    exit 1
}

$htmlContent = Get-Content -Path $HtmlFilePath -Raw
$accessToken = Get-AccessToken
Send-GraphEmail -Token $accessToken -From $FromAddress -To $ToAddress -Subject $Subject -HtmlBody $htmlContent