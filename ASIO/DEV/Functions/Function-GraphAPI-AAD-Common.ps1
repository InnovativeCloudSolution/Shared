
# User Management: Create, read, update, and delete users
function Reset-GraphUserPassword {
    param (
        [string]$userPrincipalName,
        [string]$newPassword
    )
    $body = @{
        passwordProfile = @{
            password = $newPassword
            forceChangePasswordNextSignIn = $true
        }
    } | ConvertTo-Json

    Invoke-GraphRequest -method PATCH -url "https://graph.microsoft.com/v1.0/users/$userPrincipalName" -body $body
}

function Block-GraphUserSignIn {
    param (
        [string]$userPrincipalName
    )
    $body = @{
        accountEnabled = $false
    } | ConvertTo-Json

    Invoke-GraphRequest -method PATCH -url "https://graph.microsoft.com/v1.0/users/$userPrincipalName" -body $body
}

# Groups: Manage groups and their memberships
function Remove-GraphUserFromGroup {
    param (
        [string]$userPrincipalName,
        [string]$groupName
    )
    $groupId = (Invoke-GraphRequest -method GET -url "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$groupName'" | Select-Object -ExpandProperty id)
    $userId = (Invoke-GraphRequest -method GET -url "https://graph.microsoft.com/v1.0/users/$userPrincipalName" | Select-Object -ExpandProperty id)
    
    Invoke-GraphRequest -method DELETE -url "https://graph.microsoft.com/v1.0/groups/$groupId/members/$userId/`$ref"
}

# Roles: Assign and manage roles
function Remove-GraphUserFromRole {
    param (
        [string]$userPrincipalName,
        [string]$roleName
    )
    $roleId = (Invoke-GraphRequest -method GET -url "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=displayName eq '$roleName'" | Select-Object -ExpandProperty id)
    $userId = (Invoke-GraphRequest -method GET -url "https://graph.microsoft.com/v1.0/users/$userPrincipalName" | Select-Object -ExpandProperty id)
    
    Invoke-GraphRequest -method DELETE -url "https://graph.microsoft.com/v1.0/directoryRoles/$roleId/members/$userId/`$ref"
}

# Applications, Devices, and other Azure AD-related functions would go here...

# Generic Graph API request function
function Invoke-GraphRequest {
    param (
        [string]$method,
        [string]$url,
        [string]$body = $null
    )
    
    $token = Get-GraphAPIToken
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type"  = "application/json"
    }

    if ($method -eq "GET" -or $method -eq "DELETE") {
        Invoke-RestMethod -Uri $url -Method $method -Headers $headers
    } else {
        Invoke-RestMethod -Uri $url -Method $method -Headers $headers -Body $body
    }
}

# Function to get the Graph API token
function Get-GraphAPIToken {
    # Implement your token retrieval logic here (OAuth, client credentials, etc.)
    # For example, using client credentials flow:
    $tenantId = "your-tenant-id"
    $clientId = "your-client-id"
    $clientSecret = "your-client-secret"
    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    $response = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $body -ContentType "application/x-www-form-urlencoded"
    return $response.access_token
}
