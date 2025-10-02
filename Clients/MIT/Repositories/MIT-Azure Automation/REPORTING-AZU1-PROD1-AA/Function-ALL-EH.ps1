function Invoke-EHApiCall {
    param (
        [string]$Uri,
        [string]$Authorization,
        [string]$Method = 'GET',
        [hashtable]$Body = $null,
        [string]$ContentType = 'application/json',
        [switch]$Paginate,
        [int]$Retries = 5
    )

    $BaseDelay = 5

    if ($Paginate) {
        $AllItems = @()
        $Page = 1
        $PerPage = 100
        $Done = $false
        while (-not $Done) {
            $Attempt = 0
            do {
                try {
                    $ApiArguments = @{
                        Uri             = "$($Uri)?page_index=$Page&item_per_page=$PerPage"
                        Headers         = @{ 
                            'Authorization' = $Authorization
                            'Accept'        = 'application/json'
                        }
                        Method          = $Method
                        UseBasicParsing = $true
                        TimeoutSec      = 30
                    }
                    if ($Body -and $Method -ne 'GET') {
                        $ApiArguments['Body'] = ($Body | ConvertTo-Json -Depth 10)
                        $ApiArguments['ContentType'] = $ContentType
                    }
                    $Response = Invoke-WebRequest @ApiArguments
                    $StatusCode = $Response.StatusCode
                    $content = $Response.Content
                    if ($content -is [byte[]]) {
                        $content = [System.Text.Encoding]::UTF8.GetString($content)
                    }
                    try {
                        $JsonResponse = $content | ConvertFrom-Json
                        $Data = $JsonResponse.data
                    } catch {
                        $Data = $content
                    }

                    if (200 -le $StatusCode -and $StatusCode -lt 300) {
                        if ($Data -and $Data.items) {
                            $AllItems += $Data.items
                        }
                        if (($null -eq $Data.items) -or ($Data.items.Count -eq 0) -or ($Data.page_index -eq $Data.total_pages)) {
                            $Done = $true
                        } else {
                            $Page++
                        }
                        break
                    } elseif ($StatusCode -eq 429 -or $StatusCode -eq 503) {
                        $RetryAfter = $Response.Headers["Retry-After"]
                        $WaitTime = if ($RetryAfter) { [int]$RetryAfter } else { $BaseDelay * (2 * $Attempt) }
                        Write-Warning "Rate limit exceeded. Retrying in $WaitTime seconds"
                        Start-Sleep -Seconds $WaitTime
                        $Attempt++
                        continue
                    } elseif (400 -le $StatusCode -and $StatusCode -lt 500) {
                        Write-Error "Client error Status: $StatusCode, Response: $content"
                        return @{ Success = $false; StatusCode = $StatusCode; Data = $content }
                    } elseif (500 -le $StatusCode -and $StatusCode -lt 600) {
                        Write-Warning "Server error Status: $StatusCode, attempt $($Attempt + 1) of $Retries"
                        Start-Sleep -Seconds ($BaseDelay * (2 * $Attempt))
                        $Attempt++
                        continue
                    } else {
                        Write-Error "Unexpected response Status: $StatusCode, Response: $content"
                        return @{ Success = $false; StatusCode = $StatusCode; Data = $content }
                    }
                } catch {
                    Write-Error "Exception during API call to $($Uri): $_"
                    return @{ Success = $false; StatusCode = 0; Error = $_.Exception.Message }
                }
            } while ($Attempt -lt $Retries)
            if ($Attempt -ge $Retries) {
                Write-Error "Max retries exceeded for page $Page"
                return @{ Success = $false; StatusCode = 0; Error = "Max retries exceeded" }
            }
        }
        return @{ Success = $true; StatusCode = 200; Data = $AllItems }
    } else {
        $Attempt = 0
        do {
            try {
                $ApiArguments = @{
                    Uri             = $Uri
                    Headers         = @{ 
                        'Authorization' = $Authorization
                        'Accept'        = 'application/json'
                    }
                    Method          = $Method
                    UseBasicParsing = $true
                    TimeoutSec      = 30
                }
                if ($Body -and $Method -ne 'GET') {
                    $ApiArguments['Body'] = ($Body | ConvertTo-Json -Depth 10)
                    $ApiArguments['ContentType'] = $ContentType
                }
                $Response = Invoke-WebRequest @ApiArguments
                $StatusCode = $Response.StatusCode
                $content = $Response.Content
                if ($content -is [byte[]]) {
                    $content = [System.Text.Encoding]::UTF8.GetString($content)
                }
                try {
                    $JsonResponse = $content | ConvertFrom-Json
                    $Data = $JsonResponse
                } catch {
                    $Data = $content
                }

                if (200 -le $StatusCode -and $StatusCode -lt 300) {
                    return @{ Success = $true; StatusCode = $StatusCode; Data = $Data }
                } elseif ($StatusCode -eq 429 -or $StatusCode -eq 503) {
                    $RetryAfter = $Response.Headers["Retry-After"]
                    $WaitTime = if ($RetryAfter) { [int]$RetryAfter } else { $BaseDelay * (2 * $Attempt) }
                    Write-Warning "Rate limit exceeded. Retrying in $WaitTime seconds"
                    Start-Sleep -Seconds $WaitTime
                    $Attempt++
                    continue
                } elseif (400 -le $StatusCode -and $StatusCode -lt 500) {
                    Write-Error "Client error Status: $StatusCode, Response: $content"
                    return @{ Success = $false; StatusCode = $StatusCode; Data = $content }
                } elseif (500 -le $StatusCode -and $StatusCode -lt 600) {
                    Write-Warning "Server error Status: $StatusCode, attempt $($Attempt + 1) of $Retries"
                    Start-Sleep -Seconds ($BaseDelay * (2 * $Attempt))
                    $Attempt++
                    continue
                } else {
                    Write-Error "Unexpected response Status: $StatusCode, Response: $content"
                    return @{ Success = $false; StatusCode = $StatusCode; Data = $content }
                }
            } catch {
                Write-Error "Exception during API call to $($Uri): $_"
                return @{ Success = $false; StatusCode = 0; Error = $_.Exception.Message }
            }
            break
        } while ($Attempt -lt $Retries)
        return @{ Success = $false; StatusCode = 0; Error = "Max retries exceeded" }
    }
}

function Get-EHAuthorization {
    param (
        [string]$EHclient_Id,
        [string]$EHclient_secret,
        [string]$EHcode,
        [string]$EHrefresh_token,
        [string]$EHRedirectUri
    )

    $Uri = "https://oauth.employmenthero.com/oauth2/token?client_id=$($EHclient_Id)&client_secret=$($EHclient_secret)&grant_type=refresh_token&code=$($EHcode)&redirect_uri=$($EHRedirectUri)&refresh_token=$($EHrefresh_token)"
    
    try {
        $ApiArguments = @{
            Uri             = $Uri
            Method          = 'POST'
            UseBasicParsing = $true
            TimeoutSec      = 30
        }
        
        $ApiResponse = Invoke-WebRequest @ApiArguments
        $ApiResponseContent = $ApiResponse.Content | ConvertFrom-Json
        
        if ($null -eq $ApiResponseContent.access_token) {
            Write-Error -Message "Access token is null."
            return $null
        }
        
        $AccessToken = $ApiResponseContent.access_token
        $Authorization = "Bearer $AccessToken"
        return $Authorization
    }
    catch {
        Write-Error -Message "Failed to retrieve access token: $_"
        return $null
    }
}

function Get-Employee {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EHOrganizationId,
        [Parameter(Mandatory = $true)]
        [string]$EHAuthorization,
        [Parameter(Mandatory = $true)]
        [string]$EmployeeId
    )
    try {
        if ([string]::IsNullOrEmpty($EHOrganizationId)) {
            Write-Error "EHOrganizationId cannot be null or empty"
            return $null
        }
        if ([string]::IsNullOrEmpty($EHAuthorization)) {
            Write-Error "EHAuthorization cannot be null or empty"
            return $null
        }
        if ([string]::IsNullOrEmpty($EmployeeId)) {
            Write-Error "EmployeeId cannot be null or empty"
            return $null
        }

        $Uri = "https://api.employmenthero.com/api/v1/organisations/$EHOrganizationId/employees/$EmployeeId"
        Write-Verbose "Fetching employee data from $Uri"
        $Result = Invoke-EHApiCall -Uri $Uri -Authorization $EHAuthorization -Method 'GET'

        if ($Result.Success) {
            if ($null -eq $Result.Data) {
                Write-Error "API returned success but no data for Employee $EmployeeId"
                return $null
            }
            Write-Host "DEBUG: Raw API response:" ($Result.Data | ConvertTo-Json -Depth 10)
            if ($Result.Data.data) {
                Write-Verbose "Successfully retrieved employee data for ID: $EmployeeId"
                return $Result.Data.data
            } else {
                Write-Error "API returned success but unexpected data format for Employee $EmployeeId"
                return $null
            }
        } else {
            Write-Error "Failed to fetch employee details for ID: $EmployeeId"
            return $null
        }
    }
    catch {
        Write-Error "Exception in Get-Employee: $_"
        return $null
    }
}

function Get-Teams {
    param (
        $EHOrganizationId,
        $EHAuthorization
    )
    try {
        $Uri = "https://api.employmenthero.com/api/v1/organisations/$EHOrganizationId/teams"
        $Result = Invoke-EHApiCall -Uri $Uri -Authorization $EHAuthorization -Paginate
        if ($Result.Success) {
            Write-Host "Total teams fetched: $($Result.Data.Count)"
            return $Result.Data
        } else {
            Write-Error -Message "Failed to fetch teams. Error: $($Result.Error)"
            return $null
        }
    }
    catch {
        Write-Error -Message "Exception in Get-Teams: $_"
        return $null
    }
}

function Get-TeamMembers {
    param (
        $EHOrganizationId,
        $EHAuthorization,
        $TeamId
    )
    
    $Uri = "https://api.employmenthero.com/api/v1/organisations/$EHOrganizationId/teams/$TeamId/employees"
    $Result = Invoke-EHApiCall -Uri $Uri -Authorization $EHAuthorization -Paginate
    if ($Result.Success) {
        $Count = $Result.Data.Count
        if ($Count -eq 0) {
            Write-Host "Team $($TeamId) has no members."
        } else {
            Write-Host "Total team members fetched for team $($TeamId): $Count"
        }
        return $Result.Data
    } else {
        Write-Error -Message "Failed to fetch team members for team $TeamId. Error: $($Result.Error)"
        return @()
    }
}

function Get-EmployeeTeamNames {
    param (
        $EHOrganizationId,
        $EHAuthorization,
        $EmployeeId,
        $Teams
    )
    try {
        $TeamMembersWithTeamName = @()
        $TeamCount = $Teams.Count
        $Current = 0
        foreach ($TeamItem in $Teams) {
        $Current++
        Write-Host "Processing team $Current of $($TeamCount): $($TeamItem.name) ($($TeamItem.id))"
        $TeamId = $TeamItem.id

        $startTime = Get-Date
        $TeamMembers = Get-TeamMembers -EHOrganizationId $EHOrganizationId -TeamId $TeamId -EHAuthorization $EHAuthorization
        $duration = (Get-Date) - $startTime
        Write-Host "Fetched $($TeamMembers.Count) members for team $($TeamItem.name) in $($duration.TotalSeconds) seconds"

        if ($null -eq $TeamMembers -or $TeamMembers.Count -eq 0) { continue }

            foreach ($TeamMember in $TeamMembers) {
                if ($TeamMember.id -eq $EmployeeId) {
                    $TeamMembersWithTeamName += [PSCustomObject]@{
                        TeamMember = $TeamMember
                        TeamName   = $TeamItem.name
                    }
                }
            }
        }
        if ($TeamMembersWithTeamName.Count -eq 0) {
            Write-Error -Message "No team names found for employee."
            return $null
        }
        return $TeamMembersWithTeamName | Select-Object -ExpandProperty TeamName
    }
    catch {
        Write-Error -Message "Failed to fetch employee team names. Error: $_"
        return $null
    }
}

function Get-EmployeeCustomfields {
    param (
        $EHOrganizationId,
        $EHAuthorization,
        $EmployeeId
    )

    $Uri = "https://api.employmenthero.com/api/v1/organisations/$EHOrganizationId/employees/$EmployeeId/custom_fields"
    $Result = Invoke-EHApiCall -Uri $Uri -Authorization $EHAuthorization -Paginate

    if ($Result.Success) {
        return $Result.Data
    } else {
        Write-Error -Message "Failed to fetch employee custom fields for $EmployeeId. Error: $($Result.Error)"
        return $null
    }
}

function Get-EmployeeEmploymentHistory {
    param (
        $EHOrganizationId,
        $EHAuthorization,
        $EmployeeId
    )

    $Uri = "https://api.employmenthero.com/api/v1/organisations/$EHOrganizationId/employees/$EmployeeId/employment_histories"
    $Result = Invoke-EHApiCall -Uri $Uri -Authorization $EHAuthorization -Paginate

    if ($Result.Success) {
        return $Result.Data
    } else {
        Write-Error -Message "Failed to fetch employee employment history for $EmployeeId. Error: $($Result.Error)"
        return $null
    }
}