## VARIABLES ##
$Url = $env:APIKEY_DESKDIRECTOR_URL
$ApiKey = $env:APIKEY_DESKDIRECTOR_SECRETKEY
$UserId = $env:APIKEY_DESKDIRECTOR_USERID

## FUNCTIONS ##
function New-DeskDirectorBearerToken {
    param (
        [Parameter(Mandatory)][int]$TicketId
    )

    # Build API variables
    $Arguments = @{
        Uri = "$($Url)/api/v2/user/member/$($UserId)/userkey"
        Method = 'GET'
        Headers = @{
            Authorization = "DdApi $($ApiKey)"
        }
        UseBasicParsing = $true
        ContentType = 'application/json'
    }

    # Get bearer token
    try {
        $Request = Invoke-WebRequest @Arguments | ConvertFrom-Json
        Write-Host "[#$($TicketId)] New-DeskDirectorBearerToken: Bearer token successfully generated."
        return $Request.userKey
    }
    catch {
        Write-Error "[#$($TicketId)] New-DeskDirectorBearerToken: Unable to generate bearer token : $($_)"
    }
}

function Get-DeskDirectorFormResult {
    param (
        [Parameter(Mandatory)][int]$TicketId,
        [Parameter(Mandatory)][int]$FormEntityId,
        [Parameter(Mandatory)][int]$ResultEntityId
    )

    # Get bearer token
    $BearerToken = New-DeskDirectorBearerToken -TicketId $TicketId
    
    # Build API variables
    $Arguments = @{
        Uri = "$($Url)/api/v2/ddform/forms/$($FormEntityId)/results/$($ResultEntityId)"
        Method = 'GET'
        Headers = @{
            Authorization = "DdAccessToken $($BearerToken)"
        }
        UseBasicParsing = $true
        ContentType = 'application/json'
    }

    # Get form result details
    try {
        $Request = Invoke-WebRequest @Arguments | ConvertFrom-Json
        Write-Warning "[#$($TicketId)] Get-DeskDirectorFormResult: Fetched form result for $($ResultEntityId)."
        return $Request.result
    }
    catch {
        Write-Error "[#$($TicketId)] Get-DeskDirectorFormResult: Unable to fetch form result for $($ResultEntityId) : $($_)"
    }
}

function Get-DeskDirectorFormResultByTicket {
    param (
        [Parameter(Mandatory)][int]$TicketId,
        [Parameter(Mandatory)][int]$FormEntityId
    )

    # Get bearer token
    $BearerToken = New-DeskDirectorBearerToken -TicketId $TicketId

    # Get today and yesterday's date
    $CurrentDateUnformatted = Get-Date
    $CurrentDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
    $YesterdayDateUnformatted = $CurrentDateUnformatted.AddDays(-1)
    $YesterdayDate = Get-Date -Date $YesterdayDateUnformatted -Format "yyyy-MM-ddTHH:mm:ss.fffZ"
    
    # Build API variables
    $Arguments = @{
        Uri = "$($Url)/api/v2/ddform/forms/$($FormEntityId)/results/find"
        Method = 'POST'
        Headers = @{
            Authorization = "DdAccessToken $($BearerToken)"
        }
        Body = @{
            skip = 0
            take = 1000
            timeRange = @{
                startTime = $YesterdayDate
                endTime = $CurrentDate
            }
        } | ConvertTo-Json -Depth 100
        UseBasicParsing = $true
        ContentType = 'application/json'
    }

    # Get all form submissions for the last day
    try {
        $Submissions = Invoke-WebRequest @Arguments | ConvertFrom-Json
        Write-Warning "[#$($TicketId)] Get-DeskDirectorFormResultByTicket: Fetched form results for $($FormEntityId)."
    }
    catch {
        Write-Error "[#$($TicketId)] Get-DeskDirectorFormResultByTicket: Unable to fetch form results for $($FormEntityId) : $($_)"
    }

    # Sort through to find matching submission
    foreach ($Result in $Submissions.results) {
        if ($TicketId -eq $Result.ticket.entityId) {
            Write-Warning "[#$($TicketId)] Get-DeskDirectorFormResultByTicket: Located result ID for #$($TicketId) - $($Result.entityId)"
            $FormResult = Get-DeskDirectorFormResult -TicketId $TicketId -FormEntityId $FormEntityId -ResultEntityId $Result.entityId
        }
    }

    # Return result if located
    if ($null -ne $FormResult) {
        return $FormResult
    }
}