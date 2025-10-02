
function New-CWM-TicketJson {
    param (
        [string]$Summary,
        [string]$Description,
        [string]$Status,
        [string]$Priority
    )
    
    $ticketObject = @{
        summary     = $Summary
        description = $Description
        status      = $Status
        priority    = $Priority
    }
    
    return $ticketObject | ConvertTo-Json -Depth 3
}

function New-CWM-ConfigItemJson {
    param (
        [string]$Name,
        [string]$Type,
        [string]$Status
    )
    
    $configItemObject = @{
        name   = $Name
        type   = $Type
        status = $Status
    }
    
    return $configItemObject | ConvertTo-Json -Depth 3
}

function New-CWM-ContactJson {
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$EmailAddress,
        [string]$PhoneNumber
    )
    
    $contactObject = @{
        firstName   = $FirstName
        lastName    = $LastName
        email       = $EmailAddress
        phoneNumber = $PhoneNumber
    }
    
    return $contactObject | ConvertTo-Json -Depth 3
}

Export-ModuleMember -Function New-CWM-TicketJson, New-CWM-ConfigItemJson, New-CWM-ContactJson
