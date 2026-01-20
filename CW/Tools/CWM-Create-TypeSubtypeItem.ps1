$CsvPath = ".data\Input\Connectwise Templates\CWPSA-Boards-TypeSubtypeItem-Template.csv"
$CWMUrl = "https://api-aus.myconnectwise.net"
$ApiVersion = "v4_6_release/apis/3.0"
$CompanyId = "dropbearit"
$PublicKey = "xAVcYWO20x5dRyG7"
$PrivateKey = "QUC1zTaMuUXbiJqX"
$ClientId = "1748c7f0-976c-4205-afa1-9bc9e1533565"

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath = ".logs\CWM-TypeSubtypeItem-Log_$timestamp.txt"

$script:ErrorCount = 0
$script:WarningCount = 0
$script:SuccessCount = 0

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    Add-Content -Path $script:LogPath -Value $logEntry
    
    switch ($Level) {
        "ERROR" {
            Write-Host $Message -ForegroundColor Red
            $script:ErrorCount++
        }
        "WARNING" {
            Write-Host $Message -ForegroundColor Yellow
            $script:WarningCount++
        }
        "SUCCESS" {
            Write-Host $Message -ForegroundColor Green
            $script:SuccessCount++
        }
        default {
            Write-Host $Message
        }
    }
}

function Connect-CWM {
    param(
        [string]$CWMUrl,
        [string]$ApiVersion,
        [string]$CompanyId,
        [string]$PublicKey,
        [string]$PrivateKey,
        [string]$ClientId
    )
    
    $authString = "$CompanyId+$PublicKey`:$PrivateKey"
    $encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))
    
    $script:CWMHeaders = @{
        'Authorization' = "Basic $encodedAuth"
        'clientId' = $ClientId
        'Content-Type' = 'application/json'
    }
    
    $script:CWMBaseUrl = "$CWMUrl/$ApiVersion"
    
    Write-Log "Connected to ConnectWise Manage at $CWMUrl" -Level "SUCCESS"
}

function Get-CWMBoard {
    param(
        [string]$BoardName
    )
    
    $uri = "$script:CWMBaseUrl/service/boards?conditions=name='$BoardName'"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            Write-Log "Board '$BoardName' not found" -Level "ERROR"
            return $null
        }
    } catch {
        Write-Log "Failed to retrieve board: $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMType {
    param(
        [int]$BoardId,
        [string]$TypeName
    )
    
    $escapedName = $TypeName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/types?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            return $null
        }
    } catch {
        Write-Log "Failed to retrieve type: $_" -Level "ERROR"
        return $null
    }
}

function New-CWMType {
    param(
        [int]$BoardId,
        [string]$TypeName
    )
    
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/types"
    
    $body = @{
        name = $TypeName
    } | ConvertTo-Json
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Log "Created Type: $TypeName" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to create type '$TypeName': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMSubtype {
    param(
        [int]$BoardId,
        [int]$TypeId,
        [string]$SubtypeName
    )
    
    $escapedName = $SubtypeName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/subtypes?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            foreach ($subtype in $response) {
                if ($subtype.typeAssociationIds -contains $TypeId) {
                    return $subtype
                }
            }
            return $null
        } else {
            return $null
        }
    } catch {
        Write-Log "Failed to retrieve subtype: $_" -Level "ERROR"
        return $null
    }
}

function New-CWMSubtype {
    param(
        [int]$BoardId,
        [int]$TypeId,
        [string]$SubtypeName
    )
    
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/subtypes"
    
    $body = @{
        name = $SubtypeName
        typeAssociationIds = @($TypeId)
    } | ConvertTo-Json
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Log "  Created Subtype: $SubtypeName" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to create subtype '$SubtypeName': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMItem {
    param(
        [int]$BoardId,
        [int]$TypeId,
        [int]$SubtypeId,
        [string]$ItemName
    )
    
    $escapedName = $ItemName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/items?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            return $null
        }
    } catch {
        Write-Log "Failed to retrieve item: $_" -Level "ERROR"
        return $null
    }
}

function New-CWMItem {
    param(
        [int]$BoardId,
        [int]$TypeId,
        [int]$SubtypeId,
        [string]$ItemName
    )
    
    $escapedName = $ItemName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $getUri = "$script:CWMBaseUrl/service/boards/$BoardId/items?conditions=$encodedCondition"
    $existingItem = $null
    
    try {
        $existingItems = Invoke-RestMethod -Uri $getUri -Headers $script:CWMHeaders -Method Get
        if ($existingItems.Count -gt 0) {
            $existingItem = $existingItems[0]
        }
    } catch {}
    
    if (-not $existingItem) {
        $createUri = "$script:CWMBaseUrl/service/boards/$BoardId/items"
        $body = @{
            name = $ItemName
        } | ConvertTo-Json
        
        try {
            $existingItem = Invoke-RestMethod -Uri $createUri -Headers $script:CWMHeaders -Method Post -Body $body
        } catch {
            Write-Log "Failed to create item '$ItemName': $_" -Level "ERROR"
            return $null
        }
    }
    
    Write-Log "    Created Item: $ItemName" -Level "SUCCESS"
    return $existingItem
}

Write-Log "=========================================" -Level "INFO"
Write-Log "ConnectWise Manage Type/Subtype/Item Creator" -Level "INFO"
Write-Log "Log File: $LogPath" -Level "INFO"
Write-Log "=========================================" -Level "INFO"
Write-Log "" -Level "INFO"

if (-not (Test-Path $CsvPath)) {
    Write-Log "CSV file not found: $CsvPath" -Level "ERROR"
    exit 1
}

Write-Log "CSV file loaded: $CsvPath" -Level "SUCCESS"

Connect-CWM -CWMUrl $CWMUrl -ApiVersion $ApiVersion -CompanyId $CompanyId -PublicKey $PublicKey -PrivateKey $PrivateKey -ClientId $ClientId

$csvData = Import-Csv -Path $CsvPath
Write-Log "Total CSV entries: $($csvData.Count)" -Level "INFO"

$boardGroups = $csvData | Group-Object -Property Board
Write-Log "Total boards to process: $($boardGroups.Count)" -Level "INFO"
Write-Log "" -Level "INFO"

foreach ($boardGroup in $boardGroups) {
    $boardName = $boardGroup.Name
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "Processing Board: $boardName" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    
    $board = Get-CWMBoard -BoardName $boardName
    
    if (-not $board) {
        Write-Log "Board '$boardName' not found. Skipping..." -Level "ERROR"
        continue
    }
    
    $boardId = $board.id
    Write-Log "Board ID: $boardId" -Level "INFO"
    Write-Log "Entries to process: $($boardGroup.Group.Count)" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    $processedTypes = @{}
    $processedSubtypes = @{}
    
    foreach ($row in $boardGroup.Group) {
        $typeName = $row.Type.Trim()
        $subtypeName = $row.Subtype.Trim()
        $itemName = $row.Item.Trim()
        
        if ([string]::IsNullOrWhiteSpace($typeName) -or [string]::IsNullOrWhiteSpace($subtypeName) -or [string]::IsNullOrWhiteSpace($itemName)) {
            Write-Log "Skipping row with missing data: Type='$typeName', Subtype='$subtypeName', Item='$itemName'" -Level "WARNING"
            continue
        }
    
    if (-not $processedTypes.ContainsKey($typeName)) {
        $type = Get-CWMType -BoardId $boardId -TypeName $typeName
        
            if (-not $type) {
                $type = New-CWMType -BoardId $boardId -TypeName $typeName
            } else {
                Write-Log "Type already exists: $typeName" -Level "INFO"
            }
            
            if ($type) {
                $processedTypes[$typeName] = $type
            } else {
                Write-Log "Failed to get or create type: $typeName" -Level "ERROR"
                continue
            }
        }
    
    $typeId = $processedTypes[$typeName].id
    $subtypeKey = "$typeName|$subtypeName"
    
    if (-not $processedSubtypes.ContainsKey($subtypeKey)) {
        $subtype = Get-CWMSubtype -BoardId $boardId -TypeId $typeId -SubtypeName $subtypeName
        
            if (-not $subtype) {
                $subtype = New-CWMSubtype -BoardId $boardId -TypeId $typeId -SubtypeName $subtypeName
            } else {
                Write-Log "  Subtype already exists: $subtypeName" -Level "INFO"
            }
            
            if ($subtype) {
                $processedSubtypes[$subtypeKey] = $subtype
            } else {
                Write-Log "Failed to get or create subtype: $subtypeName" -Level "ERROR"
                continue
            }
        }
    
    $subtypeId = $processedSubtypes[$subtypeKey].id
    
    $item = Get-CWMItem -BoardId $boardId -TypeId $typeId -SubtypeId $subtypeId -ItemName $itemName
    
        if (-not $item) {
            New-CWMItem -BoardId $boardId -TypeId $typeId -SubtypeId $subtypeId -ItemName $itemName
        } else {
            Write-Log "    Item already exists: $itemName" -Level "INFO"
        }
    }
    
    Write-Log "" -Level "INFO"
    Write-Log "Creating Item-to-Subtype associations for Board: $boardName..." -Level "INFO"
    Write-Log "" -Level "INFO"
    
        $allItemsUri = "$script:CWMBaseUrl/service/boards/$boardId/items"
    $allItems = @()
    try {
        $allItems = Invoke-RestMethod -Uri $allItemsUri -Headers $script:CWMHeaders -Method Get
    } catch {
        Write-Log "Failed to get items for associations: $_" -Level "ERROR"
    }
    
    if ($allItems.Count -gt 0) {
        foreach ($typeName in $processedTypes.Keys) {
            $type = $processedTypes[$typeName]
            $typeId = $type.id
            
            $subtypeIds = @()
            foreach ($subtypeKey in $processedSubtypes.Keys) {
                if ($subtypeKey.StartsWith("$typeName|")) {
                    $subtypeIds += $processedSubtypes[$subtypeKey].id
                }
            }
            
            if ($subtypeIds.Count -gt 0) {
                foreach ($item in $allItems) {
                    $assocUri = "$script:CWMBaseUrl/service/boards/$boardId/items/$($item.id)/associations/$typeId"
                    $assocBody = @{
                        id = $typeId
                        subTypeAssociationIds = $subtypeIds
                    } | ConvertTo-Json
                    
                    try {
                        Invoke-RestMethod -Uri $assocUri -Headers $script:CWMHeaders -Method Put -Body $assocBody | Out-Null
                    } catch {
                        Write-Log "Failed to create association for Item '$($item.name)' Type '$typeName': $($_.Exception.Message)" -Level "WARNING"
                    }
                }
                Write-Log "  Associated $($subtypeIds.Count) subtypes for Type: $typeName (applied to all items)" -Level "SUCCESS"
            }
        }
    }
    
    Write-Log "" -Level "INFO"
    Write-Log "Board '$boardName' processing complete!" -Level "SUCCESS"
    Write-Log "" -Level "INFO"
}

Write-Log "=========================================" -Level "INFO"
Write-Log "All boards processed successfully!" -Level "SUCCESS"
Write-Log "" -Level "INFO"
Write-Log "SUMMARY:" -Level "INFO"
Write-Log "  Total Successes: $script:SuccessCount" -Level "INFO"
Write-Log "  Total Warnings:  $script:WarningCount" -Level "INFO"
Write-Log "  Total Errors:    $script:ErrorCount" -Level "INFO"
Write-Log "" -Level "INFO"
Write-Log "All Types, Subtypes, Items, and associations have been created." -Level "INFO"
Write-Log "Type/Subtype/Item combinations are now available for ticket creation." -Level "INFO"
Write-Log "=========================================" -Level "INFO"
Write-Log "" -Level "INFO"
Write-Log "Log file saved to: $LogPath" -Level "SUCCESS"