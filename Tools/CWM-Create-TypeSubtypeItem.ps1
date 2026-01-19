param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$true)]
    [string]$BoardName,
    
    [Parameter(Mandatory=$false)]
    [string]$CWMUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$CompanyId,
    
    [Parameter(Mandatory=$false)]
    [string]$PublicKey,
    
    [Parameter(Mandatory=$false)]
    [string]$PrivateKey,
    
    [Parameter(Mandatory=$false)]
    [string]$ClientId
)

function Connect-CWM {
    param(
        [string]$CWMUrl,
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
    
    $script:CWMBaseUrl = $CWMUrl
    
    Write-Output "Connected to ConnectWise Manage at $CWMUrl"
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
            Write-Error "Board '$BoardName' not found"
            return $null
        }
    } catch {
        Write-Error "Failed to retrieve board: $_"
        return $null
    }
}

function Get-CWMType {
    param(
        [int]$BoardId,
        [string]$TypeName
    )
    
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/types?conditions=name='$TypeName'"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            return $null
        }
    } catch {
        Write-Error "Failed to retrieve type: $_"
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
        Write-Output "Created Type: $TypeName"
        return $response
    } catch {
        Write-Error "Failed to create type '$TypeName': $_"
        return $null
    }
}

function Get-CWMSubtype {
    param(
        [int]$BoardId,
        [int]$TypeId,
        [string]$SubtypeName
    )
    
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/subtypes?conditions=name='$SubtypeName' and type/id=$TypeId"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            return $null
        }
    } catch {
        Write-Error "Failed to retrieve subtype: $_"
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
        type = @{
            id = $TypeId
        }
    } | ConvertTo-Json
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Output "  Created Subtype: $SubtypeName"
        return $response
    } catch {
        Write-Error "Failed to create subtype '$SubtypeName': $_"
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
    
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/items?conditions=name='$ItemName' and type/id=$TypeId and subType/id=$SubtypeId"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            return $null
        }
    } catch {
        Write-Error "Failed to retrieve item: $_"
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
    
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/items"
    
    $body = @{
        name = $ItemName
        type = @{
            id = $TypeId
        }
        subType = @{
            id = $SubtypeId
        }
    } | ConvertTo-Json
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Output "    Created Item: $ItemName"
        return $response
    } catch {
        Write-Error "Failed to create item '$ItemName': $_"
        return $null
    }
}

if (-not (Test-Path $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}

Connect-CWM -CWMUrl $CWMUrl -CompanyId $CompanyId -PublicKey $PublicKey -PrivateKey $PrivateKey -ClientId $ClientId

$board = Get-CWMBoard -BoardName $BoardName

if (-not $board) {
    Write-Error "Cannot proceed without valid board"
    exit 1
}

$boardId = $board.id
Write-Output "Using Board: $($board.name) (ID: $boardId)"
Write-Output ""

$csvData = Import-Csv -Path $CsvPath

$processedTypes = @{}
$processedSubtypes = @{}

foreach ($row in $csvData) {
    $typeName = $row.Type.Trim()
    $subtypeName = $row.Subtype.Trim()
    $itemName = $row.Item.Trim()
    
    if ([string]::IsNullOrWhiteSpace($typeName) -or [string]::IsNullOrWhiteSpace($subtypeName) -or [string]::IsNullOrWhiteSpace($itemName)) {
        Write-Warning "Skipping row with missing data: Type='$typeName', Subtype='$subtypeName', Item='$itemName'"
        continue
    }
    
    if (-not $processedTypes.ContainsKey($typeName)) {
        $type = Get-CWMType -BoardId $boardId -TypeName $typeName
        
        if (-not $type) {
            $type = New-CWMType -BoardId $boardId -TypeName $typeName
        } else {
            Write-Output "Type already exists: $typeName"
        }
        
        if ($type) {
            $processedTypes[$typeName] = $type
        } else {
            Write-Error "Failed to get or create type: $typeName"
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
            Write-Output "  Subtype already exists: $subtypeName"
        }
        
        if ($subtype) {
            $processedSubtypes[$subtypeKey] = $subtype
        } else {
            Write-Error "Failed to get or create subtype: $subtypeName"
            continue
        }
    }
    
    $subtypeId = $processedSubtypes[$subtypeKey].id
    
    $item = Get-CWMItem -BoardId $boardId -TypeId $typeId -SubtypeId $subtypeId -ItemName $itemName
    
    if (-not $item) {
        New-CWMItem -BoardId $boardId -TypeId $typeId -SubtypeId $subtypeId -ItemName $itemName
    } else {
        Write-Output "    Item already exists: $itemName"
    }
}

Write-Output ""
Write-Output "Processing complete!"
