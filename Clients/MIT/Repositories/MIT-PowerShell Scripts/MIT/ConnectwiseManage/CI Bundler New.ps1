# Generic parts
function Get-CommonHeaders {
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("clientId", "1208536d-40b8-4fc0-8bf3-b4955dd9d3b7")
    $headers.Add("Authorization", "Basic bWl0K2Q2anR1V01VOVkzQzFKdlk6TjExczRNZzkwUVZCUG1XYg==")
    $headers.Add("Content-Type", "application/json")
    return $headers
}

# Function to remove parent configuration
function Remove-ParentConfiguration {
    param (
        [string]$CompanyName,
        [string]$ConfigName,
        [string]$TypeName,
        [hashtable]$headers
    )

    Write-Host "Removing Parent Config from all $TypeName configurations" 
    $Configurations = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/?pageSize=1000&conditions=company/name='$CompanyName' %26%26 type/name='$ConfigName'" -Method 'GET' -Headers $headers

    $nullParentConfigurationBody = "[
    `n  {
    `n    `"op`": `"replace`",
    `n    `"path`": `"parentConfigurationId`",
    `n    `"value`": null
    `n  }
    `n]"

    foreach ($Configuration in $Configurations) {
        $ConfigId = $Configuration.id
        Write-Host "Processing Config ID: $ConfigId"
        $response = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/$ConfigId" -Method 'PATCH' -Headers $headers -Body $nullParentConfigurationBody
    }
    Write-Host "Removed Parent Config from all $TypeName configurations" 
}

# Function to bundle child configurations
function Bundle-ChildConfigurations {
    param (
        [string]$CompanyName,
        [string]$PCNBConfigName,
        [string]$MWConfigName,
        [hashtable]$headers
    )

    $MWConfigurations = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/?pageSize=1000&conditions=company/name='$CompanyName' %26%26 type/name='$MWConfigName' %26%26 parentConfigurationId==null" -Method 'GET' -Headers $headers
    foreach ($MW in $MWConfigurations) {
        $MWID = $MW.id
        $ConfigName = $MW.Name
        $serialNumber = $MW.serialNumber

        if ($serialNumber -ne "") {
            $ApplicablePCNBs = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/?pageSize=1000&conditions=company/name='$CompanyName' %26%26 type/name='$PCNBConfigName' %26%26 serialNumber like '*$serialNumber*'" -Method 'GET' -Headers $headers
            if ($ApplicablePCNBs.Count -ge 1) {
                $PCNBConfigId = if ($ApplicablePCNBs.Count -eq 1) { $ApplicablePCNBs.id } else { $ApplicablePCNBs[0].id }
                
                $parentConfigurationBody = "[
                `n  {
                `n    `"op`": `"replace`",
                `n    `"path`": `"parentConfigurationId`",
                `n    `"value`": $PCNBConfigId
                `n  }
                `n]"

                Write-Host "Found applicable parent PCNB for $ConfigName $serialNumber, ParentId = $PCNBConfigId, ChildId = $MWID" -ForegroundColor DarkYellow
                $response = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/$MWID" -Method 'PATCH' -Headers $headers -Body $parentConfigurationBody
            } else {
                $ResponseConfigId = New-PCNBConfiguration -ConfigName $ConfigName -serialNumber $serialNumber -MW $MW -headers $headers
                Bundle-MWIntoPCNB -MWID $MWID -ResponseConfigId $ResponseConfigId -headers $headers
            }
        }
        Write-Host "Done, next!" -ForegroundColor Gray
    }
}

# Function to create new PCNB configuration
function New-PCNBConfiguration {
    param (
        [string]$ConfigName,
        [string]$serialNumber,
        [object]$MW,
        [hashtable]$headers
    )

    $ConfigBody = "{
    `n    `"name`": `"$ConfigName`",
    `n    `"status`": {
    `n        `"id`": $($MW.status.id)
    `n    },
    `n    `"company`": {
    `n        `"id`": $($MW.company.id)
    `n    },
    `n    `"type`": {
    `n        `"id`": 25
    `n    },
    `n    `"serialNumber`": `"$serialNumber`",
    `n    `"modelNumber`": `"$($MW.modelNumber)`",
    `n    `"tagNumber`": `"$($MW.tagNumber)`",
    `n    `"name`": `"$ConfigName`"
    `n}"

    Write-Host "Creating new PCNB for $ConfigName $serialNumber..." -ForegroundColor Yellow
    $response = Invoke-RestMethod 'https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations' -Method 'POST' -Headers $headers -Body $ConfigBody
    return $response.id
}

# Function to bundle MW into the newly created PCNB
function Bundle-MWIntoPCNB {
    param (
        [string]$MWID,
        [string]$ResponseConfigId,
        [hashtable]$headers
    )

    $parentConfigurationBody = "[
    `n  {
    `n    `"op`": `"replace`",
    `n    `"path`": `"parentConfigurationId`",
    `n    `"value`": $ResponseConfigId
    `n  }
    `n]"

    Write-Host "Bundling MW ID $MWID into PCNB ID $ResponseConfigId" -ForegroundColor Green
    $response = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/$MWID" -Method 'PATCH' -Headers $headers -Body $parentConfigurationBody
}

# Main script
$CompanyName = "Kern Group"
$PCNBConfigName = "PC/Notebook"
$MWConfigName = "Managed Workstation"
$RedoChildren = $false

$headers = Get-CommonHeaders

if ($RedoChildren) {
    Remove-ParentConfiguration -CompanyName $CompanyName -ConfigName $PCNBConfigName -TypeName "PCNB" -headers $headers
    Remove-ParentConfiguration -CompanyName $CompanyName -ConfigName $MWConfigName -TypeName "MW" -headers $headers
}

Bundle-ChildConfigurations -CompanyName $CompanyName -PCNBConfigName $PCNBConfigName -MWConfigName $MWConfigName -headers $headers
