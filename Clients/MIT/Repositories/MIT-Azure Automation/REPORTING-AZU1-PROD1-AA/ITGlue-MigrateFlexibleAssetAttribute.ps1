<#

Mangano IT - Migrate Flexible Asset Attribute
Created by: Gabriel Nugent
Version: 0.7

This runbook is designed to be used on its own.

Please manually update the script with the fields you need to edit!

#>

param(
    [Parameter(Mandatory)][int]$FlexibleAssetTypeId = 19480,
    [int]$OrganizationId
)

## SCRIPT VARIABLES ##

$ContentType = "application/vnd.api+json"

# Grab API key from vault
$ApiKey = Get-AutomationVariable -Name 'ITGlue-ApiKey'

# URL for IT Glue API
$ApiUrl = "https://api.itglue.com"

## GET ALL ASSETS ##

$GetAssetsArguments = @{
	Uri = "$ApiUrl/flexible_assets"
	Method = 'GET'
	Body = @{ 
		'page[size]' = 10000
		'filter[flexible-asset-type-id]' = $FlexibleAssetTypeId
	}
	ContentType = $ContentType
	Headers = @{ 'x-api-key'=$ApiKey }
	UseBasicParsing = $true
}

# Add organisation filter if provided
if ($OrganizationId -ne 0) {
    $GetAssetsArguments.Body += @{ 'filter[organization-id]' = $OrganizationId }
}

# Fetch all assets
try {
    Write-Warning "Fetching all flexible assets with the ID $FlexibleAssetTypeId..."
    $Assets = Invoke-WebRequest @GetAssetsArguments | ConvertFrom-Json -AsHashtable
    Write-Warning "SUCCESS: Fetched all flexible assets with the ID $FlexibleAssetTypeId."
} catch {
    Write-Error "Unable to fetch assets : $_"
}

## SORT THROUGH ASSETS AND UPDATE ##

# Go through each asset and update field
foreach ($Asset in $Assets.data) {
    if ($Asset.attributes.traits.'business-impact-old' -ne '' -and $null -ne $Asset.attributes.traits.'business-impact-old') {
        $AssetId = $Asset.id
        Write-Warning "INFO: Old field found for $AssetId."

        # Replace old field with new field
        if ($null -ne $Asset.attributes.traits.'business-impact') {
            $Asset.attributes.traits.'business-impact' = $Asset.attributes.traits.'business-impact-old'
        } else {
            $Asset.attributes.traits += @{
                'business-impact' = $Asset.attributes.traits.'business-impact-old'
            }
        }
        $Asset.attributes.traits.Remove('business-impact-old')

        # Prep custom fields with IDs for API push
        if ($null -ne $Asset.attributes.traits.'application-champion'.values.id) {
            $Asset.attributes.traits.'application-champion' = $Asset.attributes.traits.'application-champion'.values.id
        }
        if ($null -ne $Asset.attributes.traits.'application-server'.values.id) {
            $Asset.attributes.traits.'application-server' = $Asset.attributes.traits.'application-server'.values.id
        }
        if ($null -ne $Asset.attributes.traits.vendor.values.id) {
            $Asset.attributes.traits.vendor = $Asset.attributes.traits.vendor.values.id
        }
        if ($null -ne $Asset.attributes.traits.'workstation-installation-guide'.values.id) {
            $Asset.attributes.traits.'workstation-installation-guide' = $Asset.attributes.traits.'workstation-installation-guide'.values.id
        }
        if ($null -ne $Asset.attributes.traits.license.values.id) {
            $Asset.attributes.traits.license = $Asset.attributes.traits.license.values.id
        }

        # Create request body
        $ApiBody = @{
            data = @{ 
                type = "flexible-assets"
                attributes = @{
                    traits = $Asset.attributes.traits
                }
            }
        } | ConvertTo-Json -Depth 100

        Write-Output $ApiBody

        # Setup request
        $ApiArguments = @{
            Uri = "$ApiUrl/flexible_assets/$AssetId"
            Method = 'PATCH'
            Body = $ApiBody
            ContentType = $ContentType
            Headers = @{ 'x-api-key'=$ApiKey }
            UseBasicParsing = $true
        }

        # Update field
        try {
            Write-Warning "Updating field found for $AssetId..."
            Invoke-WebRequest @ApiArguments | Out-Null
            Write-Warning "SUCCESS: Updated field found for $AssetId."
        } catch {
            "Unable to update field for $AssetId : $_"
        }
    }
}