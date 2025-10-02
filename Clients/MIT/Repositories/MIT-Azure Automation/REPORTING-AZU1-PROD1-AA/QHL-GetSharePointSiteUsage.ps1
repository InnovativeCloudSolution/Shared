<#

Queensland Hydro - Get SharePoint Site Usage
Created by: Gabriel Nugent
Version: 1.0.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

## SCRIPT VARIABLES ##

# Connection info
$SiteUrl = "https://queenslandhydro-admin.sharepoint.com"
$ApplicationClientId = 'ead975fb-be9d-4516-aa6a-a355e301e454'
$CertificateThumbprint = '4B5A28D92DA290B0C1D4752C40C5729FCB8C7254'
$Tenant = 'queenslandhydro.onmicrosoft.com'

# Site info to be exported
$Sites = @()

## GET SITE INFO ##

# Build connection info
$Connection = @{
    Url = $SiteUrl
    ClientId = $ApplicationClientId
    Thumbprint = $CertificateThumbprint
    Tenant = $Tenant
}

# Connect to SharePoint Online
try {
    Write-Warning "Connecting to SharePoint Online site '$($SiteUrl)'..."
    Connect-PnPOnline @Connection | Out-Null
    Write-Warning "SUCCESS: Connected to SharePoint Online site '$($SiteUrl)'."
} catch {
    Write-Error "Unable to connect to SharePoint : $($_)"
}

# Get all site collections
$SiteCollections = Get-PnPSite

# Loop through each site collection
foreach ($Site in $SiteCollections) {
    # Get data size (in MB)
    #$dataSize = $Site.StorageUsageCurrent
    
    # Get other data (example: Last Content Modified Date)
    #$lastContentModifiedDate = $Site.LastContentModifiedDate
    
    # Add site info to list of sites
    #$Sites += [ordered]@{
    #    SiteUrl = $Site.Url
    #    DataSizeInMB = $dataSize
    #    LastContentModifiedDate = $lastContentModifiedDate
    #}

    # Add site to sites
    $Sites += $Site
}

## WRITE OUTPUT TO FLOW ##

if ($Sites -ne @()) {
    Write-Output $Sites | ConvertTo-Json -Depth 100
}