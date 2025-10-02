function Get-MSGraph-SPSite {
    param (
        [string]$SiteUrl,
        [string]$SiteName
    )

    $Site = Get-MgSite -Search $SiteName | Where-Object { $_.WebUrl -eq $SiteUrl }
    return $Site
}

function Get-MSGraph-SPList {
    param (
        $Site,
        [string]$ListName
    )

    $List = Get-MgSiteList -SiteId $Site.Id -Filter "DisplayName eq '$ListName'"
    return $List
}

function Get-MSGraph-SPFilteredList {
    param (
        $Site,
        $List
    )

    $ListItems = Get-MgSiteListItem -SiteId $Site.Id -ListId $List.Id -Select 'fields' -ExpandProperty 'fields'

    $FieldsToExclude = @(
        "@odata.etag", "ContentType", "Modified", "Created", "AuthorLookupId", 
        "EditorLookupId", "_UIVersionString", "Attachments", "Edit", 
        "ItemChildCount", "FolderChildCount", "_ComplianceFlags", 
        "_ComplianceTag", "_ComplianceTagWrittenTime", "_ComplianceTagUserId", 
        "AppAuthorLookupId", "AppEditorLookupId", "LinkTitle", "LinkTitleNoMenu",
        "Title"
    )

    $FilteredItems = $ListItems | ForEach-Object {
        $filteredProperties = @{}
        foreach ($key in $_.Fields.AdditionalProperties.Keys) {
            if ($key -notin $FieldsToExclude) {
                $value = $_.Fields.AdditionalProperties[$key]
                if ($value -is [System.Array]) {
                    $filteredProperties[$key] = $value
                }
                else {
                    $filteredProperties[$key] = @($value)
                }
            }
        }
        New-Object PSObject -Property $filteredProperties
    }

    return $FilteredItems
}

function Get-MSGraph-SPItem {
    param (
        [string]$SiteUrl,
        [string]$SiteName,
        [string]$FolderPath,
        [string]$FileName
    )
    try {
        $site = Get-MSGraph-SPSite -SiteUrl $SiteUrl -SiteName $SiteName
        $drives = Get-MgSiteDrive -SiteId $site.Id

        $drive = $drives | Where-Object { $_.Name -eq "Documents" }
        if (-not $drive) {
            throw "Documents library not found in the specified SharePoint site."
        }

        $ItemPath = if ($FolderPath) { "$FolderPath/$FileName" } else { $FileName }

        $driveItem = Get-MgDriveItem -DriveId $drive.Id -Filter "name eq '$FileName'" | Where-Object { $_.Name -eq $FileName }

        if (-not $driveItem) {
            throw "File '$FileName' not found in folder '$FolderPath'."
        }

        $downloadUrl = $driveItem.AdditionalProperties.'@microsoft.graph.downloadUrl'
        $response = Invoke-RestMethod -Uri $downloadUrl -Method Get

        return $response
    }
    catch {
        throw "Failed to retrieve file from SharePoint: $_"
    }
}

function Get-MSGraph-SPPersonaData {
    param (
        [array]$CSVFromSharePoint,
        [string]$Persona
    )
    try {
        $filteredData = $CSVFromSharePoint | Where-Object { $_.Persona -eq $Persona }
        if (-not $filteredData) {
            throw "No rows found for Persona='$Persona'"
        }

        function Join-WithSemicolon {
            param (
                [array]$Values
            )
            if ($Values -and $Values.Count -gt 0) {
                return ($Values -join ";") + ";"
            }
            else {
                return ""
            }
        }

        $groupedData = [PSCustomObject]@{
            Persona           = $FilteredData[0].persona
            LicenseGroups     = ($FilteredData.licensegroup | Where-Object { $_ } | Select-Object -Unique) -join ","
            LicenseSKU        = ($FilteredData.licensesku | Where-Object { $_ } | Select-Object -Unique) -join ","
            M365Groups        = ($FilteredData.m365groupaccess | Where-Object { $_ } | Select-Object -Unique) -join ","
            Sharepoint        = ($FilteredData.'sharepointaccess-1' + $FilteredData.'sharepointaccess-2' | Where-Object { $_ } | Select-Object -Unique) -join ","
            DistributionLists = ($FilteredData.'distributionlistaccess-1' + $FilteredData.'distributionlistaccess-2' | Where-Object { $_ } | Select-Object -Unique) -join ","
            CalendarAccess    = ($FilteredData.calendaraccess | Where-Object { $_ } | Select-Object -Unique) -join ","
            SharedMailbox     = ($FilteredData.sharedmailboxaccess | Where-Object { $_ } | Select-Object -Unique) -join ","
            AADGroups         = ($FilteredData.aadgroupaccess | Where-Object { $_ } | Select-Object -Unique) -join ","
            ADGroups          = ($FilteredData.adgroupaccess | Where-Object { $_ } | Select-Object -Unique) -join ","
            AADSoftwareGroups = ($FilteredData.aadsoftwaregroupaccess | Where-Object { $_ } | Select-Object -Unique) -join ","
            ADSoftwareGroups  = ($FilteredData.adsoftwaregroupaccess | Where-Object { $_ } | Select-Object -Unique) -join ","
            OU                = ($FilteredData.ou | Where-Object { $_ } | Select-Object -Unique) -join ","
            HomeDrive         = ($FilteredData.homedrive | Where-Object { $_ } | Select-Object -Unique) -join ","
            HomeDriveLetter   = ($FilteredData.homedriveletter | Where-Object { $_ } | Select-Object -Unique) -join ","
        }

        return $groupedData
    }
    catch {
        throw "Failed to process Persona data: $_"
    }
}

function Get-MSGraph-SPClientData {
    param (
        [array]$CSVFromSharePoint,
        [string]$Client
    )
    try {
        $filteredData = $CSVFromSharePoint | Where-Object { $_.client -eq $Client }
        if (-not $filteredData) {
            throw "No rows found for Client='$Client'"
        }

        function Join-WithSemicolon {
            param (
                [array]$Values
            )
            if ($Values -and $Values.Count -gt 0) {
                return ($Values -join ";") + ";"
            }
            else {
                return ""
            }
        }

        $groupedData = [PSCustomObject]@{
            Client    = $FilteredData[0].client
            UNformat1 = ($FilteredData.unformat1 | Where-Object { $_ } | Select-Object -Unique) -join ","
            UNformat2 = ($FilteredData.unformat2 | Where-Object { $_ } | Select-Object -Unique) -join ","
            UNformat3 = ($FilteredData.unformat3 | Where-Object { $_ } | Select-Object -Unique) -join ","
            UPNformat = ($FilteredData.upnformat | Where-Object { $_ } | Select-Object -Unique) -join ","
            DNformat  = ($FilteredData.dnformat | Where-Object { $_ } | Select-Object -Unique) -join ","
            MNformat  = ($FilteredData.mnformat | Where-Object { $_ } | Select-Object -Unique) -join ","
        }

        return $groupedData
    }
    catch {
        throw "Failed to process Client data: $_"
    }
}

function Get-MSGraph-SPManualData {
    param (
        [array]$CSVFromSharePoint
    )
    try {
        if (-not $CSVFromSharePoint) {
            throw "CSV data is empty or not provided."
        }

        $manualData = $CSVFromSharePoint.manual | Where-Object { $_ }

        return ,$manualData
    }
    catch {
        throw "Failed to process manual data: $_"
    }
}

function Save-MSGraph-SPJson {
    param (
        [string]$SiteUrl,
        [string]$SiteName,
        [string]$FolderPath,
        [string]$FileName,
        [PSCustomObject]$JsonData
    )
    try {
        $Site = Get-MSGraph-SPSite -SiteUrl $SiteUrl -SiteName $SiteName
        $Drives = Get-MgSiteDrive -SiteId $Site.Id
        $Drive = $Drives | Where-Object { $_.Name -eq "Documents" }

        if (-not $Drive) {
            throw "Documents library not found in the specified SharePoint site."
        }

        $FolderItem = Get-MgDriveItem -DriveId $Drive.Id -Filter "Name eq '$FolderPath'" | Where-Object { $_.Name -eq $FolderPath -and $_.Folder }
 
        if (-not $FolderItem) {
            throw "Folder '$FolderPath' not found in SharePoint Documents."
        }

        $JsonContent = $JsonData | ConvertTo-Json -Depth 10 -Compress
        $JsonBytes = [System.Text.Encoding]::UTF8.GetBytes($JsonContent)
        
        Invoke-MgGraphRequest -Method Put -Uri "https://graph.microsoft.com/v1.0/drives/$($Drive.Id)/root:/$($FolderPath)/$($FileName):/content" -Body $JsonBytes -Headers @{ "Content-Type" = "application/json" }
    }
    catch {
        throw "Failed to save JSON to SharePoint: $_"
    }
}
