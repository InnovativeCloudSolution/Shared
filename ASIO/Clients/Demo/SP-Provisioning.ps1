# Import required Microsoft Graph modules
$RequiredModules = @("Microsoft.Graph.Sites", "Microsoft.Graph.Files", "Microsoft.Graph.Groups")

foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -Name $Module -ListAvailable)) {
        Import-Module $Module -ErrorAction Stop
    }
}

# Connect to Microsoft Graph with required permissions
Connect-MgGraph -Scopes "Sites.Manage.All", "Sites.ReadWrite.All", "Files.ReadWrite.All", "Group.ReadWrite.All", "Directory.ReadWrite.All"

# Define constants
$TenantDomain = "M365x24689869.onmicrosoft.com"
$SiteUrl = "https://m365x24689869.sharepoint.com/sites/Project"
$LibraryName = "Documents"

$Projects = @(
    @{ Code="RD102"; Name="HighwayExpansion"; Year="2025" },
    @{ Code="WS201"; Name="WaterPipelineUpgrade"; Year="2025" },
    @{ Code="BR512"; Name="SuspensionBridgeUpgrade"; Year="2026" },
    @{ Code="RD203"; Name="UrbanRoadReconstruction"; Year="2024" },
    @{ Code="WS312"; Name="StormwaterManagement"; Year="2024" },
    @{ Code="BLD150"; Name="SportsComplexConstruction"; Year="2025" },
    @{ Code="GT505"; Name="SoilStabilizationStudy"; Year="2025" },
    @{ Code="RW801"; Name="HighSpeedRailLink"; Year="2025" },
    @{ Code="EN301"; Name="SolarFarmInfrastructure"; Year="2025" },
    @{ Code="BR620"; Name="FootbridgeConstruction"; Year="2024" },
    @{ Code="WS418"; Name="WastewaterTreatmentPlant"; Year="2026" },
    @{ Code="BLD275"; Name="SchoolBuildingRenovation"; Year="2024" },
    @{ Code="GT612"; Name="ErosionControlInitiative"; Year="2024" },
    @{ Code="RW915"; Name="MetroLineExtension"; Year="2024" },
    @{ Code="EN405"; Name="WindTurbineFoundation"; Year="2024" },
    @{ Code="BR725"; Name="TollBridgeConstruction"; Year="2026" },
    @{ Code="RD315"; Name="InterchangeUpgrade"; Year="2026" },
    @{ Code="WS520"; Name="ReservoirRehabilitation"; Year="2025" },
    @{ Code="BLD390"; Name="GovernmentOfficeTower"; Year="2026" }
)

$M365RoleGroups = @(
    "SG.Role.ProjectManagers",
    "SG.Role.Engineers",
    "SG.Role.SiteSupervisors",
    "SG.Role.Finance",
    "SG.Role.HealthSafety",
    "SG.Role.Consultants"
)

$AccessTypes = @{
    "FC" = "FullControl"
    "ED" = "Edit"
    "RO" = "Read"
}

$Folders = @("Admin", "Design", "Construct", "Safety", "Finance", "Archive")

Function Get-SiteId {
    param ($SiteUrl)
    $SiteName = ($SiteUrl -split '/')[-1]
    $Site = Get-MgSite -Search $SiteName | Where-Object { $_.WebUrl -eq $SiteUrl }
    if ($Site -and $Site.Id) { return $Site.Id }
    Write-Host "Error: Could not retrieve Site ID for $SiteUrl" -ForegroundColor Red
    return $null
}

Function Get-DriveId {
    param ($SiteId, $LibraryName)
    $Drive = Get-MgSiteDrive -SiteId $SiteId | Where-Object { $_.Name -eq $LibraryName }
    if ($Drive -and $Drive.Id) { return $Drive.Id }
    Write-Host "Error: Could not retrieve Drive ID for '$LibraryName'" -ForegroundColor Red
    return $null
}

Function New-M365Group {
    param ($GroupName)
    $ExistingGroup = Get-MgGroup -Filter "DisplayName eq '$GroupName'"
    if (-not $ExistingGroup) {
        New-MgGroup -DisplayName $GroupName -MailEnabled:$false -SecurityEnabled:$true -MailNickname $GroupName
        Write-Host "Created M365 Security Group: $GroupName"
    } else {
        Write-Host "M365 Security Group already exists: $GroupName"
    }
}

Function New-SPFolder {
    param (
        [string]$DriveId,
        [string]$FolderPath
    )

    # Validate DriveId
    if (-not $DriveId) {
        Write-Host "Error: DriveId is null. Unable to create folders." -ForegroundColor Red
        return
    }

    # Initialize the base path from root
    $CurrentPath = ""

    # Split folder path into segments for nested folder creation
    $PathSegments = $FolderPath -split "/"

    foreach ($Segment in $PathSegments) {
        # Construct full folder path dynamically
        $CurrentPath = if ($CurrentPath -eq "") { $Segment } else { "$CurrentPath/$Segment" }

        # Check if the folder already exists using Graph API
        $CheckUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$($CurrentPath)"
        Try {
            $ExistingFolder = Invoke-MgGraphRequest -Method GET -Uri $CheckUri -Headers @{ "Content-Type" = "application/json" }

            if ($ExistingFolder) {
                Write-Host "Folder exists: $CurrentPath"
                Continue
            }
        } Catch {
            Write-Host "Folder '$CurrentPath' does not exist. Creating it now..." -ForegroundColor Yellow

            # Define folder creation parameters
            $Body = @{
                "name" = $Segment
                "folder" = @{}
                "@microsoft.graph.conflictBehavior" = "rename"
            } | ConvertTo-Json -Depth 2

            # Construct URI for folder creation
            $CreateUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$($CurrentPath):/children"

            # Create the folder
            Try {
                $NewFolder = Invoke-MgGraphRequest -Method POST -Uri $CreateUri -Body $Body -Headers @{ "Content-Type" = "application/json" }
                Write-Host "Created Folder: $CurrentPath"
            } Catch {
                Write-Host "Error creating folder '$CurrentPath': $_" -ForegroundColor Red
                return
            }
        }
    }
}

Function New-SPGroup {
    param (
        [string]$DriveId,     # ID of the SharePoint document library
        [string]$GroupName,   # Name of the group to create or use
        [string]$AccessType,  # Access type (e.g., "RO", "ED", "FC")
        [string]$FolderPath   # Folder path where permissions should be assigned
    )

    # Check if the group already exists
    $ExistingGroup = Get-MgGroup -Filter "DisplayName eq '$GroupName'" -ErrorAction SilentlyContinue
    if (-not $ExistingGroup) {
        $NewGroup = New-MgGroup -DisplayName $GroupName -MailEnabled:$false -SecurityEnabled:$true -MailNickname $GroupName -ErrorAction Stop
        $GroupId = $NewGroup.Id
        Write-Host "Created SharePoint Group: $GroupName"
    } else {
        $GroupId = $ExistingGroup.Id
        Write-Host "SharePoint Group already exists: $GroupName"
    }

    # Map the access type to a valid role
    $Role = switch ($AccessType) {
        "RO" { "read" }
        "ED" { "write" }
        "FC" { "fullControl" }
        default {
            Write-Host "Invalid Access Type: $AccessType" -ForegroundColor Red
            return
        }
    }

    # Retrieve folder ID using Invoke-MgGraphRequest
    Try {
        $Folder = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$FolderPath" -ErrorAction Stop
        $FolderId = $Folder.id
    } Catch {
        Write-Host "Error retrieving folder ID for '$FolderPath': $_" -ForegroundColor Red
        return
    }

    # Assign permissions using direct API call
    $PermissionBody = @{
        roles = @($Role)
        grantedToIdentitiesV2 = @(
            @{
                "@odata.type" = "#microsoft.graph.identitySet"
                group = @{
                    id = $GroupId
                }
            }
        )
    } | ConvertTo-Json -Depth 3

    $PermissionUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$FolderId/permissions"

    Try {
        Invoke-MgGraphRequest -Method POST -Uri $PermissionUri -Body $PermissionBody -Headers @{ "Content-Type" = "application/json" }
        Write-Host "Assigned '$Role' access to: $GroupName on folder '$FolderPath'"
    } Catch {
        Write-Host "Error assigning permissions for '$GroupName' on folder '$FolderPath': $_" -ForegroundColor Red
    }
}

Write-Host "Starting SharePoint & M365 Group Setup..."

# Create M365 Groups
foreach ($M365Group in $M365RoleGroups) {
    New-M365Group -GroupName $M365Group
}

# Retrieve SharePoint Site ID
$SiteId = Get-SiteId -SiteUrl $SiteUrl
if ($null -eq $SiteId) { exit }

# Retrieve Document Library Drive ID
$DriveId = Get-DriveId -SiteId $SiteId -LibraryName $LibraryName
if ($null -eq $DriveId) { exit }

# Loop through each project
foreach ($Project in $Projects) {
    $ProjectCode = $Project["Code"]
    $ProjectName = $Project["Name"]
    $Year = $Project["Year"]

    Write-Host "Processing Project: $ProjectCode - $ProjectName - $Year"

    # Define project folder path
    $ProjectFolder = "$ProjectCode-$ProjectName-$Year"

    # Automatically handle folder creation with built-in checker
    New-SPFolder -DriveId $DriveId -FolderPath $ProjectFolder

    # Assign project folder-level permissions (FC, ED, RO)
    foreach ($AccessType in $AccessTypes.Keys) {
        $SPGroupName = "SG.SP.$ProjectCode.$AccessType"
        New-SPGroup -DriveId $DriveId -GroupName $SPGroupName -AccessType $AccessType -FolderPath $ProjectFolder
    }

    # Create Subfolders inside the project folder and assign permissions
    foreach ($Folder in $Folders) {
        $ProjectSubFolder = "$ProjectFolder/$Folder"
        New-SPFolder -DriveId $DriveId -FolderPath $ProjectSubFolder

        # Assign Access Types to each SharePoint Group for subfolders
        foreach ($AccessType in $AccessTypes.Keys) {
            $SPGroupName = "SG.SP.$ProjectCode.$Folder.$AccessType"
            New-SPGroup -DriveId $DriveId -GroupName $SPGroupName -AccessType $AccessType -FolderPath $ProjectSubFolder
        }
    }
}

Write-Host "SharePoint Structure Setup Completed!"
