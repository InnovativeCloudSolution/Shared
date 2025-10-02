param(
    [string]$Email = '',
    [string]$Team = '',
    [string]$TenantId = '',
    [string]$ScriptAdmin = ''
)

Import-Module MicrosoftTeams
Write-Host "`nConnecting to Microsoft Teams..."
Connect-MicrosoftTeams -TenantId $TenantId -AccountId $ScriptAdmin

# Function to add a user to a Teams channel
function Add-UserToChannel {
    param (
        $Email = '',
        $ChannelName = ''
    )

    try {
        $Channel = Get-Team -DisplayName $ChannelName
        Add-TeamUser -GroupId $Channel.GroupId -User $Email
        Write-Output "$Email has been added to $ChannelName.`r`n"
    }
    catch { Write-Output "ERROR: $Email has not been added to $ChannelName.`r`n" }
}

# Define channels that every user is added to
$StandardChannels = @(
    'Team Mangano'
    'Tech Team'
)

# Define channels for the service delivery team
$ServiceDeliveryChannels = @(
    
)

# Define channels for the L1 members of the service delivery team
$ServiceDeliveryL1Channels = @(
    
)

# Define channels for the L2 members of the service delivery team
$ServiceDeliveryL2Channels = @(
    
)

# Define channels for the L3 members of the service delivery team
$ServiceDeliveryL3Channels = @(
    
)

# Define channels for the sales team
$SalesChannels = @(
    'Marketing Team'
    'Sales Team'
)

# Define channels for the project team
$ProjectsChannels = @(
    'Project Delivery'
)

# Define channels for the leadership team
$LeadershipChannels = @(
    'Internal Systems'
    'Leadership Team'
    'Marketing Team'
    'Project Delivery'
    'Recruitment'
    'Sales Team'
)

# Adds user to standard channels
foreach ($ChannelName in $StandardChannels) { Add-UserToChannel -Email $Email -ChannelName $ChannelName }

# Switch statement for adding user account to channels based on team
switch ($Team) {
    "Service Team Level 1" {
        foreach ($ChannelName in $ServiceDeliveryChannels) { Add-UserToChannel -Email $Email -ChannelName $ChannelName }
        foreach ($ChannelName in $ServiceDeliveryL1Channels) { Add-UserToChannel -Email $Email -ChannelName $ChannelName }
        break
    }
    "Service Team Level 2" {
        foreach ($ChannelName in $ServiceDeliveryChannels) { Add-UserToChannel -Email $Email -ChannelName $ChannelName }
        foreach ($ChannelName in $ServiceDeliveryL2Channels) { Add-UserToChannel -Email $Email -ChannelName $ChannelName }
        break
    }
    "Service Team Level 3" {
        foreach ($ChannelName in $ServiceDeliveryChannels) { Add-UserToChannel -Email $Email -ChannelName $ChannelName }
        foreach ($ChannelName in $ServiceDeliveryL3Channels) { Add-UserToChannel -Email $Email -ChannelName $ChannelName }
        break
    }
    "Projects Team" {
        foreach ($ChannelName in $ProjectsChannels) { Add-UserToChannel -Email $Email -ChannelName $ChannelName }
        break
    }
    "Sales Team" {
        foreach ($ChannelName in $SalesChannels) { Add-UserToChannel -Email $Email -ChannelName $ChannelName }
        break
    }
    "Leadership Team" {
        foreach ($ChannelName in $LeadershipChannels) { Add-UserToChannel -Email $Email -ChannelName $ChannelName }
        break
    }
}