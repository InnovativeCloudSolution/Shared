param(
    # General variables
    [string]$City,
    [string]$Site,
    [string]$Company,

    # Security groups
    [string]$SecGroup_SnagIt,
    [string]$SecGroup_ez2view,
    [string]$SecGroup_RProject,
    [string]$SecGroup_Convene,
    [string]$SecGroup_NEMSight,
    [string]$SecGroup_NEMWatch,
    [string]$SecGroup_NEOexpress,
    [string]$SecGroup_NEMReview,
    [string]$SecGroup_MSVisio,
    [string]$SecGroup_MSProject,
    [string]$SecGroup_MSWhiteboard,
    [string]$SecGroup_Eikon,
    [string]$SecGroup_StandardCitrix,
    [string]$SecGroup_TraderCitrix,
	[string]$SecGroup_SapReadOnly,
    [string]$SecGroup_RedEyeDMS,
    [string]$SecGroup_OverseasAccess
)

## ESTABLISH VARIABLES ##

[string]$SecurityGroups = ''
[string]$PublicFolders = ''
[string]$SharedMailboxes = ''
[string]$DistributionGroups = ''
[string]$MailEnabledSecurityGroups = ''
[string]$TeamsChannels = ''
[string]$SharePointGroups = ''
[string]$Licenses = ''

## SECURITY GROUPS ##

# Add template groups
$SecurityGroups += ';SG.App.O365x64;SG.Role.CQL.AllUsers;SG.Role.IntuneUser;SG.Role.PrinterAccess;SG.Role.VpnAccess;SG.Role.NAC.Wireless;SG.Role.NAC.Wired;SG.Role.WorkspaceOneEnrollment;SG.App.Bitwarden;SG.App.AdobeReader'

if ($SecGroup_SnagIt -eq 'True') { $SecurityGroups += ";SG.App.Snagit" }
if ($SecGroup_ez2view -eq 'True') { $SecurityGroups += ";SG.App.ez2view" }
if ($SecGroup_RProject -eq 'True') { $SecurityGroups += ";SG.App.RProject" }
if ($SecGroup_Convene -eq 'True') { $SecurityGroups += ";SG.App.Convene" }
if ($SecGroup_NEMSight -eq 'True') { $SecurityGroups += ";SG.App.NEM-Sight" }
if ($SecGroup_NEMWatch -eq 'True') { $SecurityGroups += ";SG.App.NEM-Watch" }
if ($SecGroup_NEOexpress -eq 'True') { $SecurityGroups += ";SG.App.NEOexpress" }
if ($SecGroup_NEMReview -eq 'True') { $SecurityGroups += ";SG.App.NEM-Review" }
if ($SecGroup_MSVisio -eq 'True') { $SecurityGroups += ";SG.App.O365Visio" }
if ($SecGroup_MSProject -eq 'True') { $SecurityGroups += ";SG.App.O365Project" }
if ($SecGroup_MSWhiteboard -eq 'True') { $SecurityGroups += ";SG.App.Whiteboard" }
if ($SecGroup_Eikon -eq 'True') { $SecurityGroups += ";SG.App.Eikon" }
if ($SecGroup_StandardCitrix -eq 'Yes') { $SecurityGroups += ";SG.Role.CitrixVDIAccess.Standard" }
if ($SecGroup_TraderCitrix -eq 'True') { $SecurityGroups += ";SG.Role.CitrixAccess.Trader" }
if ($SecGroup_SapReadOnly -eq 'Yes') { $SecurityGroups += ";SG.File.SAP.ReadOnly" }
if ($SecGroup_RedEyeDMS -eq 'Yes') { $SecurityGroups += ";SG.App.RedEyeDMS" }
if ($SecGroup_OverseasAccess -eq 'Yes') { $SecurityGroups += ";SG.Policy.AzureCA.BlockOffshore.Exclude" }
if (($City -eq "Wivenhoe Pocket") -or ($City -eq "Swanbank") -or ($City -eq "Kuranda") -or ($City -eq "Cardstone")) { $SecurityGroups += ";SG.Role.TSAOnboardedUsers.SCL" }
switch ($City) {
    "Caravonica" {
        $SecurityGroups += ";SG.Site.AllUsersBarron"
        $SecurityGroups += ";SG.Print.BarrongorgePrinters"
        $SecurityGroups += ";SG.Print.BarrongorgeEpasTagPrinters"
        Break
    }
    "Brisbane" {
        $SecurityGroups += ";SG.Site.AllUsersBrisbane"
        Break
    }
    "Cardstone" {
        $SecurityGroups += ";SG.Site.AllUsersKareeya"
        $SecurityGroups += ";SG.Print.KareeyaPrinters"
        $SecurityGroups += ";SG.Print.KareeyaEpasTagPrinters"
        Break
    }
    "Swanbank" {
        $SecurityGroups += ";SG.Site.AllUsersSwanbank"
        $SecurityGroups += ";SG.Print.SwanbankPrinters"
        $SecurityGroups += ";SG.Print.SwanbankEpasTagPrinters"
        Break
    }
    "Wivenhoe Pocket" {
        $SecurityGroups += ";SG.Site.AllUsersWivenhoe"
        $SecurityGroups += ";SG.Print.WivenhoePrinters"
        $SecurityGroups += ";SG.Print.WivenhoeEpasTagPrinters"
        Break
    }
}

# Overseas access based on company
switch ($Company) {
    'PwC' { $SecurityGroups += ";SG.Role.External.PWC.Offshore;SG.Role.CitrixVDIAccess.Vendor.PwC" }
}

## DISTRIBUTION GROUPS ##

if ($Company -eq "CleanCo") {
    switch -Wildcard ($Site) {
        "BRIS3*" {
            $DistributionGroups += ";BRI.Staff@cleancoqld.com.au"
            break
        }
        "KARY1*" {
            $DistributionGroups += ";KAR.Staff@cleancoqld.com.au"
            break
        }
        "BARG1*" {
            $DistributionGroups += ";BAR.Staff@cleancoqld.com.au"
            break
        }
        "SWNB1*" {
            $DistributionGroups += ";SBK.Staff@cleancoqld.com.au"
            break
        }
        "WIVE1*" {
            $DistributionGroups += ";WIV.Staff@cleancoqld.com.au"
            break
        }
        default {
            $DistributionGroups += ";AllStaff@cleancoqld.com.au"
        }
    }
} else {
    $DistributionGroups += ";BRIExtPartners@cleancoqld.com.au"
}

## SEND DATA TO FLOW ##

$Output = @{
    SecurityGroups = $SecurityGroups
	PublicFolders = $PublicFolders
	SharedMailboxes = $SharedMailboxes
	DistributionGroups = $DistributionGroups
	MailEnabledSecurityGroups = $MailEnabledSecurityGroups
	TeamsChannels = $TeamsChannels
    SharePointGroups = $SharePointGroups
    Licenses = $Licenses
}

Write-Output $Output | ConvertTo-Json