param(
    # General variables
    [string]$City,
    [string]$Site,
    [string]$Company,
    [string]$ElevateToM365
)

## ESTABLISH VARIABLES ##

[string]$SecurityGroups = ''
[string]$Licenses = ''

# Define list of companies that get certain licenses
$M365E5Companies = @('CleanCo', 'BAPL', 'CAPGEMINI', 'Centum Services', 'Cosol', 'Ernest & Young', 'Tally Group', 'UGL Limited')
$ProjectPlan3Companies = @('Sensei')
$M365AppsForEnterpriseCompanies = @('PwC')

# Define common license types
$License_M365E5 = 'Microsoft 365 E5:Telstra:null:SPE_E5:Monthly'
$License_EMSE5 = 'Enterprise Mobility and Security E5:Telstra:null:EMSPREMIUM:Monthly'
$License_ProjectPlan3 = 'Project Plan 3:Telstra:null:PROJECTPROFESSIONAL:Monthly'
$License_M365AppsForEnterprise = 'Microsoft 365 Apps for Enterprise:Telstra:null:OFFICESUBSCRIPTION:Monthly'

## LICENSES ##

if ($M365E5Companies.Contains($Company) -or $ElevateToM365 -eq 'Increase license to Microsoft 365 E5') {
    $SecurityGroups += "SG.License.Microsoft365.E5"
    $Licenses += "$License_M365E5;"
} else {
    $SecurityGroups += "SG.License.Microsoft365.EMS_E5"
    $Licenses += "$License_EMSE5;"
}

if ($ProjectPlan3Companies.Contains($Company)) {
    $SecurityGroups += "SG.License.Microsoft365.ProjectPlan3"
    $Licenses += "$License_ProjectPlan3;"
}

if ($M365AppsForEnterpriseCompanies.Contains($Company)) {
    $SecurityGroups += "SG.License.Microsoft365.AppsForEnterprise"
    $Licenses += "$License_M365AppsForEnterprise;"
}

## SEND DATA TO FLOW ##

$Output = @{
    SecurityGroups = $SecurityGroups
    Licenses = $Licenses
}

Write-Output $Output | ConvertTo-Json