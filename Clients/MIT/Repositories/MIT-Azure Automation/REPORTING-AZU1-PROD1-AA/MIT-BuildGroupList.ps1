<#

Mangano IT - Export List of Required Groups for Internal New Users
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory)][string]$Title,
    [Parameter(Mandatory)][string]$Team
)

## ESTABLISH VARIABLES ##

[string]$DistributionGroups = ''
[string]$M365Groups = ''
[string]$SharedMailboxes = ''
[string]$Licenses = ''

## BASELINE GROUPS AND MAILBOXES ##

# Add minimum groups
$M365Groups += "SG.Role.SSO.CWManage;SG.LastPass.SynchronizedUsers;SG.Device.Windows.WiFi;SG.Role.AllStaff;SG.Role.PaloAlto.VPN"

# Add temporary groups for initial setup
$M365Groups += ";SG.Policy.AzureCA.DisableMFA;SG.Policy.AzureCA.BlockNonManganoIP"

 # Add minimum licenses
$M365Groups += ";SG.License.M365E5.StandardUser"

# Add minimum Teams
$M365Groups += ";Team Mangano;Tech Team;Information Hub"

# Add minimum DLs
$DistributionGroups += "AllTeam@manganoit.com.au;emailsignaturegroup@manganoit.com.au"

# Add minimum shared mailboxes
$SharedMailboxes += "sms@manganoit.com.au:FullAccess"

## ADDITIONAL GROUPS AND MAILBOXES BASED ON TEAM ##

switch ($Team) {
    "Automation & Internal Systems" {
        # AAD groups
        $M365Groups += ";SG.App.AzureCLI;SG.App.Bullphish;SG.Device.Windows.EnableUSBData;SG.LastPass.InternalSystems;SG.Officevibe.InternalSystems"
        $M365Groups += ";SG.Role.DevOps.Administrators;SG.Role.DevOps.Users.Subscriber;SG.Role.InternalSystems;SG.Role.PilotUsers;SG.Role.SSO.CWSell"
        $M365Groups += ";SG.Role.SSO.Pia;SG.Role.SSO.ThreatLocker"

        # Groups to deploy apps
        $M365Groups += ";SG.App.YubicoAuthenticator"

        # Licenses
        #$M365Groups += ";SG.License.PowerBIPro"
        #$Licenses += ";Power BI Pro:Microsoft:null:POWER_BI_PRO:null"

        # Teams
        $M365Groups += ";Internal Systems;Power Automate"

        # Shared mailboxes
        $SharedMailboxes += ";cybersecurity@manganoit.com.au:FullAccess;alerts@manganoit.com.au:FullAccess;automate@manganoit.com.au:FullAccess"
        $SharedMailboxes += ";backup.servers@manganoit.com.au:FullAccess;Hosting@manganoit.com.au:FullAccess;internal@manganoit.com.au:FullAccess"
        $SharedMailboxes += ";InternalAdmins@manganoit.com.au:FullAccess;support@manganoit.com.au:FullAccess;voicemail@manganoit.com.au:FullAccess"

        # Distribution groups
        $DistributionGroups += ";AllTechs@manganoit.com.au;manganoadmins@manganoit.com.au"
    }
    "Azure & Modern Work Practice" {
        switch ($Title) {
            "Technical Consultant" {
                # AAD groups
                $M365Groups += ";"

                # Groups to deploy apps
                $M365Groups += ";"

                # Licenses
                $M365Groups += ";"

                # Teams
                $M365Groups += ";"

                # Shared mailboxes
                $SharedMailboxes += ";"

                # Distribution groups
                $DistributionGroups += ";"
            }
            "Senior Technical Consultant" {
                # AAD groups
                $M365Groups += ";"

                # Groups to deploy apps
                $M365Groups += ";"

                # Licenses
                $M365Groups += ";"

                # Teams
                $M365Groups += ";"

                # Shared mailboxes
                $SharedMailboxes += ";"

                # Distribution groups
                $DistributionGroups += ";"
            }
            Default {
                # AAD groups
                $M365Groups += ";SG.LastPass.Projects;SG.Role.AutomationAccount.AutomationOperator;SG.Role.Projects;SG.Officevibe.Projects"

                # Groups to deploy apps
                $M365Groups += ";SG.App.Microsoft365.Project;SG.App.Microsoft365.Visio"

                # Licenses
                $M365Groups += ";SG.License.PowerApp;SG.License.Visio.Plan2;SG.License.Project.Plan5"
                $Licenses += ";Visio Plan 2:Microsoft:null:VISIOCLIENT:null"
                $Licenses += ";Project Plan 5:Microsoft:null:PROJECTPREMIUM:null"

                # Teams
                $M365Groups += ";Project Delivery"

                # Shared mailboxes
                $SharedMailboxes += ";projectupdates@manganoit.com.au:FullAccess;"

                # Distribution groups
                $DistributionGroups += ";AllTechs@manganoit.com.au"
            }
        }
    }
    "Corporate Services" {
        # AAD groups
        $M365Groups += ";SG.App.Bullphish;SG.Role.Sales;SG.Role.SSO.CWSell"

        # Groups to deploy apps
        #$M365Groups += ";SG.App.PowerBI"

        # Licenses
        #$M365Groups += ";SG.License.PowerBIPro"
        #$Licenses += ";Power BI Pro:Microsoft:null:POWER_BI_PRO:null"

        # Teams
        $M365Groups += ";Accounts Team;Human Resources Team;IBA - Incredible Business Academy;Marketing Team;Recruitment;Sales Team"

        # Shared mailboxes
        $SharedMailboxes += ";"

        # Distribution groups
        $DistributionGroups += ";"
    }
    "Managed Services Practice" {
        switch ($Title) {
            "Service Desk Analyst" {
                # AAD groups
                $M365Groups += ";SG.Role.ServiceDesk.L1;SG.Officevibe.Service.L1"
            }
            "Technical Consultant" {
                # AAD groups
                $M365Groups += ";SG.Device.Windows.EnableUSBData;SG.Role.ServiceDesk.L2;SG.Officevibe.Service.L2"
            }
            "Technical Account Manager" {
                # AAD groups
                $M365Groups += ";SG.Device.Windows.EnableUSBData;SG.Role.ServiceDesk.L3;SG.Officevibe.Service.L3"

                # Groups to deploy apps
                $M365Groups += ";SG.App.Microsoft365.Project;SG.App.Microsoft365.Visio"

                # Licenses
                $M365Groups += ";SG.License.PowerApp;SG.License.Visio.Plan2;SG.License.Project.Plan5"
                $Licenses += ";Visio Plan 2:Microsoft:null:VISIOCLIENT:null"
                $Licenses += ";Project Plan 5:Microsoft:null:PROJECTPREMIUM:null"

                # Shared mailboxes
                $SharedMailboxes += ";sdm@manganoit.com.au:FullAccess"
            }
            Default {
                # AAD groups
                $M365Groups += ";SG.LastPass.ServiceDesk"

                # Distribution groups
                $DistributionGroups += ";AllTechs@manganoit.com.au;ServiceDelivery@manganoit.com.au"
            }
        }
    }
    "Project Management" {
        # AAD groups
        $M365Groups += ";SG.LastPass.Projects;SG.Role.Projects;SG.Officevibe.Projects"

        # Groups to deploy apps
        $M365Groups += ";SG.App.Microsoft365.Project;SG.App.Microsoft365.Visio"

        # Licenses
        $M365Groups += ";SG.License.PowerApp;SG.License.Visio.Plan2;SG.License.Project.Plan5"
        $Licenses += ";Visio Plan 2:Microsoft:null:VISIOCLIENT:null"
        $Licenses += ";Project Plan 5:Microsoft:null:PROJECTPREMIUM:null"

        # Teams
        $M365Groups += ";Project Delivery"

        # Shared mailboxes
        $SharedMailboxes += ";projectupdates@manganoit.com.au:FullAccess;"

        # Distribution groups
        $DistributionGroups += ";AllTechs@manganoit.com.au"
    }
    "Professional Services" {
        # AAD groups
        $M365Groups += ";SG.LastPass.ServiceDesk"

        # Groups to deploy apps
        $M365Groups += ";SG.App.Microsoft365.Project;SG.App.Microsoft365.Visio"

        # Licenses
        $M365Groups += ";SG.License.PowerApp;SG.License.Visio.Plan2;SG.License.Project.Plan5"
        $Licenses += ";Visio Plan 2:Microsoft:null:VISIOCLIENT:null"
        $Licenses += ";Project Plan 5:Microsoft:null:PROJECTPREMIUM:null"

        # Teams
        $M365Groups += ";Project Delivery"

        # Shared mailboxes
        $SharedMailboxes += ";projectupdates@manganoit.com.au:FullAccess;"

        # Distribution groups
        $DistributionGroups += ";AllTechs@manganoit.com.au"
    }
    "Sales & Marketing" {
        # AAD groups
        $M365Groups += ";SG.App.DuoSecurity;SG.App.Bullphish;SG.Role.SSO.CWSell;SG.Role.Sales;SG.Role.AutomationAccount.AutomationOperator"
        $M365Groups += ";SG.Officevibe.Sales;!All Sales"

        # Groups to deploy apps
        $M365Groups += ";SG.App.Microsoft365.Visio"

        # Licenses
        $M365Groups += ";SG.License.PowerApp;SG.License.Visio.Plan2"
        $Licenses += ";Visio Plan 2:Microsoft:null:VISIOCLIENT:null"

        # Teams
        $M365Groups += ";Marketing Team;Sales Team"

        # Shared mailboxes
        $SharedMailboxes += ";sales@manganoit.com.au:FullAccess;agmt.balance@manganoit.com.au:FullAccess;tenders@manganoit.com.au:FullAccess"
        $SharedMailboxes += ";salesadmin@manganoit.com.au:FullAccess"
    }
}

## SEND DATA TO FLOW ##

$Output = @{
    M365Groups = $M365Groups
    SharedMailboxes = $SharedMailboxes
    DistributionGroups = $DistributionGroups
    Licenses = $Licenses
}

Write-Output $Output | ConvertTo-Json