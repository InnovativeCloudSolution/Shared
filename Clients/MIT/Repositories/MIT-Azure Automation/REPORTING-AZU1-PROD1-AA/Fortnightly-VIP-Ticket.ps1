<#

Mangano IT - Get List of CleanCo VIPs Fortnightly
Created by: Alex Williams
Maintained by: Gabriel Nugent
Version: 2.0.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

## SCRIPT VARIABLES ##

Set-StrictMode -Off

# Create a credential object
$Connection = .\CWM-CreateConnectionObject.ps1

# Connect to Manage server
Connect-CWM @Connection
$Ticket = Get-CWMTicket -condition 'company/name like "CleanCo Queensland Ltd" and closedFlag = false and summary like "CleanCo - Send VIP Users to Telstra BTSM"'
$Users = Get-CWMContact -condition 'company/name like "CleanCo Queensland Ltd"' -all

foreach ($User in $Users) {
    try {
        # Get the VIP? field
        foreach ($customField in $User.customFields) {
            if ($customField.caption -eq "VIP?") {
                $VIPField = $customField.value
            }
        }

        # Check if the VIP field is true, if so, proceed
        if ($VIPField -eq "True" -and !$User.inactiveFlag) {
            $First = $User.firstName
            $Last = $User.lastName
            #Get the default phone number
            foreach ($communicationItem in $User.communicationItems) {
                if ($communicationItem.communicationType -eq "Phone") {
                    $Mobile = $communicationItem.value
                }
                if ($communicationItem.communicationType -eq "Email") {
                    $Email = $communicationItem.value
                }
            }


            $User = " - $First $Last. Phone Number: $Mobile. Email: $Email`r`n"
            $Output += $User
         }
     }catch{
     }
}

# Add users to ticket as note
New-CWMTicketNote -ticketId $Ticket.id -text $Output -detailDescriptionFlag $True