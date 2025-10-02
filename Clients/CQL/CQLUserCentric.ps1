$WorkspaceName = "CQL-PRD3-SENTINEL"
$ResourceGroupName = "cql-sentinel-rg"
$Subscription = "f954f177-4db6-46db-9a99-2bf59d22c1fa"

Import-Module ActiveDirectory
Import-Module Az.Accounts
Import-Module Az.OperationalInsights

function Get-EnabledUserCount {
    $Count = (Get-ADUser -Filter * | Where-Object { $_.Enabled -eq "True" }).Count
    return $Count
}

function Get-ChangedPasswords {
    $Date = (Get-Date).AddDays(-7)
    $Passwords = Get-ADUser -Filter { pwdLastSet -gt $Date } -Properties Name, PasswordLastSet, lastLogon |
    Select-Object Name, UserPrincipalName, PasswordLastSet, @{N = 'LastLogon'; E = { [DateTime]::FromFileTime($_.LastLogon) } }
    return $Passwords
}

function Get-CreatedAccounts {
    $Date = (Get-Date).AddDays(-7)
    $Created = Get-ADUser -Filter { (Enabled -eq $True) -and (whencreated -gt $Date) } -Properties Name, whencreated |
    Select-Object Name, UserPrincipalName, whencreated
    return $Created
}

function Get-ModifiedDisabledAccounts {
    $Date = (Get-Date).AddDays(-7)
    $ModifiedDisabled = Get-ADUser -Filter { (Enabled -eq $False) -and (whenChanged -gt $Date) } -Properties Name, whenChanged |
    Select-Object Name, UserPrincipalName, whenChanged
    return $ModifiedDisabled
}

function Get-ModifiedAccounts {
    $Date = (Get-Date).AddDays(-7)
    $Modified = Get-ADUser -Filter { (Enabled -eq $True) -and (whenChanged -gt $Date) -and (whencreated -lt $Date) } -Properties Name, whenChanged |
    Select-Object Name, UserPrincipalName, whenChanged
    return $Modified
}

function Get-InactiveADUsers {
    $Date = (Get-Date).AddDays(-90)
    $InactiveADUsers = Get-ADUser -Filter { Enabled -eq $True -and LastLogonDate -lt $Date } -Properties Name, LastLogonDate |
    Select-Object Name, UserPrincipalName, LastLogonDate
    return $InactiveADUsers
}

function Get-InactiveAADUsers {
    param (
        [string]$WorkspaceName,
        [string]$ResourceGroupName,
        [string]$Subscription
    )

    Disable-AzContextAutosave -Scope Process | Out-Null

    try {
        Connect-AzAccount -Identity -Subscription $Subscription | Out-Null
    } catch {
        Write-Error "Failed to connect to Azure: $_"
        return
    }
    $Workspace = Get-AzOperationalInsightsWorkspace -ResourceGroupName $ResourceGroupName -Name $WorkspaceName

    $Query = "SigninLogs`n"
    $Query += "| where TimeGenerated < ago(90d)`n"
    $Query += "| summarize LastSignIn=max(TimeGenerated) by UserPrincipalName`n"
    $Query += "| join kind=inner (`n"
    $Query += "AADUserRiskEvents`n"
    $Query += "| where RiskLevel == 'none' // Adjust this filter based on how you define 'active' users`n"
    $Query += "| summarize by UserPrincipalName`n"
    $Query += ") on UserPrincipalName`n"
    $Query += "| where LastSignIn < ago(90d)`n"
    $Query += "| project UserPrincipalName, LastSignIn"
    
    $QueryResults = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query

    Write-Output $QueryResults.Results
}

function Get-DisabledAccounts {
    param (
        [string]$WorkspaceName,
        [string]$ResourceGroupName,
        [string]$Subscription
    )

    Disable-AzContextAutosave -Scope Process | Out-Null

    try {
        Connect-AzAccount -Identity -Subscription $Subscription | Out-Null
    } catch {
        Write-Error "Failed to connect to Azure: $_"
        return
    }
    $Workspace = Get-AzOperationalInsightsWorkspace -ResourceGroupName $ResourceGroupName -Name $WorkspaceName

    $Query = "SecurityEvent`n"
    $Query += "| where EventSourceName == 'Microsoft-Windows-Security-Auditing'`n"
    $Query += "| where EventID == 4725`n"
    $Query += "| where TimeGenerated > ago(7d)`n"
    $Query += "| extend TimeGeneratedAEST = datetime_utc_to_local(TimeGenerated, 'Australia/Sydney')`n"
    $Query += "| sort by TimeGeneratedAEST`n"
    $Query += "| project TimeGeneratedAEST, DisabledAccount = TargetUserName, Actor = SubjectUserName, Computer"

    $QueryResults = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query

    Write-Output $QueryResults.Results
}

function Convert-ToHtml {
    param (
        [Parameter(Mandatory)]
        $data,
        [Parameter(Mandatory)]
        $title
    )
    $html = "<button onclick=""toggleVisibility('$($title.Replace(' ', ''))Section')"">$title</button>`n"
    $html += "<div id=""$($title.Replace(' ', ''))Section"" style=""display:none;"">`n"
    $html += "$(($data | ConvertTo-Html -Body ""<h2>$title</h2>""))`n"
    $html += "</div>"
    

    return $html
}

function New-Report {
    $reportdate = Get-Date -Format D

    $Count = Get-EnabledUserCount
    $Passwords = Get-ChangedPasswords
    $Created = Get-CreatedAccounts
    $Disabled = Get-DisabledAccounts -WorkspaceName $WorkspaceName -ResourceGroupName $ResourceGroupName -Subscription $Subscription
    $ModifiedDisabled = Get-ModifiedDisabledAccounts
    $Modified = Get-ModifiedAccounts
    $InactiveADUsers = Get-InactiveADUsers
    $InactiveAADUsers = Get-InactiveAADUsers

    $PasswordsHTML = Convert-ToHtml $Passwords "Passwords changed this week"
    $CreatedHTML = Convert-ToHtml $Created "Accounts created this week"
    $DisabledHTML = Convert-ToHtml $Disabled "AD accounts disabled this week"
    $ModifiedDisabledHTML = Convert-ToHtml $ModifiedDisabled "Disabled accounts modified this week"
    $ModifiedHTML = Convert-ToHtml $Modified "Enabled accounts modified this week"
    $InactiveADUsersHTML = Convert-ToHtml $InactiveADUsers "Inactive AD users (90 days)"
    $InactiveAADUsersHTML = Convert-ToHtml $InactiveAADUsers "Inactive AAD users (90 days)"
    
    $htmlcontent = $PasswordsHTML + $CreatedHTML + $DisabledHTML + $ModifiedDisabledHTML + $ModifiedHTML + $InactiveADUsersHTML + $InactiveAADUsersHTML

    $htmlhead = "<html>`n"
    $htmlhead += "<body>`n"
    $htmlhead += "This is the weekly User Account Administration Report. For more information, please contact the CleanCo Service Desk at service.desk@cleancoqld.com.au or on (07) 3151 9066`n"
    $htmlhead += "<h1>Weekly Summary: CQL - CleanCo Queensland Ltd</h1>`n"
    $htmlhead += "<h3>Generated: $reportdate</h3>`n"
    $htmlhead += "<p>Total user count: $Count</p>"

    $htmltail = "</body>`n"
    $htmltail += "</html>"

    $htmlbody = $htmlhead + $htmlcontent + $htmltail

    $htmlreport = "<html>`n"
    $htmlreport += "<head>`n"
    $htmlreport += "<style>`n"
    $htmlreport += "body { font-family: Arial, sans-serif; margin: 20px; }`n"
    $htmlreport += "h1, h2, h3 { color: #333366; }`n"
    $htmlreport += "table { width: 100%; border-collapse: collapse; }`n"
    $htmlreport += "th, td { border: 1px solid #999; padding: 8px; text-align: left; }`n"
    $htmlreport += "th { background-color: #f2f2f2; }`n"
    $htmlreport += "tr:nth-child(even) { background-color: #f9f9f9; }`n"
    $htmlreport += "</style>`n"
    $htmlreport += "<script>`n"
    $htmlreport += "function toggleVisibility(id) {`n"
    $htmlreport += "var element = document.getElementById(id);`n"
    $htmlreport += "if (element.style.display === 'none') {`n"
    $htmlreport += "element.style.display = 'block';`n"
    $htmlreport += "} else {`n"
    $htmlreport += "element.style.display = 'none';`n"
    $htmlreport += "}`n"
    $htmlreport += "}`n"
    $htmlreport += "</script>`n"
    $htmlreport += "</head>`n"
    $htmlreport += "<body>`n"
    $htmlreport += "$htmlbody`n"
    $htmlreport += "</body>`n"
    $htmlreport += "</html>"

    return $htmlreport
}

$htmlreport = New-Report

Write-Output $htmlreport