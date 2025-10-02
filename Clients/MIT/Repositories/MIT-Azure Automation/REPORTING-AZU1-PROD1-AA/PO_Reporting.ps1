    param(
    [string]$CompanyIdentifier="MIT",
    [string]$PONumber = '',
    [string]$AgreementType='All',
    [string]$ReportSpan='Forever',
    [string]$Requestor='Alex.Williams@manganoit.com.au'
    )
Set-StrictMode -Off
# Set CRM variables, connect to server
$Server = 'https://api-aus.myconnectwise.net'
$Company = 'mit'
$pubkey = Get-AutomationVariable -Name 'PublicKey'
$privatekey = Get-AutomationVariable -Name 'PrivateKey'
$clientId = Get-AutomationVariable -Name 'clientId'

# Create a credential object
$Connection = @{
    Server = $Server
    Company = $Company
    pubkey = $pubkey
    privatekey = $privatekey
    clientId = $clientId
}
# Connect to Manage server
Connect-CWM @Connection

$PONumber =$PONumber -replace '\s',''
$Conditions = 'company/identifier like "%'+$CompanyIdentifier+'%" '

if($AgreementType -eq 'Managed Services'){
    $Conditions += " and agreement/name like '%Managed Services%' "
}elseif($AgreementType -eq "PO"){
    $Conditions +=' and agreement/name like "*'+$PONumber+'*" '
}

if($AgreementType -eq 'Managed Services' -or $AgreementType -eq 'All'){
    if($ReportSpan -eq '1 Month'){
        $ReportStart=(Get-Date).AddMonths(-1).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $Conditions+= " and lastUpdated >= [$ReportStart]"
    }elseif($ReportSpan -eq '3 Months'){
        $ReportStart=(Get-Date).AddMonths(-3).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $Conditions+= " and lastUpdated >= [$ReportStart]"
    }elseif($ReportSpan -eq '6 Months'){
        $ReportStart=(Get-Date).AddMonths(-6).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $Conditions+= " and lastUpdated >= [$ReportStart]"
    }else{
        $ReportStart=(Get-Date).AddMonths(-6).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $Conditions+= " and lastUpdated >= [$ReportStart]"
    }
}

Write-Output $Conditions

$TimeEntries = Get-CWMTimeEntry -Condition $Conditions -all
$AEST = [System.TimeZoneInfo]::GetSystemTimeZones().Where({$_.StandardName -eq 'AUS Eastern Standard Time'})[0]
$Results = @()

#If the time entry count is over 999, it splits it into an array of arrays. This will check if we have an array of arrays of objects or an array of objects
if($TimeEntries.GetType() -eq $TimeEntries[0].GetType()){
    foreach ($Subset in $TimeEntries){
        foreach ($TimeEntry in $Subset){
            $Cost = $TimeEntry.actualHours*$TimeEntry.hourlyRate
            if(![string]::IsNullOrEmpty($TimeEntry.notes)){
                if($TimeEntry.notes[0] -eq '-'){
                    $TimeEntry.notes = $TimeEntry.notes.substring(1)
                }
                $TimeEntry.notes =$TimeEntry.notes -replace '\n',' '
                if($TimeEntry.notes.length -gt 32000) {$TimeEntry.notes = $TimeEntry.notes.substring(0,32000)}
            }
            $timeDate = [datetime]::parseexact($TimeEntry.timeStart, 'yyyy-MM-ddTHH:mm:ssZ', $null).AddHours(10).ToString('yyyy-MM-dd');

            $Ticket=Get-CWMTicket -id $TimeEntry.chargeToId
            $Type= $Ticket.type.name
            $Subtype= $Ticket.subType.name
            $Item= $Ticket.item.name
            $Results += New-Object -TypeName PSObject -Property ([ordered]@{TimeID=$TimeEntry.id; TicketNumber=$TimeEntry.chargeToId; Summary=$TimeEntry.ticket.Summary; Technician=$TimeEntry.member.name; timeDate=$timeDate; hoursBilled=$TimeEntry.hoursBilled; Cost=$Cost; Notes=$TimeEntry.notes; Agreement=$TimeEntry.agreement.name; Type=$Type; subType=$Subtype; item=$Item; status=$Ticket.status.name})
        }
    }
}else{
    foreach ($TimeEntry in $TimeEntries){
        $Cost = $TimeEntry.actualHours*$TimeEntry.hourlyRate
        if(![string]::IsNullOrEmpty($TimeEntry.notes)){
            if($TimeEntry.notes[0] -eq '-'){
                $TimeEntry.notes = $TimeEntry.notes.substring(1)
            }
            $TimeEntry.notes =$TimeEntry.notes -replace '\n',' '
            if($TimeEntry.notes.length -gt 32000) {$TimeEntry.notes = $TimeEntry.notes.substring(0,32000)}
        }
        $timeDate = [datetime]::parseexact($TimeEntry.timeStart, 'yyyy-MM-ddTHH:mm:ssZ', $null).AddHours(10).ToString('yyyy-MM-dd');

        $Ticket=Get-CWMTicket -id $TimeEntry.chargeToId
        $Type= $Ticket.type.name
        $Subtype= $Ticket.subType.name
        $Item= $Ticket.item.name
        $Results += New-Object -TypeName PSObject -Property ([ordered]@{TimeID=$TimeEntry.id; TicketNumber=$TimeEntry.chargeToId; Summary=$TimeEntry.ticket.Summary; Technician=$TimeEntry.member.name; timeDate=$timeDate; hoursBilled=$TimeEntry.hoursBilled; Cost=$Cost; Notes=$TimeEntry.notes; Agreement=$TimeEntry.agreement.name; Type=$Type; subType=$Subtype; item=$Item; status=$Ticket.status.name})
    }
}
$Now=(Get-Date).ToString('yyyy-MM-dd')
$FileName = "$CompanyIdentifier $AgreementType$PONumber $ReportSpan Report - $Now"

$Results | Export-Csv "C:\temp\$FileName.csv" -NoTypeInformation


# Email Result
$sendGridApiKey = Get-AutomationVariable -Name 'SendGrid Azure Automation API'

$Subject = "Agreement Reconciliation: "+$FileName
$Body = "Please find attached the agreement reconciliation report for "+$FileName
$SendGridEmail = @{
	From = 'workflows@manganoit.com.au'
	To = $Requestor
	Subject = $Subject
	Body = $Body
	
	SmtpServer = 'smtp.sendgrid.net'
	Port = 587
	UseSSL = $true
	Credential = New-Object PSCredential 'apikey', (ConvertTo-SecureString $sendGridApiKey -AsPlainText -Force)	
    Attachments = "C:\temp\$FileName.csv"
}
Send-MailMessage @SendGridEmail

Disconnect-CWM

#$SharePointApp = Get-AutomationPSCredential -Name "SharePoint App"
#Connect-PnPOnline -Url "https://manganoit.sharepoint.com/clients" -AppId $SharePointApp.UserName -AppSecret $SharePointApp.GetNetworkCredential().Password
#Write-Output "Adding PNP File "+$FileName 
#$AddedFile = Add-PnPFile -Path "C:\temp\temp.csv" -Folder "Client Data/Clean Co Queensland/CleanCo-BucketReporting" -NewFileName $FileName
#Disconnect-PnPOnline