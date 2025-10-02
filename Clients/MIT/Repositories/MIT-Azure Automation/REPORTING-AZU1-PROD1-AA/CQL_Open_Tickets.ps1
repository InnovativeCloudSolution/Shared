Set-StrictMode -Off
# Set CRM variables, connect to server
$Server = 'https://api-aus.myconnectwise.net'
$Company = 'mit'
$pubkey = Get-AutomationVariable -Name 'PublicKey'
$privatekey = Get-AutomationVariable -Name 'PrivateKey'
$clientId = Get-AutomationVariable -Name 'clientId'

$Recipients = "ict.business.partners@cleancoqld.com.au"
$CCs = "Alex.Williams@manganoit.com.au","csm@manganoit.com.au","sdm@manganoit.com.au"

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
$Results = @()
$Date = (Get-Date).AddHours(10).ToString('yyyy-MM-dd')
$FileName = $Date+'_Ticket_Report.csv'

$Tickets = Get-CWMTicket -condition "company/name like '*CleanCo Queensland Ltd*' and closedFlag = false and board/name like 'CQL*' and status/name !='Closed~' and status/name !='Closed - Silent' and status/name != 'Completed~'" -all

foreach ($Ticket in $Tickets){
    $TicketID=$Ticket.id
    $DateEntered=$Ticket._info.dateEntered
    $CreatedTime = [datetime]::ParseExact($DateEntered,"yyyy-MM-ddTHH:mm:ssZ", $null)
    $Age= [Math]::Round((New-TimeSpan -Start $CreatedTime -End (Get-Date)).TotalDays,1)
    $Status=$Ticket.status.name
    $Summary=$Ticket.summary
    $Contact=$Ticket.contactName
    $Resources=$Ticket.Resources
    $LastUpdate=$Ticket._info.lastUpdated
    $TotalHours=$Ticket.actualHours
    $Type=$Ticket.type.name
    $SubType=$Ticket.subType.name
    $Item=$Ticket.item.name
    $Site=$Ticket.siteName
    $Source=$Ticket.source.name

    $Results += New-Object -TypeName PSObject -Property ([ordered]@{TicketNumber=$TicketID; Age=$Age; Status=$Status; Summary=$Summary; Contact=$Contact; Resources=$Resources; LastUpdate=$LastUpdate; TotalHours=$TotalHours; Type=$Type; SubType=$SubType; Item=$Item; Site=$Site; Source=$Source})
}

Disconnect-CWM

$Results | Export-Csv "C:\temp\CQL_Open_Tickets.csv" -NoTypeInformation

# Email Result
$sendGridApiKey = Get-AutomationVariable -Name 'SendGrid Azure Automation API'

$Subject = "CleanCo Daily Ticket Report"
$Body = @"
<html><body>Hi ICT Business Partners,<br><br>
Attached is the daily CleanCo Ticket Reporting.<br><br>
Regards,<br><br>
Mangano IT</body></html>
"@

$SendGridEmail = @{
	From = 'support@manganoit.com.au'
	To = $Recipients
    CC = $CCs
	Subject = $Subject
	Body = $Body
	
	SmtpServer = 'smtp.sendgrid.net'
	Port = 587
	UseSSL = $true
	Credential = New-Object PSCredential 'apikey', (ConvertTo-SecureString $sendGridApiKey -AsPlainText -Force)	
    Attachments = "C:\temp\CQL_Open_Tickets.csv"
}
Send-MailMessage @SendGridEmail -BodyAsHtml

#$SharePointApp = Get-AutomationPSCredential -Name "SharePoint App"
#Connect-PnPOnline -Url "https://manganoit.sharepoint.com/clients" -AppId $SharePointApp.UserName -AppSecret $SharePointApp.GetNetworkCredential().Password
#$AddedFile = Add-PnPFile -Path "C:\temp\CQL_Open_Tickets.csv" -Folder "Client Data/Clean Co Queensland/CleanCo-TicketReporting" -NewFileName $FileName
#Disconnect-PnPOnline

#$json = @"
#{
#"FileName": "$FileName"
#}
#"@
#Write-Output $json