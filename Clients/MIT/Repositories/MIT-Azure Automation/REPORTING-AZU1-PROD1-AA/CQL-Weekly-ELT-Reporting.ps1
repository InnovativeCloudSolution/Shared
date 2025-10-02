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
$Date = (Get-Date).AddHours(10).ToString('yyyy-MM-dd')

$TicketID=0
$Results = @()
$TimeEntires = @()
$FileName = $Date+'_ELT_Report.csv'
$TicketCount=0
$VIPCount=0
$NoteCount=0
$LMTotal=0
$MTDTotal=0
$MTDTickets=0
$MTDVIPTickets=0
$OpenVIPTickets=0
$MTDTicketNoteCount=0
$MTDTicketTimeCount=0

$LMTicketNoteCount=0
$LMTicketTimeCount=0

$StartLastWeek = (Get-Date).AddDays(-7).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')
$StartThisMonth=(Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')
$StartLastMonth=(Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0).AddMonths(-1).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')

$Tickets = Get-CWMTicket -condition "company/name like '*CleanCo Queensland Ltd*' and (lastUpdated >= [$StartLastMonth] or closedFlag = false)" -all
$TotalTime=0
# This week
foreach ($Ticket in $Tickets){
    $VIPFlag = ""
    $TicketID = $Ticket.id
    if($Ticket._info.lastUpdated -ge $StartThisMonth -or $Ticket.closedFlag -eq $false){
        $MTDTickets++
        foreach($types in $Contact.types){
            if($types.id -eq 17){
                $MTDVIPTickets++
            }
        }
    }
    if($Ticket._info.lastUpdated -ge $StartLastWeek -or $Ticket.closedFlag -eq $false){
        $TicketCount++
        $TicketSummary = $Ticket.summary
        $TicketHours = $Ticket.actualHours
        $WeekTicketNote = Get-CWMTicketNote -TicketID $TicketID -condition "dateCreated >= [$StartLastWeek]" -all
        $WeekTicketTime  = Get-CWMTimeEntry -Condition "ticket/id = $TicketID and dateEntered >= [$StartLastWeek]" -all

        $WeekTicketNoteCount = if($WeekTicketNote -ne $null){if(($WeekTicketNote).GetType().Name -ne 'PSCustomObject'){($WeekTicketNote).Count} else {1}} else{0}
        $WeekTicketTimeCount = if($WeekTicketTime -ne $null){if(($WeekTicketTime).GetType().Name -ne 'PSCustomObject'){($WeekTicketTime).Count} else {1}} else{0}
        $WeekTotal = $WeekTicketNoteCount+$WeekTicketTimeCount
        $NoteCount+=$WeekTotal
        $TotalTime+=$WeekTicketTimeCount

        $TotalTicketNote = Get-CWMTicketNote -TicketID $TicketID -all
        $TotalTicketTime  = Get-CWMTimeEntry -Condition "ticket/id = $TicketID" -all

        $TotalTicketNoteCount = if($TotalTicketNote -ne $null){if(($TotalTicketNote).GetType().Name -ne 'PSCustomObject'){($TotalTicketNote).Count} else {1}} else{0}
        $TotalTicketTimeCount = if($TotalTicketTime -ne $null){if(($TotalTicketTime).GetType().Name -ne 'PSCustomObject'){($TotalTicketTime).Count} else {1}} else{0}
        $Total = $TotalTicketNoteCount+$TotalTicketTimeCount

        $ContactID = $Ticket.contact.id
        
        $Contact = Get-CWMContact -Condition "id=$ContactID"

        foreach($types in $Contact.types){
            if($types.id -eq 17){
                $VIPFlag = $types.name;
                $VIPCount++
                if($Ticket.closedFlag -eq $false){
                    $OpenVIPTickets++
                }
            }
        }

        $LastUpdated = $Ticket._info.lastUpdated
        $Results += New-Object -TypeName PSObject -Property ([ordered]@{TicketNumber=$TicketID; Summary=$TicketSummary; Hours=$TicketHours; TotalComms=$Total; WeeklyComms=$WeekTotal; Owner=$Ticket.owner.name; Priority=$Ticket.priority.name;Status=$Ticket.status.name;Agreement=$Ticket.agreement.name;VIP=$VIPFlag; LastUpdated=$LastUpdated})
    }

    #MTD Month To Date
    $MTDTicketNote = Get-CWMTicketNote -TicketID $TicketID -condition "dateCreated >= [$StartThisMonth]" -all
    $MTDTicketTime  = Get-CWMTimeEntry -Condition "ticket/id = $TicketID and timeStart >= [$StartThisMonth]" -all
    $MTDTicketNoteCount += if($MTDTicketNote -ne $null){if(($MTDTicketNote).GetType().Name -ne 'PSCustomObject'){($MTDTicketNote).Count} else {1}} else{0}
    $MTDTicketTimeCount += if($MTDTicketTime -ne $null){if(($MTDTicketTime).GetType().Name -ne 'PSCustomObject'){($MTDTicketTime).Count} else {1}} else{0}

    #LM Last Month
    $LMTicketNote = Get-CWMTicketNote -TicketID $TicketID -condition "dateCreated >= [$StartLastMonth] and dateCreated <[$StartThisMonth]" -all
    $LMTicketTime  = Get-CWMTimeEntry -Condition "ticket/id = $TicketID and timeStart >= [$StartLastMonth] and timeStart <[$StartThisMonth]" -all
    $LMTicketNoteCount += if($LMTicketNote -ne $null){if(($LMTicketNote).GetType().Name -ne 'PSCustomObject'){($LMTicketNote).Count} else {1}} else{0}
    $LMTicketTimeCount += if($LMTicketTime -ne $null){if(($LMTicketTime).GetType().Name -ne 'PSCustomObject'){($LMTicketTime).Count} else {1}} else{0}
    
}
Disconnect-CWM


$Results | Export-Csv "C:\temp\$FileName" -NoTypeInformation

$LMTotal=$LMTicketNoteCount+$LMTicketTimeCount
$MTDTotal=$MTDTicketNoteCount+$MTDTicketTimeCount


# Email Result
$sendGridApiKey = Get-AutomationVariable -Name 'SendGrid Azure Automation API'

$Subject = "CleanCo Weekly Support Interaction Report"
$Body = @"
<html><body>Hi CleanCo ELT,<br>
Attached is the Weekly Mangano IT Reporting.<br>
Here is an overview of the statistics from this week:<br>
Tickets worked on last week: $TicketCount<br>
Ticket interactions last week: $NoteCount<br>
VIP Tickets last week: $VIPCount<br>
VIP Tickets Open: $OpenVIPTickets<br>
Month To Date Tickets: $MTDTickets<br>
Month To Date VIP Tickets: $MTDVIPTickets<br>
Month To Date Interactions: $MTDTotal<br>
Last Month Interactions: $LMTotal<br><br>
Regards,<br><br>
Mangano IT</body></html>
"@
$SendGridEmail = @{
	From = 'Mangano IT - Workflows <workflows@manganoit.com.au>'
	To = "ict@cleancoqld.com.au"
    CC = "Robert.Mangano@manganoit.com.au","Alex.Williams@manganoit.com.au","csm@manganoit.com.au","sdm@manganoit.com.au","ict.business.partners@cleancoqld.com.au","cybersecurity@cleancoqld.com.au"
	Subject = $Subject
	Body = $Body
	
	SmtpServer = 'smtp.sendgrid.net'
	Port = 587
	UseSSL = $true
	Credential = New-Object PSCredential 'apikey', (ConvertTo-SecureString $sendGridApiKey -AsPlainText -Force)	
    Attachments = "C:\temp\$FileName"
}
Send-MailMessage @SendGridEmail -BodyAsHtml


#$SharePointApp = Get-AutomationPSCredential -Name "SharePoint App"
#Connect-PnPOnline -Url "https://manganoit.sharepoint.com/clients" -AppId $SharePointApp.UserName -AppSecret $SharePointApp.GetNetworkCredential().Password
#$AddedFile = Add-PnPFile -Path "C:\temp\temp.csv" -Folder "Client Data/Clean Co Queensland/CleanCo-WeeklyReporting" -NewFileName $FileName
#Disconnect-PnPOnline

$json = @"
{
"TicketCount": "$TicketCount",
"NoteCount": "$NoteCount",
"VIPCount": "$VIPCount",
"MTDVIPTickets": "$MTDVIPTickets",
"MTDTickets": "$MTDTickets",
"OpenVIPTickets": "$OpenVIPTickets",
"MTDTicketNoteCount": "$MTDTicketNoteCount",
"MTDTicketTimeCount": "$MTDTicketTimeCount",
"MTDTotal": "$MTDTotal",
"LMTicketNoteCount": "$LMTicketNoteCount",
"LMTicketTimeCount": "$LMTicketTimeCount",
"LMTotal":"$LMTotal",
"FileName": "$FileName"
}
"@
Write-Output $json