param(
    [string]$CompanyIdentifier = "ICS",
    [int]$Year  = 2025,
    [int]$Month = 12,
    [int]$TicketsPerDay = 5
)

. "$PSScriptRoot\CWM-Common.ps1"

# ============================================================
# TICKET SCENARIOS
# Each scenario maps to a real board/type/subtype combination
# and includes a pool of realistic summaries + conversation
# threads between the client contact and Glenn Farnsworth.
# ============================================================
$Scenarios = @(
    @{
        Weight  = 20
        Board   = "Service Desk"; Type = "Access"; SubType = "Password"; Item = "Incident"
        Summaries = @(
            "Unable to log in - password not working",
            "Account locked out after multiple failed attempts",
            "Password reset required - forgot credentials",
            "Cannot access email - password rejected"
        )
        ContactOpeners = @(
            "Hi, I can't log into my account. I've tried resetting my password but it's still not working. Can you help?",
            "My account seems to be locked. I entered my password a few times and now it won't let me in at all.",
            "I need a password reset please. I've been locked out since this morning.",
            "I can't access my email. It keeps saying my password is incorrect but I haven't changed it."
        )
        TechReplies = @(
            "Hi {contact}, thanks for reaching out. I can see your account is locked. I've reset it now - please try logging in with the temporary password I've sent to your mobile.",
            "Hey {contact}, no worries. I've unlocked your account and sent a password reset link to your recovery email. Let me know once you're back in.",
            "Hi {contact}, I've reset your credentials. Your new temporary password has been sent via SMS. Please change it on first login.",
            "Hi {contact}, I can see the issue. Your account was flagged after 5 failed attempts. I've cleared the lock - try again now."
        )
        ContactFollowUps = @(
            "That worked, thank you! All good now.",
            "I'm back in, thanks Glenn!",
            "Perfect, I can log in now. Thanks for the quick fix.",
            "All sorted, cheers!"
        )
        TechResolutions = @(
            "Great! Account unlocked and access restored. Closing this ticket.",
            "Glad that worked. I've noted the issue. Closing off now.",
            "Access restored. If this happens again let us know - may be worth reviewing MFA settings.",
            "Resolved. Closing ticket."
        )
        Duration = @(0.25, 0.5, 0.5, 0.75)
        BillableOption = "Billable"
        Priority = "Priority 3 - Normal Response"
    },
    @{
        Weight  = 15
        Board   = "Service Desk"; Type = "Collaboration"; SubType = "Email"; Item = "Incident"
        Summaries = @(
            "Not receiving emails from external senders",
            "Emails going to spam - client complaints",
            "Unable to send emails - getting bounce-back errors",
            "Outlook not syncing emails on mobile"
        )
        ContactOpeners = @(
            "Hi, clients are telling me they've emailed me but I'm not receiving anything. Can you check?",
            "My emails keep going into people's spam folders. It's really causing issues with clients.",
            "I'm getting bounce-back errors when I try sending emails externally. Error code 550.",
            "My Outlook on my phone stopped syncing about an hour ago. I'm missing emails."
        )
        TechReplies = @(
            "Hi {contact}, I'm checking your mail flow now. Can you confirm one sender's email address so I can trace the message?",
            "Hey {contact}, I can see a potential SPF/DKIM issue. Let me check the DNS records and I'll update you shortly.",
            "Hi {contact}, I can see the 550 errors. The recipient domain is rejecting your emails. I'm investigating whether it's an IP reputation issue.",
            "Hi {contact}, I've checked the Exchange sync status. There's a policy blocking ActiveSync on your mobile profile. I'll fix that now."
        )
        ContactFollowUps = @(
            "The sender was john@acmecorp.com.au - he's sent 3 emails in the last hour.",
            "Thanks, let me know what you find. It's been happening since yesterday.",
            "It's happening to all external addresses, not just one.",
            "Thanks. My phone is an iPhone 14 if that helps."
        )
        TechResolutions = @(
            "Mail flow restored. The email was caught by our anti-spam rule - I've whitelisted the sender. Closing ticket.",
            "SPF record updated and mail flow confirmed clean. Closing off.",
            "IP removed from blocklist and outbound routing adjusted. Email sending restored. Closing ticket.",
            "ActiveSync profile updated. Mobile sync confirmed working. Closing ticket."
        )
        Duration = @(0.5, 0.75, 1.0, 0.5)
        BillableOption = "Billable"
        Priority = "Priority 3 - Normal Response"
    },
    @{
        Weight  = 15
        Board   = "Service Desk"; Type = "Endpoint"; SubType = "Performance"; Item = "Incident"
        Summaries = @(
            "Computer running very slow - taking long to load applications",
            "Laptop fan running loud and machine overheating",
            "PC freezing randomly throughout the day",
            "Applications crashing frequently since Windows update"
        )
        ContactOpeners = @(
            "My computer has been incredibly slow all week. It takes 5 minutes just to open Excel.",
            "My laptop fan is constantly going flat out and it gets really hot. Something feels wrong.",
            "My PC keeps freezing every hour or so. I have to hard restart it. Losing work.",
            "Ever since the update last Tuesday everything keeps crashing. Word, Chrome, everything."
        )
        TechReplies = @(
            "Hi {contact}, I'm remoting in now to check. Please leave your computer on and unlocked.",
            "Hey {contact}, sounds like possible thermal issues. Can you tell me how old the device is? I'll run some diagnostics.",
            "Hi {contact}, I can see the freeze logs. Looks like a memory issue. Running a RAM diagnostic remotely now.",
            "Hi {contact}, I can see the update installed last Tuesday. I'll roll back the problematic components and test."
        )
        ContactFollowUps = @(
            "I've left it on, go ahead.",
            "It's about 3 years old. The fan has been loud for a few months.",
            "OK, I'll leave it running. Should I save everything first?",
            "Yes please, it's really disruptive. I've lost work twice today."
        )
        TechResolutions = @(
            "Cleared temp files, disabled startup bloat, defragged SSD. Performance back to normal. Closing ticket.",
            "Fan and heatsink cleaned remotely confirmed. Thermal paste refresh scheduled. Running fine now. Closing.",
            "RAM slot issue confirmed. Reseated via remote guidance. Stability confirmed over 30 min test. Closing.",
            "Rolled back KB5034441 update. All apps stable. Closing ticket."
        )
        Duration = @(0.75, 1.0, 1.0, 1.25)
        BillableOption = "Billable"
        Priority = "Priority 3 - Normal Response"
    },
    @{
        Weight  = 12
        Board   = "Service Desk"; Type = "Applications"; SubType = "Error"; Item = "Incident"
        Summaries = @(
            "Microsoft Teams crashes on startup",
            "SharePoint file sync not working",
            "OneDrive showing sync errors - files not uploading",
            "Adobe Acrobat failing to open PDFs"
        )
        ContactOpeners = @(
            "Teams keeps crashing as soon as I open it. I've reinstalled it twice and it keeps happening.",
            "SharePoint has stopped syncing my files. The desktop app shows a red X.",
            "OneDrive is showing sync errors on about 20 files. They won't upload.",
            "I can't open any PDFs today. Acrobat opens and then immediately closes."
        )
        TechReplies = @(
            "Hi {contact}, I can see Teams crash logs in your tenant. There's a known conflict with your GPU driver version. Let me push a fix.",
            "Hey {contact}, I can see the sync client is throwing authentication errors. Let me re-link your account.",
            "Hi {contact}, I can see the OneDrive errors. Some files have characters in the name that OneDrive can't sync. I'll identify them.",
            "Hi {contact}, Acrobat has a known issue with the latest Reader DC update. I'll push a patched version now."
        )
        ContactFollowUps = @(
            "Thanks, I rely on Teams for client calls so this is urgent.",
            "It's been like this since yesterday morning.",
            "OK, there are files I need urgently in that folder.",
            "I have documents to sign today so this is quite urgent."
        )
        TechResolutions = @(
            "GPU driver updated and Teams cache cleared. Teams confirmed launching successfully. Closing.",
            "SharePoint account re-authenticated via app. Sync confirmed running. Closing ticket.",
            "Renamed 4 files with invalid characters. All 20 files now synced. Closing ticket.",
            "Acrobat rolled back to 23.x and update blocked. PDFs opening normally. Closing ticket."
        )
        Duration = @(0.5, 0.75, 0.5, 0.75)
        BillableOption = "Billable"
        Priority = "Priority 3 - Normal Response"
    },
    @{
        Weight  = 10
        Board   = "Service Desk"; Type = "Network"; SubType = "Internet"; Item = "Incident"
        Summaries = @(
            "Internet connection dropping intermittently",
            "Very slow internet speeds affecting productivity",
            "VPN not connecting - remote work impacted",
            "WiFi dropping out every 30 minutes"
        )
        ContactOpeners = @(
            "Our internet keeps dropping every hour or so. Teams calls are getting cut off.",
            "The internet is incredibly slow today. Pages are taking ages to load and downloads are crawling.",
            "I can't connect to the VPN from home. It just times out every time.",
            "The WiFi disconnects every 30 minutes or so. I have to reconnect manually each time."
        )
        TechReplies = @(
            "Hi {contact}, I'm checking the router logs remotely now. Can you confirm if it's affecting all devices or just yours?",
            "Hey {contact}, I can see reduced throughput on your connection. Running a line test now. What speeds are you getting?",
            "Hi {contact}, I can see the VPN gateway logs. There's an auth issue with your certificate. I'll push a fix to your machine.",
            "Hi {contact}, I can see the AP logs. Your device is roaming between two access points causing drops. I'll adjust the settings."
        )
        ContactFollowUps = @(
            "It's affecting everyone in the office, not just me.",
            "Speed test shows 2Mbps down. Our plan is 100Mbps.",
            "I've been working from home for 3 days and this just started today.",
            "It happens in the main office area. The back office seems fine."
        )
        TechResolutions = @(
            "Router firmware updated and DHCP lease table cleared. Connection stable for 2 hours. Closing.",
            "ISP confirmed line issue and resolved on their end. Speeds back to 98Mbps. Closing ticket.",
            "VPN certificate renewed and client config pushed. Connection confirmed. Closing ticket.",
            "AP roaming thresholds adjusted. Sticky client issue resolved. WiFi stable. Closing ticket."
        )
        Duration = @(0.75, 0.5, 0.75, 1.0)
        BillableOption = "Billable"
        Priority = "Priority 2 - Quick Response"
    },
    @{
        Weight  = 10
        Board   = "Service Desk"; Type = "Access"; SubType = "MFA"; Item = "Request"
        Summaries = @(
            "MFA setup request for new phone",
            "MFA codes not arriving via SMS",
            "Locked out - lost access to MFA authenticator app",
            "MFA prompting every login - want to review session settings"
        )
        ContactOpeners = @(
            "I've got a new phone and need to set up the authenticator app again.",
            "I'm not receiving the SMS verification codes. I've waited 10 minutes.",
            "I've lost my old phone and can't access the authenticator app. I'm locked out.",
            "MFA is prompting me every single time I log in. Can we adjust the session trust settings?"
        )
        TechReplies = @(
            "Hi {contact}, no problem. I'll send you a QR code link to re-register your authenticator. Which app are you using - Microsoft or Google?",
            "Hey {contact}, I can see the SMS delivery logs - there's a delay from the carrier. As a workaround I'm switching you to app-based MFA now.",
            "Hi {contact}, I'll verify your identity and then reset your MFA. Can you confirm your employee ID and answer your security question?",
            "Hi {contact}, I can adjust the MFA session persistence. Currently set to require auth every login. Want me to extend to trusted devices for 14 days?"
        )
        ContactFollowUps = @(
            "I'm using the Microsoft Authenticator app.",
            "OK thanks, I'll switch to the app. Can you walk me through it?",
            "My employee ID is {empId}. Happy to answer questions to verify.",
            "Yes, 14 days would be perfect. I'm always on my work laptop."
        )
        TechResolutions = @(
            "Authenticator re-registered on new device. MFA confirmed working. Closing ticket.",
            "App-based MFA configured. Tested and confirmed. Old SMS method disabled. Closing.",
            "MFA reset after identity verified. New authenticator registered. Access restored. Closing.",
            "Session trust updated to 14 days for compliant devices. MFA prompts reduced. Closing ticket."
        )
        Duration = @(0.25, 0.5, 0.5, 0.25)
        BillableOption = "Billable"
        Priority = "Priority 3 - Normal Response"
    },
    @{
        Weight  = 8
        Board   = "Service Desk"; Type = "Endpoint"; SubType = "Printing"; Item = "Incident"
        Summaries = @(
            "Printer not responding - jobs stuck in queue",
            "Print quality degraded - faded/streaky output",
            "Unable to print to shared office printer",
            "Printer offline after power outage"
        )
        ContactOpeners = @(
            "The printer isn't printing anything. Jobs are stuck in the queue and nothing is coming out.",
            "The printer output is really faded and streaky. It's making reports look unprofessional.",
            "I can't print to the office printer. It was working yesterday but today it's not showing up.",
            "The printer went offline after the power went out last night. Can't get it back online."
        )
        TechReplies = @(
            "Hi {contact}, I'm checking the print server now. I can see a stuck job from yesterday that's blocking the queue. Clearing it now.",
            "Hey {contact}, sounds like a low toner or drum issue. Can you check the printer panel and tell me what error code it shows?",
            "Hi {contact}, I can see the printer isn't mapping on your machine. The print server had a hiccup. Pushing the driver again now.",
            "Hi {contact}, the printer IP may have changed after the outage. I'm checking the DHCP lease and will update the port config."
        )
        ContactFollowUps = @(
            "Great, thanks! The queue just cleared and it's printing now.",
            "It says 'Toner Low' on the screen.",
            "OK, let me know when it's ready to test.",
            "That makes sense. The power was out for about 2 hours last night."
        )
        TechResolutions = @(
            "Print queue cleared and spooler restarted. All queued jobs printed. Closing ticket.",
            "Toner replacement ordered. Interim workaround: print at draft quality. Closing ticket - new toner arriving tomorrow.",
            "Printer driver redeployed via GPO. Mapping confirmed on all affected machines. Closing.",
            "Printer IP updated to static assignment. Port config corrected. Online and printing. Closing ticket."
        )
        Duration = @(0.25, 0.25, 0.5, 0.5)
        BillableOption = "Billable"
        Priority = "Priority 3 - Normal Response"
    },
    @{
        Weight  = 5
        Board   = "Service Desk"; Type = "Security"; SubType = "Phishing"; Item = "Incident"
        Summaries = @(
            "Suspicious email received - possible phishing attempt",
            "Clicked a link in an email - concerned about security",
            "Received invoice email from unknown sender",
            "Email asking for credentials - flagged as suspicious"
        )
        ContactOpeners = @(
            "I've received a very suspicious email asking me to click a link to verify my account. I haven't clicked it.",
            "I accidentally clicked a link in an email before I realized it looked suspicious. What should I do?",
            "I got an invoice email from someone I don't recognize. The amount is $4,800. Is this legit?",
            "Someone emailed me asking for my Microsoft login credentials claiming to be from IT. This seems wrong."
        )
        TechReplies = @(
            "Hi {contact}, good call not clicking it. Can you forward it to security@dropbear-it.com.au? I'll analyze the headers immediately.",
            "Hey {contact}, don't panic. Can you tell me what the link looked like and whether you entered any details on the page that opened?",
            "Hi {contact}, that is definitely suspicious. Do NOT approve or respond to it. I'm checking your accounts for any related activity now.",
            "Hi {contact}, correct - that is a phishing attempt. We will never ask for credentials via email. I'm quarantining that email now."
        )
        ContactFollowUps = @(
            "Forwarded. The email came from support@micros0ft-account.com.",
            "I clicked it and a page opened that looked like a Microsoft login. I didn't enter anything though.",
            "I don't recognize the sender at all. It came from invoice@billingdept-au.net.",
            "I thought so. I didn't reply. I'll report it to the team."
        )
        TechResolutions = @(
            "Email confirmed phishing. Domain blocked tenant-wide. No compromise detected. User advised. Closing ticket.",
            "No credentials entered. Ran malware scan - clean. URL blocked across tenant. Closing ticket.",
            "Email quarantined. Sender domain blocked. No financial action taken. Finance team alerted. Closing.",
            "Phishing email quarantined and reported to Microsoft. No credential compromise. Closing ticket."
        )
        Duration = @(0.5, 0.75, 0.5, 0.5)
        BillableOption = "Billable"
        Priority = "Priority 2 - Quick Response"
    },
    @{
        Weight  = 5
        Board   = "Service Desk"; Type = "General"; SubType = "How-To Question"; Item = "Request"
        Summaries = @(
            "How to set up email signature in Outlook",
            "How to share a folder in SharePoint",
            "How to set up out of office reply",
            "How to join a Teams meeting from a room device"
        )
        ContactOpeners = @(
            "Hi, can you help me set up my email signature in Outlook? I want it to include my title and mobile number.",
            "I need to share a specific folder in SharePoint with an external user. Can you walk me through it?",
            "I'm going on leave next week and need to set up an out of office reply. How do I do that?",
            "We have a new meeting room device and I'm not sure how to use it to join Teams meetings. Can you help?"
        )
        TechReplies = @(
            "Hi {contact}, of course! In Outlook go to File > Options > Mail > Signatures. Click New, name it, and paste in your details. Want me to remote in and help?",
            "Hey {contact}, sure! In SharePoint go to the folder, click the three dots, select Share, then enter the external email and set permissions to 'Can View' or 'Can Edit'.",
            "Hi {contact}, easy one! In Outlook go to File > Automatic Replies. Turn it on, set your dates, and type your message. Let me know if you want me to set it up.",
            "Hi {contact}, on the room device tap 'Join' on the main screen, then enter the meeting ID or sign in with your Microsoft account. I can walk you through it remotely."
        )
        ContactFollowUps = @(
            "Yes please, I can never find where things are in Outlook.",
            "Thanks! I'll try that. What permission level should I give them? They just need to read files.",
            "I think I can do it myself with those instructions. Thanks!",
            "Can you remote into the room device? I'm not sure which account to use."
        )
        TechResolutions = @(
            "Email signature set up via remote session. User confirmed happy with result. Closing ticket.",
            "SharePoint folder shared with Can View permissions. External user confirmed access. Closing.",
            "Out of office confirmed configured and tested. Closing ticket.",
            "Room device signed in with shared room account. Teams join tested and working. Closing ticket."
        )
        Duration = @(0.25, 0.25, 0.25, 0.5)
        BillableOption = "Billable"
        Priority = "Priority 4 - Low Priority"
    }
)

# ============================================================
# HELPER FUNCTIONS
# ============================================================

function Get-WeightedRandom {
    param([array]$Items)
    $totalWeight = 0
    foreach ($i in $Items) { $totalWeight += $i.Weight }
    $roll = Get-Random -Minimum 1 -Maximum ($totalWeight + 1)
    $cumulative = 0
    foreach ($item in $Items) {
        $cumulative += $item.Weight
        if ($roll -le $cumulative) { return $item }
    }
    return $Items[-1]
}

function Get-WorkingDays {
    param([int]$Year, [int]$Month)
    $days = @()
    $daysInMonth = [DateTime]::DaysInMonth($Year, $Month)
    for ($d = 1; $d -le $daysInMonth; $d++) {
        $date = [DateTime]::new($Year, $Month, $d)
        if ($date.DayOfWeek -ne 'Saturday' -and $date.DayOfWeek -ne 'Sunday') {
            $days += $date
        }
    }
    return $days
}

function Get-RandomBusinessTime {
    param([DateTime]$Date)
    # Business hours: 8:30am - 5:00pm AEST, weighted toward morning
    $hourWeights = @(
        @{ Hour = 8;  Minute = 30; Weight = 5  },
        @{ Hour = 9;  Minute = 0;  Weight = 20 },
        @{ Hour = 10; Minute = 0;  Weight = 15 },
        @{ Hour = 11; Minute = 0;  Weight = 10 },
        @{ Hour = 13; Minute = 0;  Weight = 10 },
        @{ Hour = 14; Minute = 0;  Weight = 15 },
        @{ Hour = 15; Minute = 0;  Weight = 10 },
        @{ Hour = 16; Minute = 0;  Weight = 5  }
    )
    $slot = Get-WeightedRandom -Items $hourWeights
    $minute = $slot.Minute + (Get-Random -Minimum 0 -Maximum 55)
    return $Date.AddHours($slot.Hour).AddMinutes($minute)
}

function Get-CWMMemberByIdentifier {
    param([string]$Identifier)
    $encoded = [Uri]::EscapeDataString("identifier='$Identifier'")
    $uri = "$script:CWMBaseUrl/system/members?conditions=$encoded"
    try {
        $r = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($r.Count -gt 0) { return $r[0] }
        return $null
    } catch {
        Write-Log "Failed to get member '$Identifier': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMCompanyContacts {
    param([int]$CompanyId)
    $encoded = [Uri]::EscapeDataString("company/id=$CompanyId and inactiveFlag=false")
    $uri = "$script:CWMBaseUrl/company/contacts?conditions=$encoded&pageSize=50"
    try {
        $r = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        return $r
    } catch {
        Write-Log "Failed to get contacts for company $CompanyId`: $_" -Level "ERROR"
        return @()
    }
}

function Get-CWMBoardStatus {
    param([int]$BoardId, [string]$StatusName)
    $encoded = [Uri]::EscapeDataString("name='$StatusName'")
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/statuses?conditions=$encoded"
    try {
        $r = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($r.Count -gt 0) { return $r[0] }
        return $null
    } catch {
        Write-Log "Failed to get board status '$StatusName': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMPriority {
    param([string]$PriorityName)
    $encoded = [Uri]::EscapeDataString("name='$PriorityName'")
    $uri = "$script:CWMBaseUrl/service/priorities?conditions=$encoded"
    try {
        $r = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($r.Count -gt 0) { return $r[0] }
        return $null
    } catch {
        Write-Log "Failed to get priority '$PriorityName': $_" -Level "ERROR"
        return $null
    }
}

function New-CWMTicket {
    param([hashtable]$TicketData)
    $uri = "$script:CWMBaseUrl/service/tickets"
    $body = $TicketData | ConvertTo-Json -Depth 10
    try {
        $r = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Log "  Created ticket #$($r.id): $($r.summary)" -Level "SUCCESS"
        return $r
    } catch {
        Write-Log "  Failed to create ticket '$($TicketData.summary)': $_" -Level "ERROR"
        return $null
    }
}

function New-CWMTicketNote {
    param([int]$TicketId, [hashtable]$NoteData)
    $uri = "$script:CWMBaseUrl/service/tickets/$TicketId/notes"
    $body = $NoteData | ConvertTo-Json -Depth 10
    try {
        $r = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        return $r
    } catch {
        Write-Log "    Failed to add note to ticket $TicketId`: $_" -Level "ERROR"
        return $null
    }
}

function New-CWMTimeEntry {
    param([hashtable]$EntryData)
    $uri = "$script:CWMBaseUrl/time/entries"
    $body = $EntryData | ConvertTo-Json -Depth 10
    try {
        $r = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Log "    Time entry created: $($EntryData.actualHours)h" -Level "SUCCESS"
        return $r
    } catch {
        Write-Log "    Failed to create time entry: $_" -Level "ERROR"
        return $null
    }
}

function Update-CWMTicketStatus {
    param([int]$TicketId, [int]$StatusId)
    $uri = "$script:CWMBaseUrl/service/tickets/$TicketId"
    $op = @{ op = "replace"; path = "/status/id"; value = $StatusId } | ConvertTo-Json -Depth 5 -Compress
    $body = "[$op]"
    try {
        Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Patch -Body $body | Out-Null
    } catch {
        Write-Log "    Failed to update ticket $TicketId status: $_" -Level "WARNING"
    }
}

# ============================================================
# MAIN
# ============================================================

function Generate-Tickets {
    $logPath = Initialize-Logging -LogName "CWM-Generate-Tickets"

    Write-Log "=========================================" -Level "INFO"
    Write-Log "CWM Ticket Generator" -Level "INFO"
    Write-Log "Company  : $CompanyIdentifier" -Level "INFO"
    Write-Log "Period   : $Month/$Year" -Level "INFO"
    Write-Log "Per Day  : $TicketsPerDay" -Level "INFO"
    Write-Log "Log      : $logPath" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"

    Connect-CWM

    # -- Look up company --
    $company = Get-CWMCompanyByIdentifier -Identifier $CompanyIdentifier
    if (-not $company) {
        Write-Log "Company '$CompanyIdentifier' not found in CWM. Aborting." -Level "ERROR"
        return
    }
    Write-Log "Company found: $($company.name) (ID: $($company.id))" -Level "SUCCESS"

    # -- Look up contacts for company --
    $contacts = Get-CWMCompanyContacts -CompanyId $company.id
    if ($contacts.Count -eq 0) {
        Write-Log "No contacts found for company $($company.name). Aborting." -Level "ERROR"
        return
    }
    Write-Log "Contacts loaded: $($contacts.Count)" -Level "SUCCESS"

    # -- Look up Glenn Farnsworth (tech member) --
    $glennMember = Get-CWMMemberByIdentifier -Identifier "GFarnsworth"
    if (-not $glennMember) {
        # Try by name if identifier is different
        $encoded = [Uri]::EscapeDataString("firstName='Glenn' and lastName='Farnsworth'")
        $r = Invoke-RestMethod -Uri "$script:CWMBaseUrl/system/members?conditions=$encoded" -Headers $script:CWMHeaders -Method Get
        if ($r.Count -gt 0) { $glennMember = $r[0] }
    }
    if (-not $glennMember) {
        Write-Log "Glenn Farnsworth member not found. Aborting." -Level "ERROR"
        return
    }
    Write-Log "Technician: $($glennMember.firstName) $($glennMember.lastName) (ID: $($glennMember.id))" -Level "SUCCESS"

    # -- Look up Service Desk board --
    $board = Get-CWMServiceBoard -BoardName "Service Desk"
    if (-not $board) {
        Write-Log "Service Desk board not found. Aborting." -Level "ERROR"
        return
    }
    Write-Log "Board: $($board.name) (ID: $($board.id))" -Level "SUCCESS"

    # -- Pre-load statuses --
    $statusNew    = Get-CWMBoardStatus -BoardId $board.id -StatusName "New"
    $statusClosed = Get-CWMBoardStatus -BoardId $board.id -StatusName "Closed"

    # -- Pre-load priorities --
    $priorityCache = @{}

    # -- Get working days --
    $workingDays = Get-WorkingDays -Year $Year -Month $Month
    Write-Log "Working days in $Month/$Year`: $($workingDays.Count)" -Level "INFO"
    Write-Log "" -Level "INFO"

    $totalTickets = 0

    foreach ($day in $workingDays) {
        Write-Log "--- $($day.ToString('dddd dd MMM yyyy')) ---" -Level "INFO"

        for ($t = 1; $t -le $TicketsPerDay; $t++) {

            # Pick scenario + content
            $scenario     = Get-WeightedRandom -Items $Scenarios
            $summaryIdx   = Get-Random -Minimum 0 -Maximum $scenario.Summaries.Count
            $summary      = $scenario.Summaries[$summaryIdx]
            $opener       = $scenario.ContactOpeners[$summaryIdx]
            $techReply    = $scenario.TechReplies[$summaryIdx]
            $followUp     = $scenario.ContactFollowUps[$summaryIdx]
            $resolution   = $scenario.TechResolutions[$summaryIdx]
            $duration     = $scenario.Duration[$summaryIdx]

            # Pick a random contact
            $contact = $contacts | Get-Random
            $contactName = "$($contact.firstName) $($contact.lastName)"

            # Substitute contact name into tech reply
            $techReply = $techReply -replace '\{contact\}', $contact.firstName

            # Randomise ticket time through the day
            $ticketTime   = Get-RandomBusinessTime -Date $day
            $replyOffset  = Get-Random -Minimum 10 -Maximum 40   # tech replies within 10-40 min
            $followOffset = Get-Random -Minimum 5  -Maximum 20   # contact follows up 5-20 min later
            $closeOffset  = Get-Random -Minimum 5  -Maximum 30   # resolution shortly after

            $timeOpener     = $ticketTime
            $timeTechReply  = $ticketTime.AddMinutes($replyOffset)
            $timeFollowUp   = $timeTechReply.AddMinutes($followOffset)
            $timeResolution = $timeFollowUp.AddMinutes($closeOffset)

            # Resolve priority
            $priorityName = $scenario.Priority
            if (-not $priorityCache.ContainsKey($priorityName)) {
                $p = Get-CWMPriority -PriorityName $priorityName
                $priorityCache[$priorityName] = $p
            }
            $priority = $priorityCache[$priorityName]

            # -- Build ticket payload --
            $ticketData = @{
                summary = $summary
                company = @{ id = $company.id }
                contact = @{ id = $contact.id }
                board   = @{ id = $board.id }
            }
            if ($statusNew)  { $ticketData.status   = @{ id = $statusNew.id } }
            if ($priority)   { $ticketData.priority  = @{ id = $priority.id } }

            $ticket = New-CWMTicket -TicketData $ticketData
            if (-not $ticket) { continue }
            $ticketId = $ticket.id

            # -- Note 1: Client opens the ticket --
            New-CWMTicketNote -TicketId $ticketId -NoteData @{
                text                  = $opener
                detailDescriptionFlag = $true
                externalFlag          = $true
                contact               = @{ id = $contact.id }
            } | Out-Null

            # -- Note 2: Tech responds --
            New-CWMTicketNote -TicketId $ticketId -NoteData @{
                text                 = $techReply
                internalAnalysisFlag = $true
                member               = @{ id = $glennMember.id }
            } | Out-Null

            # -- Note 3: Client follow-up --
            New-CWMTicketNote -TicketId $ticketId -NoteData @{
                text                  = $followUp
                detailDescriptionFlag = $true
                externalFlag          = $true
                contact               = @{ id = $contact.id }
            } | Out-Null

            # -- Note 4: Tech resolution --
            New-CWMTicketNote -TicketId $ticketId -NoteData @{
                text           = $resolution
                resolutionFlag = $true
                member         = @{ id = $glennMember.id }
            } | Out-Null

            # -- Time entry (Glenn Farnsworth) --
            $timeStart = $timeTechReply
            $timeEnd   = $timeStart.AddHours($duration)

            New-CWMTimeEntry -EntryData @{
                chargeToId     = $ticketId
                chargeToType   = "ServiceTicket"
                member         = @{ id = $glennMember.id }
                company        = @{ id = $company.id }
                timeStart      = $timeStart.ToString("yyyy-MM-ddTHH:mm:ssZ")
                timeEnd        = $timeEnd.ToString("yyyy-MM-ddTHH:mm:ssZ")
                actualHours    = $duration
                billableOption = $scenario.BillableOption
                notes          = "Investigated and resolved: $summary"
            } | Out-Null

            # -- Close the ticket --
            if ($statusClosed) {
                Update-CWMTicketStatus -TicketId $ticketId -StatusId $statusClosed.id
            }

            $totalTickets++
        }

        Write-Log "" -Level "INFO"
    }

    Write-Log "=========================================" -Level "SUCCESS"
    Write-Log "Generation complete. Total tickets: $totalTickets" -Level "SUCCESS"
    Write-LogSummary
}

Generate-Tickets
