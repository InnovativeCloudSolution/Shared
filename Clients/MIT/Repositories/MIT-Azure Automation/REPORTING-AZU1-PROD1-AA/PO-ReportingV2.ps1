param(
    [string]$CompanyIdentifier,
    [string]$PONumber, 
    [string]$AgreementType,
    [string]$ReportType,
    [string]$ReportSpan
)

# Turn off strict mode to allow for more flexible variable usage
Set-StrictMode -Off

# Setup API variables
$CWMClientId = Get-AutomationVariable -Name 'clientId'
$CWMPublicKey = Get-AutomationVariable -Name 'PublicKey'
$CWMPrivateKey = Get-AutomationVariable -Name 'PrivateKey'
$CWMCompany = Get-AutomationVariable -Name 'CWManageCompanyId'
$CWMUrl = Get-AutomationVariable -Name 'CWManageUrl'
$CWMCredentials = "$($CWMCompany+"+"+$CWMPublicKey):$($CWMPrivateKey)"
$CWMEncodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($CWMCredentials))
$CWMAuthentication = "Basic $CWMEncodedCredentials"

# Create a credential object used in TimeEntryReport Function
$Connection = @{
    Server     = $CWMUrl
    Company    = $CWMCompany
    pubkey     = $CWMPublicKey
    privatekey = $CWMPrivateKey
    clientId   = $CWMClientId
}

# Construct API Call string as $Conditions
$PONumber = $PONumber -replace '\s', ''
$Conditions = 'company/identifier like "%' + $CompanyIdentifier + '%" '

# Add conditions based on AgreementType and ReportSpan
if ($AgreementType -eq 'Managed Services') {
    $Conditions += "and agreement/name like '%Managed Services%' "
}
elseif ($AgreementType -eq "PO") {
    $Conditions += 'and agreement/name like "*' + $PONumber + '*" '
}

# Add conditions based on ReportSpan for Managed Services or All
if ($AgreementType -eq 'Managed Services' -or $AgreementType -eq 'All') {
    # Add conditions based on ReportSpan duration
    if ($ReportSpan -eq '1 Month' -or $ReportSpan -eq '3 Months' -or $ReportSpan -eq '6 Months') {
        $ReportStart = (Get-Date).AddMonths( - [int]$ReportSpan.Split()[0]).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $Conditions += "and lastUpdated >= [$ReportStart]"
    }
    else {
        $ReportStart = (Get-Date).AddMonths(-6).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $Conditions += "and lastUpdated >= [$ReportStart]"
    }
}

# Initialize an array to store the results
$Results = @()

# Function to generate Time Entry report
function TimeEntryReport() {
    Write-Host "Running TimeEntryReport"
    Write-Host $Conditions

    # Connect to Manage server
    Connect-CWM @Connection

    # Set headers for API call
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("clientId", "$CWMClientId")
    $headers.Add("Authorization", "$CWMAuthentication")

    # Construct API call string
    $apiString = "$CWMUrl/v4_6_release/apis/3.0/time/entries" + "?pageSize=1000&conditions=$Conditions"
    Write-Host $apiString

    # Invoke API call and get response
    $response = Invoke-RestMethod $apiString -Method 'GET' -Headers $headers

    # Process each Time Entry and create objects
    $Results = @()
    foreach ($TimeEntry in $response) {
        # Process Time Entry details
        $Cost = $TimeEntry.actualHours * $TimeEntry.hourlyRate
        if (![string]::IsNullOrEmpty($TimeEntry.notes)) {
            if ($TimeEntry.notes[0] -eq '-') {
                $TimeEntry.notes = $TimeEntry.notes.substring(1)
            }
            $TimeEntry.notes = $TimeEntry.notes -replace '\n', ' '
            if ($TimeEntry.notes.length -gt 32000) {
                $TimeEntry.notes = $TimeEntry.notes.substring(0, 32000)
            }
        }

        # Parse the date string and add 10 hours
        $timeDate = (Get-Date $TimeEntry.timeStart).AddHours(10).ToString('yyyy-MM-dd')

        $Ticket = Get-CWMTicket -id $TimeEntry.chargeToId
        $Type = $Ticket.type.name
        $Subtype = $Ticket.subType.name
        $Item = $Ticket.item.name

        $Results += [PSCustomObject]@{
            # Time Entry properties
            TimeID       = $TimeEntry.id
            TicketNumber = $TimeEntry.chargeToId
            Summary      = $TimeEntry.ticket.Summary
            Technician   = $TimeEntry.member.name
            DateTime     = $timeDate
            hoursBilled  = $TimeEntry.hoursBilled
            Cost         = $Cost
            Notes        = $TimeEntry.notes
            Agreement    = $TimeEntry.agreement.name
            Type         = $Type
            subType      = $Subtype
            item         = $Item
            status       = $Ticket.status.name
        }
    }

    # Disconnect from ConnectWise Manage
    Disconnect-CWM

    return $Results
}

# Function to generate Expense Entry report
function ExpenseEntryReport() {
    Write-Host "Running ExpenseEntryReport"
    Write-Host $Conditions

    # Connect to Manage server
    Connect-CWM @Connection

    # Set headers for API call
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("clientId", "$CWMClientId")
    $headers.Add("Authorization", "$CWMAuthentication")

    # Construct API call string
    $apiString = "$CWMUrl/v4_6_release/apis/3.0/expense/entries" + "?pageSize=1000&conditions=$Conditions"
    Write-Host $apiString

    # Invoke API call and get response
    $response = Invoke-RestMethod $apiString -Method 'GET' -Headers $headers

    # Process each Expense Entry and create objects
    $Results = @()

    foreach ($expense in $response) {
        # Process Expense Entry details
        if ($null -eq $expense.company -or 
            -not $expense.agreement.id -is [int] -or 
            -not $expense.ticket.id -is [int]) {
            # Handle errors and continue to the next iteration
            Write-Error "Invalid expense entry:" $expense
            Continue
        }

        # Retrieve Agreement and Ticket details
        $expenseAgreement = Get-CWMAgreement -id $expense.agreement.id
        $Ticket = Get-CWMTicket -id $expense.ticket.id

        # Parse and adjust the date
        $timeDate = (Get-Date $expense.date).AddHours(10).ToString('yyyy-MM-dd')

        # Create an object with selected properties
        $Results += [PSCustomObject]@{
            # Expense Entry properties
            ExpenseID     = $expense.id
            Technician    = $expense.member.name
            DateTime      = $timeDate
            ExpenseType   = $expense.type.name
            Bill          = $expense.billableOption
            Notes         = $expense.notes
            WIPAmt        = $expense.billAmount
            InvAmt        = $expense.invoiceAmount
            taxCode       = $expenseAgreement.taxCode.name
            ExpenseStatus = $expense.status
            Location      = $Ticket.location.name
            BusinessUnit  = $Ticket.department.name
            TicketNumber  = $expense.chargeToId
            Summary       = $expense.ticket.Summary
            TicketStatus  = $Ticket.status.name
            siteName      = $Ticket.siteName
            Agreement     = $expense.agreement.name
            AgreementType = $expenseAgreement.type.name
        }
    }

    # Disconnect from ConnectWise Manage
    Disconnect-CWM

    return $Results
}

# Conditionally execute TimeEntryReport or ExpenseEntryReport based on ReportType
if ($ReportType -eq 'Time') {
    $Output = @()
    $Output = TimeEntryReport
    $Count = "Found " + $Output.count + " TimeEntrys"
}

if ($ReportType -eq 'Expense') {
    $Output = @()
    $Output = ExpenseEntryReport
    $Count = "Found " + $Output.count + " Expenses"
}

Write-Host $Count

## SEND DETAILS BACK TO FLOW ##
Write-Output $Output | ConvertTo-Json