param(
    [int]$Days = 30,
    [string]$CompanyIdentifier = ""
)

. .\CWM-Common.ps1

function Get-TicketsWithTime {
    $logPath  = Initialize-Logging -LogName "CWM-Get-Tickets-TimeEntries"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $CsvPath  = ".data\output\CWM-Tickets-TimeEntries_$timestamp.csv"

    Write-Log "=========================================" -Level "INFO"
    Write-Log "CWM Ticket + Time Entry Report" -Level "INFO"
    Write-Log "Days back      : $Days" -Level "INFO"
    Write-Log "Company filter : $(if ($CompanyIdentifier) { $CompanyIdentifier } else { 'All' })" -Level "INFO"
    Write-Log "Output CSV     : $CsvPath" -Level "INFO"
    Write-Log "Log            : $logPath" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"

    Connect-CWM

    $fromDate = (Get-Date).AddDays(-$Days).ToString("yyyy-MM-ddT00:00:00Z")

    # ---- Build ticket conditions ----
    $conditions = "dateEntered>=[{0}]" -f $fromDate
    if ($CompanyIdentifier) {
        $company = Get-CWMCompanyByIdentifier -Identifier $CompanyIdentifier
        if (-not $company) {
            Write-Log "Company '$CompanyIdentifier' not found. Aborting." -Level "ERROR"
            return
        }
        $conditions += " and company/id=$($company.id)"
        Write-Log "Company resolved: $($company.name) (ID: $($company.id))" -Level "SUCCESS"
    }

    # ---- Paginate through all tickets ----
    Write-Log "Fetching tickets since $fromDate..." -Level "INFO"
    $allTickets = @()
    $page       = 1
    $pageSize   = 100
    $encoded    = [Uri]::EscapeDataString($conditions)

    do {
        $uri      = "$script:CWMBaseUrl/service/tickets?conditions=$encoded&page=$page&pageSize=$pageSize&fields=id,summary,company,contact,board,type,subType,item,status,priority,dateEntered,closedDate,actualHours,resources"
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) {
            $allTickets += $response
            Write-Log "  Page $page — $($response.Count) tickets" -Level "INFO"
            $page++
        }
    } while ($response.Count -eq $pageSize)

    Write-Log "Total tickets found: $($allTickets.Count)" -Level "SUCCESS"

    if ($allTickets.Count -eq 0) {
        Write-Log "No tickets found for the specified criteria." -Level "WARNING"
        return
    }

    # ---- Fetch all time entries for the date range in one paginated call ----
    Write-Log "" -Level "INFO"
    Write-Log "Fetching time entries since $fromDate..." -Level "INFO"

    $timeConditions = "chargeToType='ServiceTicket' and timeStart>=[{0}]" -f $fromDate
    if ($CompanyIdentifier -and $company) {
        $timeConditions += " and company/id=$($company.id)"
    }
    $encodedTime = [Uri]::EscapeDataString($timeConditions)

    $allTimeEntries = @()
    $tPage          = 1

    do {
        $tUri      = "$script:CWMBaseUrl/time/entries?conditions=$encodedTime&page=$tPage&pageSize=1000&fields=id,chargeToId,actualHours,billableOption,member,workType,timeStart,timeEnd,notes"
        $tResponse = Invoke-RestMethod -Uri $tUri -Headers $script:CWMHeaders -Method Get
        if ($tResponse.Count -gt 0) {
            $allTimeEntries += $tResponse
            Write-Log "  Page $tPage — $($tResponse.Count) time entries" -Level "INFO"
            $tPage++
        }
    } while ($tResponse.Count -eq 1000)

    Write-Log "Total time entries found: $($allTimeEntries.Count)" -Level "SUCCESS"

    # ---- Group time entries by ticket ID ----
    $timeByTicket = @{}
    foreach ($entry in $allTimeEntries) {
        $tid = $entry.chargeToId
        if (-not $timeByTicket.ContainsKey($tid)) {
            $timeByTicket[$tid] = @()
        }
        $timeByTicket[$tid] += $entry
    }

    # ---- Build CSV rows ----
    Write-Log "" -Level "INFO"
    Write-Log "Building report..." -Level "INFO"

    $rows = foreach ($ticket in $allTickets | Sort-Object id) {
        $tid     = $ticket.id
        $entries = if ($timeByTicket.ContainsKey($tid)) { $timeByTicket[$tid] } else { @() }

        $totalHours     = [math]::Round(($entries | Measure-Object -Property actualHours -Sum).Sum, 2)
        $billableHours  = [math]::Round(($entries | Where-Object { $_.billableOption -eq 'Billable' } | Measure-Object -Property actualHours -Sum).Sum, 2)
        $noBillHours    = [math]::Round(($entries | Where-Object { $_.billableOption -ne 'Billable' } | Measure-Object -Property actualHours -Sum).Sum, 2)
        $entryCount     = $entries.Count

        $members = ($entries | Where-Object { $_.member } | ForEach-Object { $_.member.name } | Select-Object -Unique) -join "; "
        $workTypes = ($entries | Where-Object { $_.workType } | ForEach-Object { $_.workType.name } | Select-Object -Unique) -join "; "

        [PSCustomObject]@{
            TicketID        = $tid
            Summary         = $ticket.summary
            Company         = if ($ticket.company)  { $ticket.company.name }  else { "" }
            Contact         = if ($ticket.contact)  { $ticket.contact.name }  else { "" }
            Board           = if ($ticket.board)    { $ticket.board.name }    else { "" }
            Type            = if ($ticket.type)     { $ticket.type.name }     else { "" }
            SubType         = if ($ticket.subType)  { $ticket.subType.name }  else { "" }
            Item            = if ($ticket.item)     { $ticket.item.name }     else { "" }
            Status          = if ($ticket.status)   { $ticket.status.name }   else { "" }
            Priority        = if ($ticket.priority) { $ticket.priority.name } else { "" }
            DateEntered     = $ticket.dateEntered
            ClosedDate      = $ticket.closedDate
            TimeEntries     = $entryCount
            TotalHours      = $totalHours
            BillableHours   = $billableHours
            NonBillableHours = $noBillHours
            Technicians     = $members
            WorkTypes       = $workTypes
        }
    }

    # ---- Export CSV ----
    $null = New-Item -ItemType Directory -Path (Split-Path $CsvPath) -Force
    $rows | Export-Csv -Path $CsvPath -NoTypeInformation
    Write-Log "CSV exported: $CsvPath" -Level "SUCCESS"

    # ---- Summary ----
    $grandTotal    = [math]::Round(($rows | Measure-Object -Property TotalHours -Sum).Sum, 2)
    $grandBillable = [math]::Round(($rows | Measure-Object -Property BillableHours -Sum).Sum, 2)
    $withTime      = ($rows | Where-Object { $_.TimeEntries -gt 0 }).Count
    $noTime        = ($rows | Where-Object { $_.TimeEntries -eq 0 }).Count

    Write-Log "" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "REPORT SUMMARY" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "Tickets total        : $($rows.Count)" -Level "INFO"
    Write-Log "  With time entries  : $withTime" -Level "INFO"
    Write-Log "  No time entries    : $noTime" -Level "INFO"
    Write-Log "Total hours          : $grandTotal h" -Level "INFO"
    Write-Log "Billable hours       : $grandBillable h" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-LogSummary
}

Get-TicketsWithTime
