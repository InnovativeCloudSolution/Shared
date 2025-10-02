param(
    [string]$startInput,
    [string]$endInput
)

Set-StrictMode -Off

$CWMClientId = Get-AutomationVariable -Name 'clientId'
$CWMPublicKey = Get-AutomationVariable -Name 'PublicKey'
$CWMPrivateKey = Get-AutomationVariable -Name 'PrivateKey'
$CWMCompanyId = Get-AutomationVariable -Name 'CWManageCompanyId'
$CWMUrl = Get-AutomationVariable -Name 'CWManageUrl'

$CWMCredentials = "$($CWMCompanyId+"+"+$CWMPublicKey):$($CWMPrivateKey)"
$CWMEncodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($CWMCredentials))
$CWMAuthentication = "Basic $CWMEncodedCredentials"

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module -Name ImportExcel

try {
    $startDate = [datetime]::ParseExact($startInput, "dd/MM/yyyy", $null)
    $endDate = [datetime]::ParseExact($endInput, "dd/MM/yyyy", $null).AddDays(1)
}
catch {
    Write-Error "Invalid date format. Use dd/MM/yyyy."
    exit
}

$timestamp = (Get-Date -Format "yyyy-MM-dd")
$excelFilePath = "C:\Scripts\OutputData\$timestamp-MIT-Reports-TimeExpenseEntries.xlsx"

if (Test-Path $excelFilePath) {
    try {
        Remove-Item $excelFilePath -ErrorAction Stop
    }
    catch {
        Write-Warning "Unable to remove $excelFilePath. Please close any process using it and try again."
    }
}

$startUTC = $startDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$endUTC = $endDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
$dateCondition = "timeStart >= [$startUTC] AND timeStart <= [$endUTC]"

function Convert-DynamicProperties {
    param (
        [Parameter(Mandatory = $true)]
        [Object]$JsonObject,

        [Parameter(Mandatory = $false)]
        [bool]$isExpense = $false
    )

    function Convert-ToAEST {
        param (
            [Parameter(Mandatory = $true)]
            [datetime]$UtcDateTime
        )
        $timezone = [System.TimeZoneInfo]::FindSystemTimeZoneById("AUS Eastern Standard Time")
        return [System.TimeZoneInfo]::ConvertTimeFromUtc($UtcDateTime, $timezone)
    }

    $result = @{}

    foreach ($property in $JsonObject.PSObject.Properties) {
        $value = $property.Value

        if ($property.Name -eq "ticket" -and $value -is [psobject] -and $value.PSObject.Properties.Name -contains "id") {
            $result["TicketID"] = $value.id
        }
        elseif ($isExpense -and $property.Name -eq "notes") {
            $result["Reason"] = $value
        }
        elseif ($isExpense -and $property.Name -eq "_info" -and $value -is [psobject] -and $value.PSObject.Properties.Name -contains "updatedBy") {
            $approver = $value.updatedBy.ToUpper()
            if ($approver -eq "JANITAJONES") {
                $result["Approver"] = $approver
            } else {
                $result["Approver"] = ""
            }
        }
        elseif ($value -is [array]) {
            $result[$property.Name] = ($value | Where-Object { $_.PSObject.Properties.Name -contains "name" } | ForEach-Object { $_.name }) -join ', '
        }
        elseif ($value -is [psobject] -and $value.PSObject.Properties.Name -contains "name") {
            $result[$property.Name] = $value.name
        }
        elseif ($property.Name -in @("timeStart", "timeEnd", "date", "dateEntered") -and $value) {
            try {
                $utcTime = [datetime]$value
                $aestTime = Convert-ToAEST -UtcDateTime $utcTime
                $result[$property.Name] = $aestTime.ToString("yyyy-MM-dd HH:mm:ss")
            } catch {
                $result[$property.Name] = $value
            }
        }
        else {
            $result[$property.Name] = $value
        }
    }

    return [PSCustomObject]$result
}

function Get-CWMTimeEntry {
    param ([string]$Condition)

    $headers = @{
        "clientId" = "$CWMClientId"
        "Authorization" = "$CWMAuthentication"
    }

    $page = 1
    $pageSize = 1000
    $allEntries = @()

    do {
        $apiString = "$CWMUrl/time/entries?pageSize=$pageSize&page=$page&conditions=$Condition"

        try {
            $response = Invoke-RestMethod -Uri $apiString -Method 'GET' -Headers $headers
        } catch {
            Write-Error "Failed to retrieve time entries: $_"
            return @()
        }

        if ($response.Count -gt 0) {
            $allEntries += $response
            $page++
        } else {
            break
        }

    } while ($response.Count -eq $pageSize)

    return $allEntries
}

function Get-CWMExpenseEntry {
    Write-Host "Fetching Expense Entries"

    $headers = @{
        "clientId" = "$CWMClientId"
        "Authorization" = "$CWMAuthentication"
    }

    $conditions = "date >= [$startUTC] AND date <= [$endUTC]"
    $apiString = "$CWMUrl/expense/entries?pageSize=1000&conditions=$conditions"

    try {
        $response = Invoke-RestMethod -Uri $apiString -Method 'GET' -Headers $headers
    } catch {
        Write-Error "Failed to fetch expenses: $_"
        return @()
    }

    return $response | ForEach-Object { Convert-DynamicProperties -JsonObject $_ -isExpense $true }
}

function Reorder-Fields {
    param(
        [Parameter(Mandatory = $true)]
        [array]$entries,
        [array]$preferredOrder
    )

    return $entries | ForEach-Object {
        $ordered = [ordered]@{}

        foreach ($field in $preferredOrder) {
            if ($_.PSObject.Properties.Name -contains $field) {
                $ordered[$field] = $_.$field
            }
        }

        foreach ($prop in $_.PSObject.Properties.Name) {
            if (-not $ordered.Contains($prop)) {
                $ordered[$prop] = $_.$prop
            }
        }

        [PSCustomObject]$ordered
    }
}

$timeEntryHeaders = @(
    "company", "timeStart", "member", "timeEnd", "timesheet", "workType", "status", "id", "chargeToType", "actualHours"
)

$expenseEntryHeaders = @(
    "type", "invoiceAmount", "company", "Reason", "TicketID", "amount", "status", "member", "agreement", "billAmount", "classification"
)

$WorkTypes = @(
    "3 - After Hours",
    "930 - Annual Leave",
    "931 - Sick/Carers Leave (Personal)",
    "932 - Public Holiday",
    "933 - M Day",
    "934 - Leave Without Pay",
    "935 - Compassionate Leave",
    "936 - Long Service Leave",
    "937 - Self Development Leave",
    "938 - Time off in lieu"
)

# --- Export Master Time Entries ---
$allEntries = Get-CWMTimeEntry -Condition $dateCondition
$processedAllEntries = $allEntries | ForEach-Object { Convert-DynamicProperties -JsonObject $_ }
$processedAllEntries = Reorder-Fields -entries $processedAllEntries -preferredOrder $timeEntryHeaders
$processedAllEntries | Export-Excel -Path $excelFilePath -WorksheetName "Master" -AutoSize

# --- Export Expenses if any ---
$ExpenseEntries = Get-CWMExpenseEntry
if ($ExpenseEntries.Count -gt 0) {
    $ExpenseEntries = Reorder-Fields -entries $ExpenseEntries -preferredOrder $expenseEntryHeaders
    $ExpenseEntries | Export-Excel -Path $excelFilePath -WorksheetName "Expenses" -AutoSize -Append
}

# --- Export Specific WorkTypes ---
foreach ($workType in $WorkTypes) {
    $workTypeCondition = "workType/name='$workType' AND $dateCondition"
    $entries = Get-CWMTimeEntry -Condition $workTypeCondition
    $sheetName = ($workType -replace '^\d+\s*-\s*', '').Trim()

    if ($entries -and $entries.Count -gt 0) {
        $processedEntries = $entries | ForEach-Object { Convert-DynamicProperties -JsonObject $_ }
        $processedEntries = Reorder-Fields -entries $processedEntries -preferredOrder $timeEntryHeaders
        $processedEntries | Export-Excel -Path $excelFilePath -WorksheetName $sheetName -AutoSize -Append
    } else {
        [PSCustomObject]@{ Note = "No data available for '$workType' in selected date range" } |
        Export-Excel -Path $excelFilePath -WorksheetName $sheetName -AutoSize -Append
    }
}

Write-Host "Excel report created: $excelFilePath"
