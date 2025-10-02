param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath
)

$CsvPath = "C:\Scripts\InputData\BBH-ContactImport.csv"

Clear-Host
Import-Module 'ConnectWiseManageAPI'

Write-Host "`nMangano IT - Bulk Import Contacts to ConnectWise Manage" -ForegroundColor Yellow
Write-Host "Version: " -ForegroundColor Yellow -NoNewLine; Write-Host "1.0"

$CsvHeaders = 'FirstName', 'LastName', 'Email', 'JobTitle', 'DirectPhone', 'MobilePhone', 'SiteId', 'CompanyName', 'ContactTypeName'
$CWMServer = 'https://api-aus.myconnectwise.net'
$CWMCompany = 'mit'
$CWMPublicKey = 'mBemoCno7IwHElgT'
$CWMPrivateKey = 'BmC0AR9dN7eJFouT'
$CWMClientId = '1208536d-40b8-4fc0-8bf3-b4955dd9d3b7'

function Connect-CWMServer {
    param (
        [string]$Server,
        [string]$Company,
        [string]$PublicKey,
        [string]$PrivateKey,
        [string]$ClientId
    )
    try {
        Write-Host "`nConnecting to the ConnectWise Manage server..."
        Connect-CWM -Server $Server -Company $Company -pubkey $PublicKey -privatekey $PrivateKey -clientId $ClientId
        Write-Host "Connection successful."
        return $true
    }
    catch {
        Write-Host "Unable to connect to ConnectWise Manage. Stopping script." -ForegroundColor red
        return $false
    }
}

function Import-ContactCsv {
    param (
        [string]$Path,
        [array]$Headers
    )
    try {
        Write-Host "`nAttempting to import the CSV from " -NoNewLine
        Write-Host "$Path " -ForegroundColor blue -NoNewline
        Write-Host "now, please wait..."
        $csv = Import-Csv -Path $Path -Delimiter ',' -Header $Headers
        Write-Host "CSV imported successfully."
        return $csv
    }
    catch {
        Write-Host "CSV was not able to be imported. Stopping script." -ForegroundColor red
        return $null
    }
}

function Format-PhoneNumber {
    param ([string]$Number)
    $clean = $Number -replace '[^0-9]', ''
    if ($clean.Length -eq 9 -and $clean[0] -ne '0') {
        $clean = '0' + $clean
    }
    return $clean
}

function New-CommunicationItems {
    param (
        [string]$Email,
        [string]$DirectPhone,
        [string]$MobilePhone,
        [string]$LinkedIn,
        [string]$Facebook,
        [string]$Twitter
    )

    $items = @()

    if ($Email) {
        $items += @{
            type              = @{ id = 1; name = 'Email' }
            value             = $Email
            defaultFlag       = $true
            communicationType = 'Email'
        }
    }

    if ($DirectPhone -and $DirectPhone -ne 'null' -and $DirectPhone -ne '0') {
        $items += @{
            type              = @{ id = 2; name = 'Direct' }
            value             = $DirectPhone
            defaultFlag       = $false
            communicationType = 'Phone'
        }
    }

    if ($MobilePhone -and $MobilePhone -ne 'null' -and $MobilePhone -ne '0') {
        $items += @{
            type              = @{ id = 4; name = 'Mobile' }
            value             = $MobilePhone
            defaultFlag       = $true
            communicationType = 'Phone'
        }
    }

    # Exclude LinkedIn, Facebook, Twitter unless valid types are defined
    return $items | Where-Object { $_.value -and $_.value.Trim() -ne '' }
}

function Update-ContactSocialUrls {
    param (
        [int]$ContactId,
        [string]$LinkedIn,
        [string]$Facebook,
        [string]$Twitter
    )

    if ($LinkedIn -and $LinkedIn.Trim()) {
        Write-Host "Updating LinkedIn URL..."
        Update-CWMCompanyContact -Id $ContactId -Operation replace -Path 'linkedInUrl' -Value $LinkedIn
    }

    if ($Facebook -and $Facebook.Trim()) {
        Write-Host "Updating Facebook URL..."
        Update-CWMCompanyContact -Id $ContactId -Operation replace -Path 'facebookUrl' -Value $Facebook
    }

    if ($Twitter -and $Twitter.Trim()) {
        Write-Host "Updating Twitter URL..."
        Update-CWMCompanyContact -Id $ContactId -Operation replace -Path 'twitterUrl' -Value $Twitter
    }
}

function Find-OrCreateContact {
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$Title,
        [int]$CompanyId,
        [int]$SiteId,
        [array]$CommItems,
        [int]$ContactTypeId,
        [string]$LinkedIn,
        [string]$Facebook,
        [string]$Twitter
    )

    Write-Host "Checking for existing contact: $FirstName $LastName ($Title)..."

    $ContactCheck = Get-CWMCompanyContact -Condition "company/id = $CompanyId AND firstName = '$FirstName' AND lastName = '$LastName'" -All |
    Where-Object { $_.title -eq $Title -and $_.site.id -eq $SiteId }

    if (-not $ContactCheck -and $CommItems) {
        $primaryEmail = $CommItems | Where-Object { $_.communicationType -eq "Email" } | Select-Object -First 1
        if ($primaryEmail) {
            $ContactCheck = Get-CWMCompanyContact -Condition "company/id = $CompanyId" -All | Where-Object {
                $_.communicationItems | Where-Object { $_.value -eq $primaryEmail.value }
            } | Select-Object -First 1
            if ($ContactCheck) {
                Write-Host "Fallback match found using email [$($primaryEmail.value)]"
            }
        }
    }

    if (-not $ContactCheck) {
        Write-Host "Creating new contact for $FirstName $LastName..."
        try {
            $params = @{
                firstName          = $FirstName
                lastName           = $LastName
                title              = $Title
                company            = @{ id = $CompanyId }
                site               = @{ id = $SiteId }
                communicationItems = $CommItems
                inactiveFlag       = $false
            }
            if ($ContactTypeId) {
                $params.types = @(@{ id = $ContactTypeId })
            }
            if ($LinkedIn -and $LinkedIn.Trim()) { $params.linkedInUrl = $LinkedIn }
            if ($Facebook -and $Facebook.Trim()) { $params.facebookUrl = $Facebook }
            if ($Twitter -and $Twitter.Trim()) { $params.twitterUrl = $Twitter }

            return New-CWMCompanyContact @params
        }
        catch {
            Write-Host "ERROR creating contact: $_" -ForegroundColor Red
            return $null
        }
    }

    Write-Host "Updating existing contact for $FirstName $LastName..."
    $ContactId = $ContactCheck.id

    try {
        Update-CWMCompanyContact -Id $ContactId -Operation replace -Path 'inactiveFlag' -Value $false
        Write-Host "Set inactiveFlag to false"

        if ($ContactTypeId) {
            $existingTypeIds = @()
            if ($ContactCheck.types) {
                $existingTypeIds = $ContactCheck.types | ForEach-Object { $_.id }
            }
            if ($existingTypeIds -notcontains $ContactTypeId) {
                $allTypes = Get-CWMContactType -All
                $mergedTypes = @($existingTypeIds + $ContactTypeId | Sort-Object -Unique | ForEach-Object {
                    $type = $allTypes | Where-Object { $_.id -eq $_ }
                    if ($type) { @{ id = $type.id; name = $type.name } }
                })
                Write-Host "Adding contact type ID $ContactTypeId"
                try {
                    Update-CWMCompanyContact -Id $ContactId -Operation replace -Path 'types' -Value $mergedTypes
                }
                catch {
                    Write-Host "WARNING: Failed to update types due to internal ConnectWise error. Skipping update." -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "Contact type ID $ContactTypeId already assigned"
            }
        }

        Update-CWMCompanyContact -Id $ContactId -Operation replace -Path 'communicationItems' -Value @()
        Write-Host "Cleared existing communicationItems"

        $updatedItems = $CommItems | Where-Object { $_.value -and $_.value.Trim() -ne '' }
        foreach ($item in $updatedItems) {
            Write-Host "Setting communication item: [$($item.type.name)] $($item.value)"
        }

        Update-CWMCompanyContact -Id $ContactId -Operation replace -Path 'communicationItems' -Value $updatedItems
        Write-Host "Updated communicationItems"

        Update-ContactSocialUrls -ContactId $ContactId -LinkedIn $LinkedIn -Facebook $Facebook -Twitter $Twitter

        return $ContactCheck
    }
    catch {
        Write-Host "ERROR during Find-OrCreateContact: $_" -ForegroundColor Red
        return $null
    }
}

function main {
    if (-not (Connect-CWMServer -Server $CWMServer -Company $CWMCompany -PublicKey $CWMPublicKey -PrivateKey $CWMPrivateKey -ClientId $CWMClientId)) {
        return
    }

    $ContactsCsv = Import-ContactCsv -Path $CsvPath -Headers $CsvHeaders
    $ContactsCsv = $ContactsCsv | Where-Object { $_.FirstName -ne "FirstName" }
    if (-not $ContactsCsv) {
        return
    }

    $AllContactTypes = Get-CWMContactType -All
    $ContactsCount = 0

    foreach ($Contact in $ContactsCsv) {
        if ($Contact.FirstName -eq 'FirstName') {
            continue
        }

        $FirstName = $Contact.FirstName
        $LastName = $Contact.LastName
        $Email = $Contact.Email
        $Title = if ($Contact.JobTitle) { $Contact.JobTitle } else { " " }
        $CompanyName = $Contact.CompanyName
        $SiteId = $Contact.SiteId
        $ContactTypeName = $Contact.ContactTypeName

        $LinkedIn = $Contact.LinkedIn
        $Facebook = $Contact.Facebook
        $Twitter = $Contact.Twitter

        $ContactTypeId = $null
        if ($ContactTypeName) {
            $typeMatch = $AllContactTypes | Where-Object { $_.description -eq $ContactTypeName -or $_.name -eq $ContactTypeName } | Select-Object -First 1
            if ($typeMatch) {
                $ContactTypeId = $typeMatch.id
            }
            else {
                Write-Host "Contact type '$ContactTypeName' not found. Skipping $FirstName $LastName." -ForegroundColor Red
                continue
            }
        }

        $DirectPhone = Format-PhoneNumber -Number $Contact.DirectPhone
        $MobilePhone = Format-PhoneNumber -Number $Contact.MobilePhone

        $Company = Get-CWMCompany -Condition "name LIKE '$CompanyName' AND deletedFlag=false"
        if (-not $Company) {
            Write-Host "Company '$CompanyName' not found. Skipping $FirstName $LastName." -ForegroundColor red
            continue
        }

        $CommItems = New-CommunicationItems -Email $Email -DirectPhone $DirectPhone -MobilePhone $MobilePhone -LinkedIn $LinkedIn -Facebook $Facebook -Twitter $Twitter
        $CommItems = $CommItems | Where-Object { $_.value -and $_.value.Trim() -ne '' }

        Write-Host "`nProcessing contact: " -NoNewline
        Write-Host "$FirstName $LastName ($Title)" -ForegroundColor blue -NoNewline
        Write-Host "..."

        $Result = Find-OrCreateContact `
            -FirstName $FirstName `
            -LastName $LastName `
            -Title $Title `
            -CompanyId $Company.id `
            -SiteId $SiteId `
            -CommItems $CommItems `
            -ContactTypeId $ContactTypeId `
            -LinkedIn $LinkedIn `
            -Facebook $Facebook `
            -Twitter $Twitter

        if ($Result) {
            Write-Host "Success for " -NoNewline
            Write-Host "$FirstName $LastName ($Title)" -ForegroundColor blue
            $ContactsCount++
        }
        else {
            Write-Host "Failed to process " -ForegroundColor red -NoNewline
            Write-Host "$FirstName $LastName ($Title)" -ForegroundColor blue
        }
    }

    Disconnect-CWM
    Write-Host "`nContacts imported. Number of contacts created/updated: " -NoNewline
    Write-Host "$ContactsCount" -ForegroundColor blue
}

main