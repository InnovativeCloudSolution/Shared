$TenantID="22a3df9b-ba07-4cf7-97c5-5dd4ff48c62f"
$Company = "CleanCo Queensland - Project Licensing"
# Logging in to Azure AD with Service Principal
$PartnerCenterApp = Get-AutomationPSCredential -Name 'PartnerCenterApp'
$ClientAppId = $PartnerCenterApp.UserName

$RefreshToken = Get-AutomationVariable -Name 'PartnerCenterAppRefreshToken'
$Connection = Connect-PartnerCenter -ApplicationId $ClientAppId -RefreshToken $RefreshToken -Credential $PartnerCenterApp

$UserLicenseTable = New-Object System.Data.DataTable
$UserLicenseTable.Columns.Add("User","System.String")|Out-Null
$UserLicenseTable.Columns.Add("Licenses","System.String")|Out-Null

$Users = Get-PartnerCustomerUser -CustomerId $TenantID

foreach ($User in $Users){
    $UserLicenses = Get-PartnerCustomerUserLicense -CustomerId $TenantID -UserId $User.UserId
    if($UserLicenses.Length -ge 1){
        $first = $True
        $LicenseList=""
        $ProjectUser=$false

        foreach($UserLicense in $UserLicenses){
            if($UserLicense.Name.ToLower().Contains('project') -or $UserLicense.Name.ToLower().Contains('visio')){
               if($first -ne $True){
                    $LicenseList += ", "
                }
            $first = $False
            $LicenseList+= $UserLicense.Name
            $ProjectUser=$true
            }
        }
        if($ProjectUser){
            $Row = $UserLicenseTable.NewRow()
            $Row.User = $User.DisplayName
            $Row.Licenses = $LicenseList
            $UserLicenseTable.Rows.Add($Row)
        }
    }
}

$UserLicenseHTML ="<table><tr><th>Name</th><th>Licenses</th></tr>"
foreach ($row in $UserLicenseTable.Rows)
{ 
    $UserLicenseHTML += "<tr><td>" + $row[0] + "</td><td>" + $row[1] + "</td></tr>"
}
$UserLicenseHTML += "</table>"

$LicenseUsageTable = New-Object System.Data.DataTable
$LicenseUsageTable.Columns.Add("License","System.String")|Out-Null
$LicenseUsageTable.Columns.Add("ConsumedCount","System.String")|Out-Null
$LicenseUsageTable.Columns.Add("ActiveCount","System.String")|Out-Null
$LicenseUsageTable.Columns.Add("AvailableCount","System.String")|Out-Null

$SKUs = Get-PartnerCustomerSubscribedSku -CustomerId $TenantID
foreach ($SKU in $SKUs){
    if($SKU.ProductName.ToLower().Contains('project') -or $SKU.ProductName.ToLower().Contains('visio') -or $SKU.ProductName.ToLower().Contains('power apps per app plan')){
    $Row = $LicenseUsageTable.NewRow()
    $Row.License = $SKU.ProductName
    $Row.ConsumedCount = $SKU.ConsumedUnits
    $Row.ActiveCount = $SKU.ActiveUnits
    $Row.AvailableCount = $SKU.AvailableUnits
    $LicenseUsageTable.Rows.Add($Row)
    }
}
$LicenseUsageHTML = "<h2>License Usage Information</h2><table><tr><th>License</th><th>Consumed Count</th><th>Active Count</th><th>Available Count</th></tr>"
foreach ($row in $LicenseUsageTable.Rows)
{ 
    $LicenseUsageHTML += "<tr><td>" + $row[0] + "</td><td>" + $row[1] + "</td><td>" + $row[2] + "</td><td>" + $row[3] + "</td></tr>"
}
$LicenseUsageHTML += "</table>"

$reportdate = (Get-Date).AddHours(10).ToString("D")
$UserCount = $UserLicenseTable.Rows.Count

$html = "<html><body><h1>$Company</h1>"
$html += "<h3>Generated: $reportdate</h3>"
$html += "<h2>Total Licensed Users: $UserCount</h2>"
$html += $UserLicenseHTML
$html += $LicenseUsageHTML
$html += "</body></html>"

Write-Output $html