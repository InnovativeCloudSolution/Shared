param(
    [Parameter(Mandatory=$true)]
        [string] 
        $TenantID,
    [Parameter(Mandatory=$true)]
        [string] 
        $Company
)

# Get Azure Run As Connection Name
#$connectionName = "AzureRunAsConnection"
# Get the Service Principal connection details for the Connection name
#$servicePrincipalConnection = Get-AutomationConnection -Name $connectionName         

# Logging in to Azure AD with Service Principal
$PartnerCenterApp = Get-AutomationPSCredential -Name 'PartnerCenterApp'
$ClientAppId = $PartnerCenterApp.UserName
$ClientSecret = $PartnerCenterApp.Password
$RefreshToken = Get-AutomationVariable -Name 'PartnerCenterAppRefreshToken'

$AccessToken = New-PartnerAccessToken -ApplicationId '7adad76f-0bc2-45d5-8cee-d82429dbd377' -Scopes 'https://api.partnercenter.microsoft.com/user_impersonation' -RefreshToken 'OAQABAAAAAAAGV_bv21oQQ4ROqh0_1-tAFcqM-vRzkDcEKMIQu6hSp9LFatBVS9yCV0jBF4QUq3Dr9yxLoh0grQ3pkZLhe9Lym_yfUEnM3e7x0gCjMVjgENXKIhQ-p7xmQ8DIOQDCdgvEqyIEkvG6isc_yJu0F2OA2GPO6Gfk7OU3mVnBmNweG4vVEQk2dhbak7mGnLXrvZrKr6agZmNN3DiDh7tn1ylo2aY9FvYZgj95JUrGmJDS-FkX-Hc7q_C8vqNo_AX5OJiaV8hMG8F91EqjK78ba6cCTLkvteQHdKqe6Yi18kUFmZ1s7u5rxHJLb4WE7OCyb-4BFvqkObBUtheaLMV4nI8PfYoCvDbi6ptRnsaEUwH8KloQuQu3fEx0llpL_GDxGh6ak1vIjCFWERKHse3dkNA-WUf0uAyQ4n9c6D5wboqgiSKOc6xzSkFrLnCSAwt-bhM4Ov2guLBzUA3kHgkoFGgoiIlc11-gdA4cNbNuZjf7NB4Uz1PBRt06TDQZA2P9PpA3LX7J70Emr5RLRmv_G7k3hklQ325-XLFsoTXlq_716Nw_08L1SXpo7IAFhsIWeq5Qv5Mu0B3Jl63LSYcbhC-7mLZC95JMSRAI3mYmG_9gmQ_dFdKeqosn5WbuRm3EdHNgUVDYNnNb-vddRRyGuM-aZ9xb5klkCVZsKIBYaEXSzbOPxJTzkVefFO3ocFqu_-gNPYjVbLIo6YOzavSmBgXlsbhvOGNnBtaYlHAa5hxHT9R--qGn_AecvH2B-jJuUSeq5VDByYT-Az5BbQYmF25_z5qGt-hYqjVnpvO91g68HRLAAwZFOMkqMCQAawpleF08K1Cs9yM1a8g3vtabZ3z3I9Fzr7x4E_sWW-SILNoBh7FrqeRaKYBkYzFkCdsQlu0BdmFtkfmIgvP9QO7m7yHd0A4LSIF8ST6rEHMWDc9RefkBhAl12kQ-MZZc7VUc_07VNtp4TL84cpMsRkNDRd5UBalf5f_Bp1C9Vvwq6OQpFem0_z4qs1v2b0wiF16DTlGg_yPyCldu7wi2niccnERkKpG1uHihKlESplx_OVVNLJKVwqhMgDDxIFSHrSTqH__JuqBlijVUtM_Zdx_oQaijSiouiTOq26a8sYmhN2N_Yp4uKQeWwV3XwBsFtX9aNtggAA'
Set-AutomationVariable -Name 'PartnerCenterAppRefreshToken' -Value $AccessToken.RefreshToken

#$Credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $ClientAppId, $ClientSecret

$Connection = Connect-PartnerCenter -ApplicationId $ClientAppId -RefreshToken $RefreshToken

$UserLicenseTable = New-Object System.Data.DataTable
$UserLicenseTable.Columns.Add("User","System.String")|Out-Null
$UserLicenseTable.Columns.Add("Licenses","System.String")|Out-Null

$Users = Get-PartnerCustomerUser -CustomerId $TenantID

foreach ($User in $Users){
    $UserLicenses = Get-PartnerCustomerUserLicense -CustomerId $TenantID -UserId $User.UserId
    if($UserLicenses.Length -ge 1){
        $first = $True
        $LicenseList=""

        foreach($UserLicense in $UserLicenses){
           if($first -ne $True){
                $LicenseList += ", "
            }
            $first = $False
            $LicenseList+= $UserLicense.Name
        }
        $Row = $UserLicenseTable.NewRow()
        $Row.User = $User.DisplayName
        $Row.Licenses = $LicenseList
        $UserLicenseTable.Rows.Add($Row)
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
    $Row = $LicenseUsageTable.NewRow()
    $Row.License = $SKU.ProductName
    $Row.ConsumedCount = $SKU.ConsumedUnits
    $Row.ActiveCount = $SKU.ActiveUnits
    $Row.AvailableCount = $SKU.AvailableUnits
    $LicenseUsageTable.Rows.Add($Row)
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