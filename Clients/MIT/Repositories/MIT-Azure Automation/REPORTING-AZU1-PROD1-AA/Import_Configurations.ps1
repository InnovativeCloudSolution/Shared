param(
    [string]$ConfigFileData
)
$ConfigFileData | Out-File -FilePath "C:\temp\import.csv"
$Configs = Import-CSV "C:\temp\import.csv"

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Basic bWl0K2Q2anR1V01VOVkzQzFKdlk6TjExczRNZzkwUVZCUG1XYg==")
$headers.Add("clientId", "574a65e5-fe24-4984-bfac-cadf0590923b")
$headers.Add("content-type", "application/json")

$date=(Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

foreach ($Config in $Configs){
    $name = $Config.name
    $serialNumber = $Config.serialNumber
    $modelNumber = $Config.modelNumber
    $tagNumber = $Config.tagNumber
    $configurationIdentifier = $Config.configType
    $companyIdentifier = $Config.companyIdentifier

    $companyQuery = "https://aus.myconnectwise.net/v4_6_release/apis/3.0/company/companies/?conditions=identifier = '$companyIdentifier'"
    $companyDetails = Invoke-RestMethod $companyQuery  -Method 'GET' -Headers $headers -Body $body
    $company = $companyDetails.id

    $configurationQuery = "https://aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/types/?conditions=name = '$configurationIdentifier'"
    $configurationDetails = Invoke-RestMethod $configurationQuery  -Method 'GET' -Headers $headers -Body $body
    $configuration = $configurationDetails.id

    if($company -eq $null){
        Write-Error "Unable to find company with identifier: $companyIdentifier"
    }elseif($configuration -eq $null){
        Write-Error "Unable to find configuration with identifier: $configurationIdentifier"
    }else{

        $ConfigBody = "{
        `n  `"name`": `"$name`",
        `n  `"type`": {
        `n    `"id`": $configuration
        `n  },
        `n  `"status`": {
        `n    `"id`": 1
        `n  },
        `n  `"company`": {
        `n    `"id`": $company
        `n  },
        `n  `"deviceIdentifier`": `"`",
        `n  `"serialNumber`": `"$serialNumber`",
        `n  `"modelNumber`": `"$modelNumber `",
        `n  `"tagNumber`": `"$tagNumber`",
        `n  `"installationDate`": `"$date`",
        `n  `"billFlag`": false
        `n}"
        try{
            $response = Invoke-RestMethod 'https://aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations' -Method 'POST' -Headers $headers -Body $ConfigBody
            $responsename = $response.name
            Write-Output "Imported $responsename"
        }
        catch{
            Write-Output "Failed to import $ConfigBody"
        }
    }
}