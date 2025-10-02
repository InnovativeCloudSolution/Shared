## Generic parts
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("clientId", "1208536d-40b8-4fc0-8bf3-b4955dd9d3b7")
$headers.Add("Authorization", "Basic bWl0K2Q2anR1V01VOVkzQzFKdlk6TjExczRNZzkwUVZCUG1XYg==")
$headers.Add("Content-Type", "application/json")

## End generic parts

$CompanyName = "Kern Group"
$PCNBConfigName = "PC/Notebook"
$MWConfigName = "Managed Workstation"
$RedoChildren = $false

$nullParentConfigurationBody = "[
`n  {
`n    `"op`": `"replace`",
`n    `"path`": `"parentConfigurationId`",
`n    `"value`": null
`n  }
`n]"


#Set PCNB Parent to Null
if ($RedoChildren){
    Write-Host "Removing Parent Config from all PCNBs" 
    $PCNBConfigurations = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/?pageSize=1000&conditions=company/name='$CompanyName' %26%26 type/name='$PCNBConfigName'" -Method 'GET' -Headers $headers
    foreach ($Configuration in $PCNBConfigurations){
        $ConfigId = $Configuration.id
        Write-Host $ConfigId
        Write-Host $Configuration.company.name
        $response = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/$ConfigId" -Method 'PATCH' -Headers $headers -Body $nullParentConfigurationBody
    }
    Write-Host "Removed Parent Config from all PCNBs" 
}



#Set MW Parent to Null (Only do first run)
if ($RedoChildren){
    Write-Host "Removing Parent Config from all MWs" 
    $MWConfigurations = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/?pageSize=1000&conditions=company/name='$CompanyName' %26%26 type/name='$MWConfigName'" -Method 'GET' -Headers $headers
    foreach ($Configuration in $MWConfigurations){
        $ConfigId = $Configuration.id
        $response = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/$ConfigId" -Method 'PATCH' -Headers $headers -Body $nullParentConfigurationBody
    }
    Write-Host "Removed Parent Config from all MWs" 
}

#Bundle any unbundled managed workstations into PC/Notebooks
<#
$PCNBConfigurations = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/?pageSize=1000&conditions=company/name='$CompanyName' %26%26 type/name='$PCNBConfigName'" -Method 'GET' -Headers $headers
foreach ($Configuration in $PCNBConfigurations){
    $ChildConfigs = ""
    $SerialNumber = $Configuration.serialNumber
    $ConfigurationName = $Configuration.Name 
    $ChildConfigCount = $ChildConfigs.Count
    $SerialNumber = $SerialNumber.Substring(2,$SerialNumber.Length-2)

    #Assuming the serial isnt blank
    if($SerialNumber -ne ""){

        #Get all configs that are eligible children
        $ChildConfigs = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/?pageSize=1000&conditions=company/name='$CompanyName' %26%26 type/name='$MWConfigName' %26%26 serialNumber like '*$SerialNumber*' %26%26 parentConfigurationId==null" -Method 'GET' -Headers $headers
        #Write-Output $ConfigurationName $ChildConfigCount $SerialNumber
        $ConfigurationId = $Configuration.id

        $parentConfigurationBody = "[
        `n  {
        `n    `"op`": `"replace`",
        `n    `"path`": `"parentConfigurationId`",
        `n    `"value`": $ConfigurationId
        `n  }
        `n]"

        foreach ($ChildConfig in $ChildConfigs){
            $ChildConfigId = $ChildConfig.id
            Write-Host $Configuration.name
            Write-Host $ChildConfigId
            Write-Host $parentConfigurationBody
            $response = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/$ChildConfigId" -Method 'PATCH' -Headers $headers -Body $parentConfigurationBody

        }
    }
    
}
#>


#Find any PC/Notebooks without children, see if there are duplicate PC/Notebooks
#$PCNBConfigurations = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/?pageSize=1000&conditions=company/name='$CompanyName' %26%26 type/name='$PCNBConfigName'" -Method 'GET' -Headers $headers
#foreach($PCNB in $PCNBConfigurations){
#    $PCNBConfigId = $PCNB.id
#    $Children= Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/?pageSize=1000&conditions=parentConfigurationId=$PCNBConfigId" -Method 'GET' -Headers $headers
#    if($Children.Count -eq 0){
#        Write-Output $PCNB.name $($PCNB.serialNumber)
#    }
#}



$MultiplePCNBs = @()
#Make PCNBs for MWs Without a PCNB and bundles ones where there is one
$MWConfigurations = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/?pageSize=1000&conditions=company/name='$CompanyName' %26%26 type/name='$MWConfigName' %26%26 parentConfigurationId==null" -Method 'GET' -Headers $headers
foreach ($MW in $MWConfigurations){
    $MWID = $MW.id

    $ConfigName=""
    $SerialNumber=""
    $modelNumber=""
    $tagnumber=""
    $contactId=""
    $companyId=""
    $statusId=""
    $ApplicablePCNBs=$null

    $ConfigName = $MW.Name
    $serialNumber = $MW.serialNumber
    $modelNumber = $MW.modelNumber
    $tagNumber = $MW.tagNumber
    $contactId = $MW.contact.id
    $companyId = $MW.company.id
    $statusId = $MW.status.id

        if($serialNumber -ne ""){

        $ApplicablePCNBs = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/?pageSize=1000&conditions=company/name='$CompanyName' %26%26 type/name='$PCNBConfigName' %26%26 serialNumber like '*$serialNumber*'" -Method 'GET' -Headers $headers
        if($ApplicablePCNBs.Count -ge 1){
                if($ApplicablePCNBs.Count -eq 1){
                    $PCNBConfigId = $ApplicablePCNBs.id
                }else{
                    $PCNBConfigId = $ApplicablePCNBs[0].id
                    $MultiplePCNBs += @{
                        ConfigName=$ConfigName
                        SerialNumber=$serialNumber
                    }
                    Write-Host "Multiple PCNBs Found for $ConfigName $serialNumber. Using $PCNBConfigId" -ForegroundColor Red
                }

                $parentConfigurationBody = "[
                `n  {
                `n    `"op`": `"replace`",
                `n    `"path`": `"parentConfigurationId`",
                `n    `"value`": $PCNBConfigId
                `n  }
                `n]"
                Write-Host "Found applicable parent PCNB for $ConfigName $serialNumber, ParentId = $PCNBConfigId, ChildId = $MWID" -ForegroundColor DarkYellow
                #pause
                $response = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/$MWID" -Method 'PATCH' -Headers $headers -Body $parentConfigurationBody



        }else{

            $ConfigBody = "{
            `n    `"name`": `"$ConfigName`",
            `n    `"status`": {
            `n        `"id`": $statusId
            `n    },
            `n    `"company`": {
            `n        `"id`": $companyId
            `n    },
            `n    `"type`": {
            `n        `"id`": 25
            `n    },
            `n    `"serialNumber`": `"$serialNumber`",
            `n    `"modelNumber`": `"$modelNumber`",
            `n    `"tagNumber`": `"$tagNumber`",
            `n    `"name`": `"$ConfigName`",
            `n}"

            Write-Host "Could not find applicable parent PCNB for $ConfigName $serialNumber, making..." -ForegroundColor Yellow
            #pause

            $response = Invoke-RestMethod 'https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations' -Method 'POST' -Headers $headers -Body $ConfigBody
            $ResponseConfigId = $response.id

            $parentConfigurationBody = "[
            `n  {
            `n    `"op`": `"replace`",
            `n    `"path`": `"parentConfigurationId`",
            `n    `"value`": $ResponseConfigId
            `n  }
            `n]"
        
            Write-Host "Made PCNB for $ConfigName, ID $ResponseConfigId, bundling in $MWID" -ForegroundColor Green
            #pause

            $response = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/company/configurations/$MWID" -Method 'PATCH' -Headers $headers -Body $parentConfigurationBody
        }
        Write-Host "Done, next!" -ForegroundColor Gray
        #pause
    }
}

