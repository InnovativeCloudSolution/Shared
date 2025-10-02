### THIS IS A DRAFT - IT DOES NOT FULLY FUNCTIOn
### Awaiting MUB-8916 to be completed by ConnectWise (#16572788)

<#
    For: Mangano IT
    Date: 13.10.2022
    Engineer: Alex Williams
    Version: 0.1

.SYNOPSIS
    Sets details of Service Catalogue (type/subtype/items) for selected board in ConnectWise

.DESCRIPTION
    Connects to the ConnectWise API
    Finds CSV in C:\Temp\Boards\BoardName.csv
    For each item, ensures that the relevant Type/Subtype/Item association exists, if not, makse

.INPUTS
    $LookForName = Board Name

.OUTPUTS

    N/A
        
#>

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("clientId", "1208536d-40b8-4fc0-8bf3-b4955dd9d3b7")
$headers.Add("Authorization", "Basic BEARER")
$headers.Add("Content-Type", "application/json")

$LookForName = "Automation (MS)"

$associationResults = Import-csv -Path "C:\Temp\Boards\$LookForName.csv"
$Board = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/service/boards/?fields=id,name&pageSize=1000&conditions=name='$LookForName'" -Method 'GET' -Headers $headers
$BoardId = $Board.id


$count=0;
Foreach ($associationResult in $associationResults){
    $type = $associationResult.typeName
    $subType = $associationResult.subTypeName
    $item = $associationResult.itemName
    $remove = $associationResult.Remove

    $query = "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/service/boards/$BoardId/typeSubTypeItemAssociations?conditions=type/name='$type' and subType/name='$subType' and item/name='$item'"
    if ($item -eq ''){
        $query = "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/service/boards/$BoardId/typeSubTypeItemAssociations?conditions=type/name='$type' and subType/name='$subType' and item/name=null"
    }

    $GoodStatus = $true
    $assocs = Invoke-RestMethod $query -Method 'GET' -Headers $headers

    if($assocs.Length -eq 0){
        ##We need to add it and associate it

        ## Get Type id (or make it)
        $query = "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/service/boards/$BoardId/types?conditions=name='$type'"
        $typeresult = Invoke-RestMethod $query -Method 'GET' -Headers $headers
        if($typeresult.Count -eq 0){
            ## We need to add it
            Write-Host "Adding Type: $type"
            $body = "{
            `n    `"name`": `"$type`"
            `n}"

            $typeresult = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/service/boards/$BoardId/types" -Method 'POST' -Headers $headers -Body $body
            Start-Sleep -Milliseconds 500
        }else{
            ## It exists already:
            Write-Host "Found $type -"$typeresult.id
        }


        ## Get subType id (or make it)
        $query = "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/service/boards/$BoardId/subTypes?conditions=name='$subType'"
        $subtyperesult = Invoke-RestMethod $query -Method 'GET' -Headers $headers
        if($subtyperesult.Count -eq 0){
            ## We need to add it
            Write-Host "Adding Subtype $subtype"
            $body = "{
            `n    `"name`": `"$subtype`"
            `n}"

            $subtyperesult = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/service/boards/$BoardId/subtypes" -Method 'POST' -Headers $headers -Body $body
            Start-Sleep -Milliseconds 500
        }else{
            ## It exists already:
            Write-Host "Found $subtype -"$subtyperesult.id
        }


        ## Get Item id (or make it)
        $query = "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/service/boards/$BoardId/items?conditions=name='$item'"
        $itemresult = Invoke-RestMethod $query -Method 'GET' -Headers $headers

        if($itemresult.Count -eq 0){
            ## We need to add it
            Write-Host "Adding Item $item"
            $body = "{
            `n    `"name`": `"$item`"
            `n}"

            $itemresult = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/service/boards/$BoardId/items" -Method 'POST' -Headers $headers -Body $body
            Start-Sleep -Milliseconds 500
        }else{
            ## It exists already:
            Write-Host "Found $item -"$itemresult.id
        }
        ## Add new mapping
        ## Ensure Type-Subtype Mapping
        $TypeSubtypes = $subtyperesult.typeAssociationIds
        $TypeSubtypes += $typeresult.id
        $TypeSubtypes = $TypeSubtypes | Select -Unique

        $TypeSubtypesString = $TypeSubtypes -join ","
        $body = "[{
        `n    `"op`": `"replace`",
        `n    `"path`":`"typeAssociationIds`",
        `n    `"value`": [$TypeSubtypesString]
        `n}]"
        Write-Host $body

        $response = Invoke-RestMethod 'https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/service/boards/67/subtypes/2981' -Method 'PATCH' -Headers $headers -Body $body
        Start-Sleep -Milliseconds 500


    }elseif($remove -eq "TRUE"){
        Write-Host "Would remove $type $subType $item $remove" -ForegroundColor Red
    }
    else{

        Write-Host "Do nothing $type $subType $item $remove" -ForegroundColor Green

    }


    $count++

}
Write-Host "Count "$count