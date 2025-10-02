
<#
    For: Mangano IT
    Date: 13.10.2022
    Engineer: Alex Williams
    Version: 0.1

.SYNOPSIS
    Gets details of Service Catalogue (type/subtype/items) for all boards in ConnectWise

.DESCRIPTION
    Connects to the ConnectWise API
    Gets all active boards
    Exports all Type/Subtype/Item associations to CSV

.INPUTS
    Bearer authorisation token

.OUTPUTS

    C:\temp\boards\BOARDNAME.csv
        
#>



$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("clientId", "1208536d-40b8-4fc0-8bf3-b4955dd9d3b7")
$headers.Add("Authorization", "Basic BEARER")

$Boards = Invoke-RestMethod 'https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/service/boards/?fields=id,name&pageSize=1000&conditions=inactiveFlag=false' -Method 'GET' -Headers $headers


foreach($Board in $Boards){
    $BoardId = $Board.id
    $BoardName = $Board.name
    $typeSubTypeItemAssociations = Invoke-RestMethod "https://api-aus.myconnectwise.net/v4_6_release/apis/3.0/service/boards/$BoardId/typeSubTypeItemAssociations/?pageSize=1000" -Method 'GET' -Headers $headers
    $associationResults = @()
    Foreach ($typeSubTypeItemAssociation in $typeSubTypeItemAssociations){


                    $associationResults += New-Object -TypeName PSObject -Property ([ordered]@{
                    typeName=$typeSubTypeItemAssociation.type.name;
                    subTypeName=$typeSubTypeItemAssociation.subType.name;
                    itemName=$typeSubTypeItemAssociation.item.name;
                    Remove=""
                    })

    }

    $associationResults | Export-csv -NoTypeInformation -Path "C:\Temp\Boards\$BoardName.csv"

}