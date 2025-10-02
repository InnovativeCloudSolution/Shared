# Use $headers from Authentication.ps1

# Used to get the value of a DN Property by it's name.
# From DN Properties tab of https://voip.manganoit.com.au:5001/#/office/advanced/parameters

#Specify the desired dnNumber and DNProperty to fetch. 
$dnNumber = "9041"
$dnPropertyName = "SHARED_BLFS"

$sourceResponse = Invoke-RestMethod "https://voip.manganoit.com.au:5001/xapi/v1/DNProperties/Pbx.GetDNPropertyByName(dnNumber='$dnNumber',name='$dnPropertyName')" -Method 'GET' -Headers $headers

$sourceValue = $sourceResponse.Value
$sourceValue | ConvertTo-Json