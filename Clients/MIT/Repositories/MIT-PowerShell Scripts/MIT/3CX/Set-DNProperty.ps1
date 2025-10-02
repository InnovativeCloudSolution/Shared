# Use $headers from Authentication.ps1

# Used to copy the value of a DNProperty from one dnNumber to another.

# You can use this to update the BLFS Template User if desired.

# Specify the dnNumber to use as a source, and the DNProperty to update.
$sourceDnNumber = "9041"
$dnPropertyName = "SHARED_BLFS"
# Specify the target dnNumber to update, 7002 is the BLFs template user.
$dnNumber = "7002"


# Get the source DNProperty values
$sourceResponse = Invoke-RestMethod "https://voip.manganoit.com.au:5001/xapi/v1/DNProperties/Pbx.GetDNPropertyByName(dnNumber='$sourceDnNumber',name='$dnPropertyName')" -Method 'GET' -Headers $headers

# We need to get the specific users Id for their DNProperty, every DNProperty has a unique ID. 
$targetResponse = Invoke-RestMethod "https://voip.manganoit.com.au:5001/xapi/v1/DNProperties/Pbx.GetDNPropertyByName(dnNumber='$dnNumber',name='$dnPropertyName')" -Method 'GET' -Headers $headers

# Creates a new body, using the Target Users DNPropertyID, and the value form a Source dnNumber. 
$body = @"
{
  `"dnNumber`": `"$dnNumber`",
  `"property`": {
    `"Description`": `"$($targetResponse.Description)`",
    `"Id`": `"$($targetResponse.Id)`",
    `"Name`": `"$dnPropertyName`",
    `"Value`": $($sourceResponse.Value | ConvertTo-Json)
  }
}
"@

$setResponse = Invoke-RestMethod 'https://voip.manganoit.com.au:5001/xapi/v1/DNProperties/Pbx.UpdateDNProperty' -Method 'POST' -Headers $headers -Body $body
$setResponse | ConvertTo-Json # In testing there was no response.