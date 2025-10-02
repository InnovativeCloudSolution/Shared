# Use $headers from Authentication.ps1

# Used to copy a DNProperty value, such as BLFS from one user to many. 

# Specify the dnNumber to use as a source, and the DNProperty to update. 7002 is the BLFs template user, or use your own number after setting them correctly. 
$sourceDnNumber = "7002"
$dnPropertyName = "SHARED_BLFS"

# Specify the dnNumbers to update by removing numbers from the 100 block.
# $numbersToRemove = 9000..9019 + 9028 + 9043 + 9050 + 9060 + 9064 + 9066 + 9068..9097 + 9099
# NumbersToRemove last updated on 2024/11/20
$numbersToRemove = @(
    9000
    9028
    9050
    9064
    9066
    9077
    9088
    9099
)
# Get current 9xxx numbers
$numbers = Invoke-RestMethod 'https://voip.manganoit.com.au:5001/xapi/v1/Users?%24top=100&%24skip=0&%24orderby=Number&%24select=Number&%24filter=startswith(Number, ''9'')' -Method 'GET' -Headers $headers
# Create array without un-used numbers.
$dnNumbersArray = $numbers.value.Number | Where-Object { $_ -notin $numbersToRemove }

# Get the source DNProperty value
$sourceResponse = Invoke-RestMethod "https://voip.manganoit.com.au:5001/xapi/v1/DNProperties/Pbx.GetDNPropertyByName(dnNumber='$sourceDnNumber',name='$dnPropertyName')" -Method 'GET' -Headers $headers

# Use foreach to update each dnNumber in the array. 
foreach ($dnNumber in $dnNumbersArray) {
    # Get the target dnProperty Id and current value, every DNProperty has a unique ID - so this is required to update it.
    $targetResponse = Invoke-RestMethod "https://voip.manganoit.com.au:5001/xapi/v1/DNProperties/Pbx.GetDNPropertyByName(dnNumber='$dnNumber',name='$dnPropertyName')" -Method 'GET' -Headers $headers
    
    # Compare the target vs source values, update the target value if they don't match.
    if ($targetResponse.Value -ne $sourceResponse.Value) {
        # Create a new body, using the Target DNProperty ID, and the value from a Source dnNumber.
        $body = @{
            dnNumber = "$dnNumber"
            property = @{
                Description = $targetResponse.Description
                Id = $targetResponse.Id
                Name = $dnPropertyName
                Value = $sourceResponse.Value
            }
          } | ConvertTo-Json
        
        # Update the Target dnNumber so it matches the Source Response.
        Invoke-RestMethod 'https://voip.manganoit.com.au:5001/xapi/v1/DNProperties/Pbx.UpdateDNProperty' -Method 'POST' -Headers $headers -Body $body
    }
}