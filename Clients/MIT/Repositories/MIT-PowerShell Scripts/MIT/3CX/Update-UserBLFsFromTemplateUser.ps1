# 3CX Script to update all 9xxx numbers from a source, template user 7002.
# Created by Joshua Ceccato and Juan Moredo. 
# 2024/11/20

# Specify the dnNumber to use as a source, and the DNProperty to update. 7002 is the BLFs template user.
$sourceDnNumber = "7002"
$dnPropertyName = "SHARED_BLFS"
$secret3cx = "3CXUser7000" # AKV Secret name for the 3CX user #7000's password.

## Authentication 
function Get-MSGraph-Secrets {
    param (
        [string]$AzKeyVaultName,
        [string[]]$SecretNames
    )
    try {
        Connect-AzAccount -Identity | Out-Null
        $secrets = @{}
        foreach ($secretName in $SecretNames) {
            $secretValue = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name $secretName -AsPlainText
            $secrets[$secretName] = $secretValue
        }
        return $secrets
    }
    catch {
        Write-Error -Message $_.Exception.Message
        throw $_.Exception
    }
}

# Get the PW from Azure KeyVault.
$secrets = Get-MSGraph-Secrets "MIT-AZU1-PROD1-AKV1" $secret3cx
$password = $secrets[$secret3cx]
# $password ="put me in keyvault?" # https://mits.au.itglue.com/797747/passwords/3817176368334888

$headers = @{
    "Content-Type" = "application/json"
}

$body = @"
{`"SecurityCode`":`"`",`"Password`":`"$password`",`"Username`":`"7000`"}
"@

# Get the auth token. 
$response = Invoke-RestMethod 'https://voip.manganoit.com.au:5001/webclient/api/Login/GetAccessToken' -Method 'POST' -Headers $headers -Body $body
$ACCESS_TOKEN = $response.token.access_token
$headers["Authorization"] = "Bearer $ACCESS_TOKEN"

# -----
## Update BLFs code.

# Specify the 9xxx dnNumbers which should not be updated
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
# Create array without $numbersToRemove.
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