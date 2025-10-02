<#

Mangano IT - Format Valid Phone Number
Created by: Gabriel Nugent
Version: 2.2.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory)][string]$PhoneNumber,
    [string]$AreaCode,
    [bool]$IncludePlus = $true,
    [bool]$IsMobileNumber = $true,
    [bool]$KeepSpaces = $false,
    [string]$CustomFormat,
    [int]$CustomLength
)

## SET UP NUMBER DETAILS ##

# Check if an area code has been provided in the number, and create one if it exists
if (($PhoneNumber.StartsWith("+")) -and ($AreaCode -eq '') -and ($IncludePlus)) {
    Write-Warning "No area code specifically provided, but the phone number includes a plus."
    $AreaCode = $PhoneNumber.Substring(0, 3) -replace '\W', ''
    $PhoneNumber = $PhoneNumber.Substring(3, ($PhoneNumber.Length - 3))
    Write-Warning "Area code created: $($AreaCode) | Phone number: $($PhoneNumber)"
}

# Remove illegal characters
$PhoneNumber = $PhoneNumber -replace '[^0-9]', ''
$AreaCode = $AreaCode -replace '[^0-9]', ''

# Remove leading zero if it exists
if ($PhoneNumber.StartsWith('0')) {
    $PhoneNumber = $PhoneNumber.Substring(1, ($PhoneNumber.Length - 1))
    Write-Warning "Number started with a 0, and has now been trimmed. | Phone number: $($PhoneNumber)"
}

# Define phone number lengths
if ($AreaCode -ne '') {
    if ($IncludePlus) {
        $LeadingNumbers = '+' + $AreaCode
        Write-Warning "Area code provided and plus requested. | Leading digits: $($LeadingNumbers)"
        $Length = 12
    } else {
        $LeadingNumbers = $AreaCode
        Write-Warning "Area code provided and plus to be excluded. | Leading digits: $($LeadingNumbers)"
        $Length = 11
    }
} else {
    # Set default leading number
    $LeadingNumbers = '0'
    Write-Warning "No area code provided. | Leading digits: $($LeadingNumbers)"
    $Length = 10 
}

# Establish number format
if ($CustomFormat) {
    $Format = "{0:$CustomFormat}"
    $Length = $CustomLength
    if ($IncludePlus) {
        $LeadingNumbers = '+'
    } else {
        $LeadingNumbers = ''
    }
    Write-Warning "Custom format requested. | Leading digits: $($LeadingNumbers) | Format: $($Format)"
} else {
    if ($KeepSpaces) {
        $Length += 2
        if ($IsMobileNumber) {
            if ($AreaCode -ne '') {
                $Format = '{0: ### ### ###}'
                $Length += 1
            } else {
                $Format = '{0:### ### ###}'
            }
        } else {
            $Format = '{0:# #### ####}'
        }
    } else {
        $Format = '{0:#########}'
    }
    Write-Warning "Format chosen. | Format: $($Format)"
}

## FORMAT NUMBER ##

# Check if the number is longer than the # count in the format string
$FormatSansFormat = $Format -replace "#", ''
$HashCount = $Format.Length - $FormatSansFormat.Length
if ($PhoneNumber.Length -gt $HashCount) {
    $EndDigits = $PhoneNumber.Substring($HashCount, ($PhoneNumber.Length - $HashCount))
    $PhoneNumber = $PhoneNumber.Substring(0, $HashCount)
    Write-Warning "There are more digits in the phone number than hashes in the format. Overlap will be tacked on at the end. | Phone number: $($PhoneNumber) | End digits: $($EndDigits)"
}

# Format number
$Output = $LeadingNumbers + $Format -f [double]$PhoneNumber

# Add end digits if they exist
if ($null -ne $EndDigits) {
    $Output += $EndDigits
    Write-Warning "Added end digits onto output. | End digits: $($EndDigits) | Output: $($Output)"
}

# Send back to flow if number contains digits
if ($Output -match '\d') {
    Write-Warning "Valid number created. | Output: $($Output)"
    Write-Output $Output
} else {
    Write-Warning "'$($Output)' is not a valid number, and will not be written back."
    Write-Output ''
}