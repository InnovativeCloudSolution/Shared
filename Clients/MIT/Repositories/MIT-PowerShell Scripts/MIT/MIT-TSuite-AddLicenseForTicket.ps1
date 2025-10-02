# Define parameters
param(
    [string]$Customer='',
    [string]$Product='',
    [string]$Edition='',
    [string]$Sku='',
    [string]$Add='',
    [string]$Number='',
    [int]$TicketId
)

#Get ITGlue apikey
$GlueApiKey = Get-AutomationVariable -Name 'ITGlue-ApiKey'

# Set CRM variables, connect to server
$Server = 'https://crm.manganoit.com.au'
$Company = 'mits'
$PublicKey = Get-AutomationVariable -Name 'PublicKey'
$PrivateKey = Get-AutomationVariable -Name 'PrivateKey'
$ClientId = Get-AutomationVariable -Name 'clientId'

# Create a credential object
$Connection = @{
    Server = $Server
    Company = $Company
    pubkey = $PublicKey
    privatekey = $PrivateKey
    clientId = $ClientId
}

# Connect to Manage server
Connect-CWM @Connection

#Get current time
$CurrentTime = Get-Date -Format "HHmmss-ddMMyy"

#Construct Paths
$ScriptPath = "C:\scripts\tsuite"
$LogName = 'logs\' + $Customer + ' - ' + $Sku + ' - ' + $CurrentTime + '.log'
$LogPath = Join-Path -Path $ScriptPath -ChildPath $LogName

#Change to script location
Set-Location $ScriptPath

# Initialise Logs variable and do loop counter
$Logs = ''
$LoopCounter = 0

# Runs the script until a successful result is received (can take multiple attempts due to oddities in Selenium WebDriver)
# or until it times out at 10 attempts
do {
    node app.js --c $Customer --p $Product --e $Edition --s $Sku --a $Add --n $Number --k $GlueApiKey | Out-File -FilePath $LogPath
    $Logs = Get-Content -Path $LogPath
    $LoopCounter++
} until ($Logs -contains 'Edition Update Complete' -or $LoopCounter -eq 10)

if ($Logs -contains 'Edition Update Complete') {
    # Grab and format order number
    $LogsSplit = $Logs -split "Order Number is: "

    # Add ticket note to ticket with details
    New-CWMTicketNote -ticketId $TicketId -text "A $Product - $Edition (SKU $Sku) license has been purchased successfully.\n\n`
    The order number is: " + $LogsSplit[1].Substring(0, 8) -customerUpdatedFlag $false -internalAnalysisFlag $true -internalFlag $true `
    -Confirm $false
}

else {
    # Add ticket note to ticket with details of failure
    New-CWMTicketNote -ticketId $TicketId -text "The automation ran into an issue while trying to purchase $Product - $Edition (SKU $Sku)\n\n`
    Logs:\n" + $Logs -customerUpdatedFlag $true -internalAnalysisFlag $true -internalFlag $true -Confirm $false
}

Write-Output $Logs