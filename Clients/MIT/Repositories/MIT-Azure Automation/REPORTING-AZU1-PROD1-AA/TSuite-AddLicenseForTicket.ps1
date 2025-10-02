# Define parameters
param(
    [string]$Customer='',
    [string]$Product='',
    [string]$Edition='',
    [string]$Sku='',
    [string]$Add='',
    [string]$Number='',
    [int]$TicketId=0
)

#Get ITGlue apikey
$GlueApiKey = Get-AutomationVariable -Name 'ITGlue-ApiKey'

if ($null -ne $TicketId) {
	# Set CRM variables, connect to server
	$Server = Get-AutomationVariable -Name 'CWManageUrl'
	$Company = Get-AutomationVariable -Name 'CWManageCompanyId'
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
}

#Get current time
$CurrentTime = Get-Date -Format "HHmmss-ddMMyy"

#Construct Paths
$ScriptPath = "C:\Scripts\MIT-LicenseAutomation"
$LogName = 'logs\' + $Customer + ' - ' + $Sku + ' - ' + $CurrentTime + '.log'
$LogPath = Join-Path -Path $ScriptPath -ChildPath $LogName

#Change to script location
Set-Location $ScriptPath

# Initialise Logs variable and do loop counter
$Logs = ''
$LoopCounter = 1

# Runs the script until a successful result is received (can take multiple attempts due to oddities in Selenium WebDriver)
# or until it times out at 10 attempts
do {
	$Logs += "Attempt #$LoopCounter"
    $Logs += node app.js --c $Customer --p $Product --e $Edition --s $Sku --a $Add --n $Number --k $GlueApiKey | Out-String
	$Logs += "`n"
    $LoopCounter++
} until ($Logs -match 'Edition Update Complete' -or $LoopCounter -eq 11 -or $Logs -match 'is a child license, and needs to be purchased manually.')

$Logs | Out-File -FilePath $LogPath

if ($Logs -match 'Edition Update Complete') {
	if ($null -ne $TicketId) {
		# Grab and format order number and license count
		$LogsSplit = $Logs -split "Order Number is: "
		$OrderNumber = $LogsSplit[1].Substring(0, 8)
		$LicenseCountPart1 = $LogsSplit[0] -split "New count will be "
		$LicenseCountPart2 = $LicenseCountPart1[1] -split "`n"
		$LicenseCount = $LicenseCountPart2[0] -replace "`n", "."

		if ($Add -eq "True") {
			$TicketNotes = "$Product - $Edition (SKU: $Sku) license/s have been purchased successfully for $Customer." + "`n" + "`n" `
			+ "Order number: $OrderNumber" + "`n" + "Total licenses: $LicenseCount"
		}
		else {
			$TicketNotes = "$Product - $Edition (SKU: $Sku) license/s have been removed successfully for $Customer." + "`n" + "`n" `
			+ "Order number: $OrderNumber" + "`n" + "Total licenses: $LicenseCount"
		}

		# Add ticket note to ticket with details
		$NoteSuppressed = New-CWMTicketNote -ticketId $TicketId -text $TicketNotes -resolutionFlag $true -externalFlag $true

		# Add ticket note to ticket with log
		$LogsNoteSuppressed = New-CWMTicketNote -ticketId $TicketId -text $Logs -internalAnalysisFlag $true -internalFlag $true
	}
	
	Write-Output "Not Required"
}

else {
	if ($null -ne $TicketId) {
		$TicketNotes = "The automation ran into an issue while trying to purchase $Product - $Edition (SKU $Sku)" + "`n" + "`n" `
		+ "Logs:" + "`n" + $Logs

		# Add ticket note to ticket with details of failure
		$NoteSuppressed = New-CWMTicketNote -ticketId $TicketId -text $TicketNotes -resolutionFlag $true -externalFlag $true

		# Add ticket note to ticket with log
		$LogsNoteSuppressed = New-CWMTicketNote -ticketId $TicketId -text $Logs -internalAnalysisFlag $true -internalFlag $true
	}

	Write-Output "License Required"
}