param(
	[Parameter(Mandatory=$true)][string]$Username,
	[Parameter(Mandatory=$true)][string]$DisplayName,
	[Parameter(Mandatory=$true)][string]$CurrentDomain
)

# Get certificate thumbprint and App ID
$CertificateThumbprint = Get-AutomationVariable -Name 'EXO-CertificateThumbprint'
$AppId = Get-AutomationVariable -Name 'OPC-AppId'

# Construct alias
$Alias = $Username.Substring(0,$Username.Length-1) + ".epoc"

# Construct domain
$Domain = $CurrentDomain.Substring(1, $CurrentDomain.Length)

# Construct email addresses and new display name
$EpocEmail = $Username+"epocenviro.com"
$MainEmail = $Username+$Domain
$NewDisplayName = "$DisplayName - EPOC Enviro"

$Log=""
$Log += "Connecting to Exchange Online"
Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbprint -AppId $AppId -Organization opecsystems.onmicrosoft.com

# Create the shared mailbox, and assign required permissions
$Log += "Attempting to create shared mailbox with SMTP address $EpocEmail..."
try {
	New-Mailbox -Shared -Name $NewDisplayName -DisplayName $NewDisplayName -Alias $Alias -PrimarySmtpAddress $EpocEmail | `
	Set-Mailbox -GrantSendOnBehalfTo $MainEmail | Add-MailboxPermission -User $MainEmail -AccessRights FullAccess -InheritanceType All
	$Log += "New mailbox created successfully, access rights granted to $MainEmail"
}
catch {
	$Log += "New mailbox not created"
}

Write-Output $Log