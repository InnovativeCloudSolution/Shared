param(
	[Parameter(Mandatory=$true)][string]$Username,
	[Parameter(Mandatory=$true)][string]$DisplayName,
	[Parameter(Mandatory=$true)][string]$CurrentDomain,
	[Parameter(Mandatory=$true)][string]$NewDomain
)

# Get certificate thumbprint and App ID
$CertificateThumbprint = Get-AutomationVariable -Name 'EXO-CertificateThumbprint'
$AppId = Get-AutomationVariable -Name 'OPC-AppId'

# Construct alias
$Alias = $Username.Substring(0,$Username.Length-1) + ".epoc"

# Construct domains
$CurrentDomainFixed = $CurrentDomain.Substring(1, $CurrentDomain.Length - 1)
$NewDomainFixed = $NewDomain.Substring(1, $NewDomain.Length - 1)

# Create title for shared mailbox display name
$NewDomainTitle = ''

switch ($CurrentDomainFixed) {
	'opecsystems.com' { $NewDomainTitle = 'OPEC Systems' }
	'opeccbrne.com' { $NewDomainTitle = 'OPEC CBRNE' }
	'opeccollege.edu.au' { $NewDomainTitle = 'OPEC College' }
	'epocenviro.com' { $NewDomainTitle = 'EPOC Enviro' }
}

# Construct email addresses and new display name
$NewEmail = $Username+$NewDomainFixed
$MainEmail = $Username+$CurrentDomainFixed
$NewDisplayName = "$DisplayName - $NewDomainTitle"

$Log=""
$Log += "Connecting to Exchange Online"
Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbprint -AppId $AppId -Organization opecsystems.onmicrosoft.com

# Create the shared mailbox, and assign required permissions
$Log += "Attempting to create shared mailbox with SMTP address $NewEmail..."
try {
	New-Mailbox -Shared -Name $NewDisplayName -DisplayName $NewDisplayName -Alias $Alias -PrimarySmtpAddress $NewEmail | `
	Set-Mailbox -GrantSendOnBehalfTo $MainEmail | Add-MailboxPermission -User $MainEmail -AccessRights FullAccess -InheritanceType All
	$Log += "New mailbox created successfully, access rights granted to $MainEmail"
}
catch {
	$Log += "New mailbox not created"
}

Write-Output $Log