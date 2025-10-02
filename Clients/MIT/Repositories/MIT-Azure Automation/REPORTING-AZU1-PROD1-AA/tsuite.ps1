# Define parameters
param(
    [string]$customer='',
    [string]$product='',
    [string]$edition='',
    [string]$sku='',
    [string]$add='',
    [string]$number=''
)

#Get ITGlue apikey
$key = Get-AutomationVariable -Name 'ITGlue-ApiKey'

#GetTime
$time = Get-Date -Format "HHmmss-ddMMyy"

#Construct Paths
$path = "C:\Scripts\MIT-LicenseAutomation"
$logname = 'logs\' + $customer + ' - ' + $license + ' - ' + $time + '.log'
$logpath = Join-Path -Path $path -ChildPath $logname

#Change to script location
cd $path
#Run script and save to log
node app.js --c $customer --p $product --e $edition --s $sku --a $add --n $number --k $key | Out-File -FilePath $logpath
#Read log for output
Get-Content -Path $logpath