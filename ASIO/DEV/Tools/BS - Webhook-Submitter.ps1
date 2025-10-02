$webhookUrl = "https://webhook-test.com/d73a09e47509d8b62cd8e5b360a6b02c"
#COSMO 
$webhookUrl = "https://1ce978e3-0165-4625-bc15-bdcdfaeb9c7e.webhook.ae.azure-automation.net/webhooks?token=Q88nSvS565POfFpyw7Wo6fciyjG%2beX7rJjNJ%2f%2bjC0RA%3d"
#BSWebhook
$webhookUrl = "https://au.webhook.myconnectwise.net/TuOAry4fabtELziMWb0OAZw3oFnN1AgQmcMVADMg49bUG9zu3ArqoiZeDJ1EaR0ooeg8xg=="
#MIT-Offboarding
$webhookUrl = "https://au.webhook.myconnectwise.net/HLnbqXwYbb9EeW_bXr0OCsxkoAHNgF8QlcNNAmYq5YHVGo3oS1kqoSFvEEdKYiAv0rMiyQ=="
# MIT-Testing
$webhookUrl = "https://au.webhook.myconnectwise.net/GuPT_3IeOLpEfznaWb0OW84xoFmd2g8QncFNBjMr49mGEYu7IyaYok3NfGiFHOizB2MiPg=="
# MIT-Onboarding
$webhookUrl = "https://au.webhook.myconnectwise.net/HbSHrnJJbexELm2IDb0OW85moFmd1VwQzpVBVzIqtNKFGdjrlzxOqNThFH-kbyccLctvZA=="
#MIT-AZU1-COSMO-WEBHOOK
$webhookUrl = "https://1ce978e3-0165-4625-bc15-bdcdfaeb9c7e.webhook.ae.azure-automation.net/webhooks?token=xUFmMKhtDzQVjGtnqDE6Xd8zMkBi2Zabt9BnzGkCFro%3d"
#MIT-AZU1-EH-WEBHOOK
$webhookUrl = "https://1ce978e3-0165-4625-bc15-bdcdfaeb9c7e.webhook.ae.azure-automation.net/webhooks?token=sQW%2fjdKaRYKge1K4LgYMGQnysAu9AHnwOozpjdJkYSY%3d"
#SNR-Offboarding-FinalTasks
$webhookUrl = "https://au.webhook.myconnectwise.net/HbTXqXpNbr1EeGuMDb0ODclmoAKR0lsQyMRHATcvsdaFEdm9W7y_nEm-t_XTNYn-XLQubw=="
#QRL-GuestUserOnboarding
$webhookUrl = "https://au.webhook.myconnectwise.net/HLmA-CwfbLlEKziLDr0OWphjoFic0gwQycQRU2F74NPXHo_vUORnmouLKE13o83cQPwFng=="

$payload = Get-Content -path "C:\Workspace\swanny-cloud-workspace\MIT-ASIO\DEV\Tools\BS - QRL - GuestUser - DD Payload.json"

$payload = @{
    "cwpsa_ticket" = "1094716"
    #"action" = "get"
    #"request_type" = "user_offboarding"
} | ConvertTo-Json

$response = Invoke-RestMethod -Uri $webhookUrl -Method Post -Body $Payload -ContentType "application/json"

if ($response) {
    Write-Host "Webhook submitted successfully."
} else {
    Write-Host "Failed to submit webhook."
}

#Test Commit
