#Webhook Tester
$webhookUrl = "https://webhook-test.com/8058d285092b6d30aae44932204df24f"
#JM Webhook
$webhookUrl = "https://au.webhook.myconnectwise.net/TeLa-y9IPb5Eem_ZWb0OWs5loFjLhw8QnpFABWYh59DVEd66RTLqDAixew4V7y_mMIuXtg=="
#MIT-UserOffboarding
$webhookUrl = "https://au.webhook.myconnectwise.net/HLnbqXwYbb9EeW_bXr0OCsxkoAHNgF8QlcNNAmYq5YHVGo3oS1kqoSFvEEdKYiAv0rMiyQ=="
#MIT-UserOnboarding
$WebhookUrl = "https://au.webhook.myconnectwise.net/HbSHrnJJbexELm2IDb0OW85moFmd1VwQzpVBVzIqtNKFGdjrlzxOqNThFH-kbyccLctvZA=="
#MIT-AZU1-DD-WEBHOOK
$webhookUrl = "https://1ce978e3-0165-4625-bc15-bdcdfaeb9c7e.webhook.ae.azure-automation.net/webhooks?token=Q88nSvS565POfFpyw7Wo6fciyjG%2beX7rJjNJ%2f%2bjC0RA%3d"
#MIT-AZU1-COSMO-WEBHOOK
$webhookUrl = "https://1ce978e3-0165-4625-bc15-bdcdfaeb9c7e.webhook.ae.azure-automation.net/webhooks?token=xUFmMKhtDzQVjGtnqDE6Xd8zMkBi2Zabt9BnzGkCFro%3d"
#MIT-AZU1-EH-WEBHOOK
$webhookUrl = "https://1ce978e3-0165-4625-bc15-bdcdfaeb9c7e.webhook.ae.azure-automation.net/webhooks?token=sQW%2fjdKaRYKge1K4LgYMGQnysAu9AHnwOozpjdJkYSY%3d"


$WebhookData = Get-Content -Path "C:\Workspace\ics_workspace\MIT-ASIO\DEV\Tools\JM - TestPayload.json" -Raw

$response = Invoke-RestMethod -Uri $webhookUrl -Method Post -Body $WebhookData -ContentType "application/json"

if ($response) {
    Write-Output "Webhook triggered successfully. Response received:"
    $response
} else {
    Write-Output "No response received from the webhook"
}