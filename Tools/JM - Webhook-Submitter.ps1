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
#CW-RPA Bot Testing
$webhookUrl = "https://au.webhook.myconnectwise.net/HbnQoSkfbblEL22NCL0OC5tioAKc210Qn8RDVW5754bdTNi8sO0t1c0h42g0oMybohfnBQ=="
$secretToken = "S-KG_X0Yb79EfGvbUr0OD85koAHN1wsQzcUSUzd_ttaBHt7tT5gWuRObOXwpd_QhXp7_4g=="

$WebhookData = Get-Content -Path "C:\Workspace\ics_workspace\ICS-Shared\Tools\JM - TestPayload.json" -Raw

$response = Invoke-RestMethod -Uri $webhookUrl -Method Post -Body $WebhookData -ContentType "application/json" -Headers @{ "x-cw-secret-token" = $secretToken }

if ($response) {
    Write-Output "Webhook triggered successfully. Response received:"
    $response
} else {
    Write-Output "No response received from the webhook"
}