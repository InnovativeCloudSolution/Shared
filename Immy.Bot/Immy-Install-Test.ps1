$Uri = "https://au.webhook.myconnectwise.net/TubUqHIYP7BEfTjZX70OAJhroAGZhVwQysRGUzN_44XRGNzhvOy7Xzp5NtDWJ3VXv09P7Q=="

$Body = @{
    intent_name = "GLOBAL - Software - Installation Request"
    intent_fields = @{
        request_for    = "Myself"
        full_name      = "Juan Moredo"
        email_address  = "Juan.Moredo@manganoit.com.au"
        request_name   = "Google Chrome"
        request_type   = "Add"
        reason         = "Required for my role"
    }
    meta_data = @{
        ticket_id         = 1036925
        ticket_board_name = "Internal"
        ticket_board_id   = 80
        contact_id        = 18090
        contact_name      = "Juan Moredo"
        contact_email     = "juan.moredo@manganoit.com.au"
        company_id        = "2"
        company_name      = "Mangano IT"
        company_types     = @("Owner", "CloudOnly")
    }
} | ConvertTo-Json -Depth 5

$Headers = @{
    "Content-Type" = "application/json"
}

Invoke-RestMethod -Uri $Uri -Method Post -Body $Body -Headers $Headers