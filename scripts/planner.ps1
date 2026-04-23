# Authenticate to Microsoft Graph
$body = @{
    client_id     = $env:CLIENT_ID
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $env:CLIENT_SECRET
    grant_type    = "client_credentials"
}

$token = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$env:TENANT_ID/oauth2/v2.0/token" -Body $body
$headers = @{ Authorization = "Bearer $($token.access_token)" }

# ---------------------------
# GET PLANNER TASKS
# ---------------------------

$planUrl = "https://graph.microsoft.com/v1.0/planner/plans/$($env:PLAN_ID)/tasks"
$tasks = Invoke-RestMethod -Headers $headers -Uri $planUrl -Method Get

# Build HTML summary
$html = "<h2>Planner Status Update</h2><ul>"
foreach ($task in $tasks.value) {
    $html += "<li><b>$($task.title)</b> — $($task.percentComplete)% complete</li>"
}
$html += "</ul>"

# ---------------------------
# SEND EMAIL FROM APP REGISTRATION
# ---------------------------

# Build sender address: {client_id}@{tenant}.onmicrosoft.com
$tenantDomain = (Invoke-RestMethod -Uri "https://login.microsoftonline.com/$env:TENANT_ID/v2.0/.well-known/openid-configuration").issuer.Split('/')[3]
$fromAddress = "$($env:CLIENT_ID)@$tenantDomain"

$mailBody = @{
    message = @{
        subject = "Planner Status Update"
        body = @{
            contentType = "HTML"
            content     = $html
        }
        from = @{
            emailAddress = @{
                address = $fromAddress
            }
        }
        toRecipients = @(
            @{ emailAddress = @{ address = $env:BOSS_EMAIL } }
        )
    }
    saveToSentItems = "false"
}

$sendMailUrl = "https://graph.microsoft.com/v1.0/users/$fromAddress/sendMail"
Invoke-RestMethod -Headers $headers -Uri $sendMailUrl -Method Post -Body ($mailBody | ConvertTo-Json -Depth 10)

