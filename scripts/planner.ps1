# Authenticate to Microsoft Graph
$body = @{
    client_id     = $env:CLIENT_ID
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $env:CLIENT_SECRET
    grant_type    = "client_credentials"
}

$token = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$env:TENANT_ID/oauth2/v2.0/token" -Body $body
$headers = @{ Authorization = "Bearer $($token.access_token)" }

# Get Planner tasks
$tasks = Invoke-RestMethod -Headers $headers -Uri "https://graph.microsoft.com/v1.0/planner/plans/$env:PLAN_ID/tasks"

# Build HTML summary
$html = "<h2>Planner Status Update</h2><ul>"
foreach ($task in $tasks.value) {
    $html += "<li><b>$($task.title)</b> — $($task.percentComplete)% complete</li>"
}
$html += "</ul>"

# Send email FROM the app registration TO your boss
$fromAddress = "$($env:CLIENT_ID)@$((Invoke-RestMethod -Uri "https://login.microsoftonline.com/$env:TENANT_ID/v2.0/.well-known/openid-configuration").issuer.Split('/')[3])"

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

Invoke-RestMethod -Headers $headers -Uri "https://graph.microsoft.com/v1.0/sendMail" -Method Post -Body ($mailBody | ConvertTo-Json -Depth 10)



