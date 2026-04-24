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

if (-not $env:PLAN_ID) { throw "PLAN_ID environment variable is missing!" }
$planUrl = "https://graph.microsoft.com/v1.0/planner/plans/$($env:PLAN_ID.Trim())/tasks"

$tasks = Invoke-RestMethod -Headers $headers -Uri $planUrl -Method Get

# Build HTML summary
if ($tasks.value.Count -eq 0) {
    $html = "<h2>Planner Status Update</h2><p>No active tasks found in the plan at this time.</p>"
} else {
    $html = "<h2>Planner Status Update</h2><ul>"
    foreach ($task in $tasks.value) {
        $html += "<li><b>$($task.title)</b> — $($task.percentComplete)% complete</li>"
    }
    $html += "</ul>"
}

# ---------------------------
# SEND EMAIL FROM APP REGISTRATION
# ---------------------------

# Build sender address: {client_id}@{tenant}.onmicrosoft.com
$senderEmail = $env:SENDER_EMAIL 

$mailBody = @{
    message = @{
        subject = "Planner Status Update"
        body = @{
            contentType = "HTML"
            content     = $html
        }
        # Note: 'from' is usually redundant if it matches the URL path
        toRecipients = @(
            @{ emailAddress = @{ address = $env:BOSS_EMAIL } }
        )
    }
    saveToSentItems = "false"
}

# The URL must point to a real user/mailbox
$sendMailUrl = "https://graph.microsoft.com/v1.0/users/$senderEmail/sendMail"

Invoke-RestMethod -Headers $headers -Uri $sendMailUrl -Method Post -Body ($mailBody | ConvertTo-Json -Depth 10) -ContentType "application/json"

