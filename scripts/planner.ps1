# --- Authentication (Existing Logic) ---
$body = @{
    client_id     = $env:CLIENT_ID
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $env:CLIENT_SECRET
    grant_type    = "client_credentials"
}
$token = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$env:TENANT_ID/oauth2/v2.0/token" -Body $body
$headers = @{ Authorization = "Bearer $($token.access_token)" }

# --- Get Planner Tasks ---
$planId = $env:PLAN_ID.Trim()
$planUrl = "https://graph.microsoft.com/v1.0/planner/plans/$planId/tasks"
$tasks = Invoke-RestMethod -Headers $headers -Uri $planUrl -Method Get

# --- Process Tasks ---
$reportItems = @()

foreach ($task in $tasks.value) {
    $pId = $task.id
    $title = $task.title
    
    # Check for existing GitHub Issue using the Planner ID as a label or search term
    $issue = gh issue list --search "$pId" --json number,title | ConvertFrom-Json | Select-Object -First 1
    
    if (-not $issue) {
        # STEP 1: SYNC (Create issue if missing)
        $issueNumber = gh issue create --title "$title" --body "PlannerID: $pId `n---`nUpdates:"
        Write-Host "Created new issue for task: $title"
    } else {
        $issueNumber = $issue.number
    }

    # STEP 2: GET UPDATES (Grab latest comment)
    $issueData = gh issue view $issueNumber --json comments | ConvertFrom-Json
    $latestComment = $issueData.comments | Select-Object -Last 1
    $note = if ($latestComment) { $latestComment.body } else { "No updates recorded in meeting." }

    # Store for the email
    $reportItems += [PSCustomObject]@{
        Title    = $title
        Percent  = $task.percentComplete
        Notes    = $note
    }
}

# --- Post-Meeting Report (Email Logic) ---
if ($env:RUN_MODE -eq "report") {
    $html = "<h2>Monday Standup Report: $(Get-Date -Format 'MM/dd/yyyy')</h2>"
    foreach ($item in $reportItems) {
        $html += "<p><b>$($item.Title)</b> ($($item.Percent)% complete)<br/>"
        $html += "<i>Update:</i> $($item.Notes)</p><hr/>"
    }

    $mailBody = @{
        message = @{
            subject = "Project Update: $(Get-Date -Format 'D')"
            body = @{ contentType = "HTML"; content = $html }
            toRecipients = @( @{ emailAddress = @{ address = $env:BOSS_EMAIL } } )
        }
    }
    
    $sendUrl = "https://graph.microsoft.com/v1.0/users/$env:SENDER_EMAIL/sendMail"
    Invoke-RestMethod -Headers $headers -Uri $sendUrl -Method Post -Body ($mailBody | ConvertTo-Json -Depth 10) -ContentType "application/json"
    Write-Host "Report emailed to boss."
}
