# --- Authentication ---
$body = @{
    client_id     = $env:CLIENT_ID
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $env:CLIENT_SECRET
    grant_type    = "client_credentials"
}
$token = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$env:TENANT_ID/oauth2/v2.0/token" -Body $body
$headers = @{ Authorization = "Bearer $($token.access_token)" }

# 1. Get YOUR Graph User ID (needed to filter assignments)
$myEmail = "Steven.Brownlow@letselevate.tech"
$userResp = Invoke-RestMethod -Headers $headers -Uri "https://graph.microsoft.com/v1.0/users/$myEmail"
$myId = $userResp.id

# 2. Get Planner Tasks
$targetPlanId = $env:PLAN_ID.Trim()
$planUrl = "https://graph.microsoft.com/v1.0/planner/plans/$targetPlanId/tasks"
$response = Invoke-RestMethod -Headers $headers -Uri $planUrl -Method Get

# 3. Filter for: (Not Complete) AND (Assigned to YOU)
$myTasks = $response.value | Where-Object { 
    $_.percentComplete -lt 100 -and 
    $_.assignments."$myId" -ne $null 
}

# --- Process Tasks ---
$reportItems = @()

foreach ($task in $myTasks) {
    # Use a safe variable name (not $PID)
    $currentTaskId = $task.id
    $title = $task.title
    
    # Check for existing GitHub Issue
    $issue = gh issue list --search "$currentTaskId" --json number,title | ConvertFrom-Json | Select-Object -First 1
    
    if (-not $issue) {
        $issueNumber = gh issue create --title "$title" --body "PlannerID: $currentTaskId `n---`nUpdates:"
        Write-Host "Created new issue for: $title"
    } else {
        $issueNumber = $issue.number
    }

    # Get latest comment
    $issueData = gh issue view $issueNumber --json comments | ConvertFrom-Json
    $latestComment = $issueData.comments | Select-Object -Last 1
    $note = if ($latestComment) { $latestComment.body } else { "No updates recorded in meeting." }

    $reportItems += [PSCustomObject]@{
        Title    = $title
        Percent  = $task.percentComplete
        Notes    = $note
    }
}

# --- Post-Meeting Report ---
if ($env:RUN_MODE -eq "report" -and $reportItems.Count -gt 0) {
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
