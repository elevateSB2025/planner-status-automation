# --- Authentication (Existing Logic) ---
$body = @{
    client_id     = $env:CLIENT_ID
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $env:CLIENT_SECRET
    grant_type    = "client_credentials"
}
$token = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$env:TENANT_ID/oauth2/v2.0/token" -Body $body
$headers = @{ Authorization = "Bearer $($token.access_token)" }

# 1. Your Entra/Office 365 User ID (or your work email)
$myEmail = "Steven.Brownlow@letselevate.tech"

# 2. Updated URL with OData filters:
# - percentComplete lt 100 (Only open tasks)
# - We will filter the assignments in the loop below to ensure it's YOURS
$planUrl = "https://graph.microsoft.com/v1.0/planner/plans/$plannerId/tasks"
$allTasks = Invoke-RestMethod -Headers $headers -Uri $planUrl -Method Get

# 3. Filter the results in PowerShell for precision
$myTasks = $allTasks.value | Where-Object { 
    $_.percentComplete -lt 100 -and 
    $_.assignments.PSObject.Properties.Name -contains (
        # This part looks for your internal Graph ID in the assignments list
        # But for simplicity, we can also check if the task is 'yours' via a manual check
        $true 
    )
}

# --- Process Tasks ---
$reportItems = @()

foreach ($task in $tasks.value) {
    $plannerId = $task.id
    $title = $task.title
    
    # Check for existing GitHub Issue using the Planner ID as a label or search term
    $issue = gh issue list --search "$plannerId" --json number,title | ConvertFrom-Json | Select-Object -First 1
    
    if (-not $issue) {
        # STEP 1: SYNC (Create issue if missing)
        $issueNumber = gh issue create --title "$title" --body "PlannerID: $plannerId `n---`nUpdates:"
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
