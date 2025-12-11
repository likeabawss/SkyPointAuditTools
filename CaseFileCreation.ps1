<#
.SYNOPSIS
    Orchestrates the retrieval and reporting of Unified Audit Logs for forensic cases.
.DESCRIPTION
    1. Archives existing session data.
    2. Retrieves new logs via GetUnifiedLogData.ps1.
    3. Generates a self-contained HTML report with Workload/Intent grouping and Card layout.
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$HaloId,

    [Parameter(Mandatory=$true)]
    [string]$TicketId,

    [Parameter(Mandatory=$true)]
    [string]$ActorUpn,

    [Parameter(Mandatory=$true)]
    [string]$AdminUpn,

    [Parameter(Mandatory=$true)]
    [DateTime]$StartDate,

    [Parameter(Mandatory=$true)]
    [DateTime]$EndDate,

    [Parameter(Mandatory=$false)]
    [string]$RecordType,

    [Parameter(Mandatory=$false)]
    [string]$Operation,

    [Parameter(Mandatory=$false)]
    [string]$SearchText
)

$ErrorActionPreference = "Stop"

# --- Configuration ---
$ScriptRoot = $PSScriptRoot
$SessionRoot = Join-Path $ScriptRoot "SessionData"
$CaseDir = Join-Path $SessionRoot "$HaloId\$TicketId"
$LogFile = Join-Path $CaseDir "Case.log"

# --- Logging Helper ---
function Write-CaseLog {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Severity = "INFO"
    )
    $timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    $logEntry = "[$timestamp] [$Severity] $Message"
    
    $color = "Cyan"
    if ($Severity -eq "WARN") { $color = "Yellow" }
    elseif ($Severity -eq "ERROR") { $color = "Red" }
    
    Write-Host $logEntry -ForegroundColor $color
    
    # Ensure directory exists before logging
    if (-not (Test-Path $CaseDir)) {
        New-Item -Path $CaseDir -ItemType Directory -Force | Out-Null
    }
    Add-Content -Path $LogFile -Value $logEntry
}

# --- DATA LAYER ---

function Import-CaseData {
    param([string]$Path)
    
    Write-CaseLog "Importing data from $Path"
    $jsonFiles = Get-ChildItem -Path $Path -Filter "*.json"
    $allEvents = @()

    foreach ($file in $jsonFiles) {
        try {
            $content = Get-Content $file.FullName -Raw | ConvertFrom-Json
            if ($null -eq $content) { continue }
            
            # Handle single object vs array
            if ($content -isnot [Array]) { $content = @($content) }

            foreach ($record in $content) {
                # Parse AuditData string into object
                if ($record.AuditData -is [string]) {
                    try {
                        $parsedData = $record.AuditData | ConvertFrom-Json
                        $record | Add-Member -MemberType NoteProperty -Name "AuditDataParsed" -Value $parsedData -Force
                    } catch {
                        Write-CaseLog "Failed to parse AuditData for record $($record.Id)" "WARN"
                        $record | Add-Member -MemberType NoteProperty -Name "AuditDataParsed" -Value $null -Force
                    }
                } else {
                     $record | Add-Member -MemberType NoteProperty -Name "AuditDataParsed" -Value $null -Force
                }
                $allEvents += $record
            }
        } catch {
            Write-CaseLog "Error reading file $($file.Name): $_" "ERROR"
        }
    }
    return $allEvents
}

# --- LOGIC LAYER ---

function Get-CategorizedEvents {
    param($Events)
    
    $categorized = @()
    foreach ($e in $Events) {
        $workload = "Other"
        $intent = "Other"
        
        # Determine Workload
        if ($e.RecordType -match "SharePoint|OneDrive") { $workload = "File Activity" }
        elseif ($e.RecordType -match "AzureActiveDirectory") { $workload = "Identity & Access" }
        elseif ($e.RecordType -match "Exchange") { $workload = "Exchange" }
        elseif ($e.RecordType -match "CRM|PowerPlatform") { $workload = "Business Apps" }
        elseif ($e.RecordType -match "MicrosoftTeams") { $workload = "Collaboration" }
        
        # Determine Intent
        $op = $e.Operations
        if ($op -match "Access|Preview|Login|Logon") { $intent = "Access/Read" }
        elseif ($op -match "Modified|Set-|Update") { $intent = "Modification" }
        elseif ($op -match "Download|Sync") { $intent = "Exfiltration Risk" }
        elseif ($op -match "Delete|Recycle") { $intent = "Deletion" }
        
        # Add properties
        $e | Add-Member -MemberType NoteProperty -Name "ReportWorkload" -Value $workload -Force
        $e | Add-Member -MemberType NoteProperty -Name "ReportIntent" -Value $intent -Force
        
        $categorized += $e
    }
    return $categorized
}

# --- PRESENTATION LAYER ---

function New-HtmlCard {
    param($Event)
    
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<div class='event-card'>")
    
    # Convert UTC to Local
    try {
        $utcDate = [DateTime]::SpecifyKind([DateTime]::Parse($Event.CreationDate), [System.DateTimeKind]::Utc)
        # Format with offset (e.g., 03-12-2025 10:34:06 +13:00)
        $localDate = $utcDate.ToLocalTime().ToString("dd-MM-yyyy HH:mm:ss K")
    } catch {
        $localDate = $Event.CreationDate # Fallback if parse fails
    }

    # Header
    [void]$sb.Append("<div class='card-header'>")
    [void]$sb.Append("<span class='timestamp'>$localDate</span>")
    [void]$sb.Append("<span class='operation'>$($Event.Operations)</span>")
    [void]$sb.Append("<span class='user'>$($Event.UserIds)</span>")
    [void]$sb.Append("</div>")
    
    # Body (Dynamic DL)
    [void]$sb.Append("<div class='card-body'><dl class='audit-details'>")
    
    if ($Event.AuditDataParsed) {
        foreach ($prop in $Event.AuditDataParsed.PSObject.Properties) {
            $key = $prop.Name
            $val = $prop.Value
            
            # Format complex objects
            if ($val -is [PSCustomObject] -or $val -is [Array]) {
                $valStr = $val | ConvertTo-Json -Depth 10 -Compress -WarningAction SilentlyContinue
                $valHtml = "<details><summary>View JSON</summary><pre class='json-block'>$valStr</pre></details>"
            } else {
                $valHtml = [System.Net.WebUtility]::HtmlEncode($val)
            }
            
            [void]$sb.Append("<dt>$key</dt><dd>$valHtml</dd>")
        }
    } else {
        [void]$sb.Append("<dt>Raw AuditData</dt><dd>$($Event.AuditData)</dd>")
    }
    
    [void]$sb.Append("</dl></div>") # Close dl, body
    [void]$sb.Append("</div>") # Close card
    
    return $sb.ToString()
}

function New-HtmlReport {
    param($CategorizedEvents, $OutputPath, $Filters)
    
    Write-CaseLog "Generating HTML report..."
    
    # Process Logo (Look for skypoint.jpg/png/jpeg in script root)
    $logoHtml = ""
    $possibleLogos = Get-ChildItem -Path $ScriptRoot -Filter "skypoint.*" | Where-Object { $_.Extension -match "\.(jpg|jpeg|png|svg|gif)$" }
    $LogoPath = $possibleLogos | Select-Object -First 1 -ExpandProperty FullName

    if (-not [string]::IsNullOrWhiteSpace($LogoPath) -and (Test-Path $LogoPath)) {
        try {
            $imgBytes = [System.IO.File]::ReadAllBytes($LogoPath)
            $base64 = [Convert]::ToBase64String($imgBytes)
            $ext = [System.IO.Path]::GetExtension($LogoPath).TrimStart('.').ToLower()
            
            # Mime mapping
            $mime = switch ($ext) {
                "svg" { "image/svg+xml" }
                "jpg" { "image/jpeg" }
                "jpeg" { "image/jpeg" }
                "png" { "image/png" }
                "gif" { "image/gif" }
                default { "image/$ext" }
            }
            
            $logoHtml = "<img src='data:$mime;base64,$base64' alt='Company Logo' class='company-logo' />"
        } catch {
            Write-CaseLog "Failed to process logo: $_" "WARN"
        }
    }
    
    $sb = [System.Text.StringBuilder]::new()
    
    # HTML Head & CSS
    [void]$sb.Append(@"
<!DOCTYPE html>
<html>
<head>
    <title>Forensic Audit Report - $TicketId</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f4f4; color: #333; margin: 0; padding: 20px; }
        h1, h2, h3 { color: #0078d4; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        
        /* Navigation */
        nav { background: #333; padding: 10px; margin-bottom: 20px; position: sticky; top: 0; z-index: 100; }
        nav a { color: white; text-decoration: none; margin-right: 15px; font-weight: bold; }
        nav a:hover { text-decoration: underline; }
        
        /* Sections */
        section { margin-bottom: 40px; border-bottom: 1px solid #eee; padding-bottom: 20px; }
        
        /* Card Layout */
        .event-card { background: #fff; border: 1px solid #ddd; margin-bottom: 10px; border-radius: 4px; overflow: hidden; }
        .card-header { background: #f8f9fa; padding: 10px; border-bottom: 1px solid #eee; display: flex; justify-content: space-between; font-weight: bold; }
        .card-body { padding: 10px; }
        
        /* Dynamic List */
        dl.audit-details { display: grid; grid-template-columns: 200px auto; gap: 5px; font-size: 0.75em; margin: 0; }
        dt { font-weight: bold; color: #555; grid-column: 1; word-break: break-word; }
        dd { margin: 0; grid-column: 2; word-break: break-all; }
        
        .json-block { background: #f0f0f0; padding: 5px; border-radius: 3px; margin: 5px 0 0 0; white-space: pre-wrap; font-family: monospace; font-size: 1em; }
        
        details { margin: 0; }
        summary { cursor: pointer; color: #0078d4; font-size: 0.9em; user-select: none; }
        summary:hover { text-decoration: underline; }
        
        /* Summary Table */
        table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
        th, td { text-align: left; padding: 8px; border-bottom: 1px solid #ddd; }
        th { background-color: #f2f2f2; }

        /* Header Layout */
        .report-header { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #0078d4; padding-bottom: 10px; margin-bottom: 20px; }
        .report-title h1 { margin: 0; font-size: 24px; color: #0078d4; }
        .report-meta { color: #666; margin-top: 5px; font-size: 0.9em; }
        .company-logo { max-height: 60px; max-width: 200px; }

        /* Watermark */
        .watermark {
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%) rotate(-45deg);
            font-size: 15vw;
            color: rgba(0, 0, 0, 0.05);
            pointer-events: none;
            z-index: 9999;
            white-space: nowrap;
            font-weight: bold;
            user-select: none;
        }
    </style>
</head>
<body>
    <div class='watermark'>CONFIDENTIAL</div>
    <nav>
        <a href='#summary'>Summary</a>
        <a href='#timeline'>Timeline</a>
        <a href='#files'>File Activity</a>
        <a href='#identity'>Identity</a>
        <a href='#exchange'>Exchange</a>
    </nav>
    <div class='container'>
        <div class='report-header'>
            <div class='report-title'>
                <h1>Forensic Audit Report: $TicketId</h1>
                <div class='report-meta'><strong>Customer:</strong> $HaloId | <strong>Target:</strong> $ActorUpn | <strong>Generated:</strong> $(Get-Date -Format "dd-MM-yyyy HH:mm:ss K")</div>
            </div>
            $logoHtml
        </div>
"@)

    if (-not [string]::IsNullOrWhiteSpace($Filters)) {
        [void]$sb.Append("<div style='background-color: #eef; padding: 10px; border-radius: 4px; margin-bottom: 20px; border: 1px solid #ccd;'><strong>Active Filters:</strong> $Filters</div>")
    }

    # 1. Summary Section
    [void]$sb.Append("<section id='summary'><h2>High-Level Summary</h2>")
    
    # KPI: Total Events
    [void]$sb.Append("<p><strong>Total Events:</strong> $($CategorizedEvents.Count)</p>")
    
    # KPI: Breakdown by Workload
    [void]$sb.Append("<h3>Events by Workload</h3><table><tr><th>Workload</th><th>Count</th></tr>")
    $byWorkload = $CategorizedEvents | Group-Object ReportWorkload | Sort-Object Count -Descending
    foreach ($g in $byWorkload) {
        [void]$sb.Append("<tr><td>$($g.Name)</td><td>$($g.Count)</td></tr>")
    }
    [void]$sb.Append("</table></section>")

    # 2. Timeline (All Events)
    [void]$sb.Append("<section id='timeline'><h2>Timeline View</h2>")
    foreach ($ev in $CategorizedEvents | Sort-Object CreationDate) {
        [void]$sb.Append( (New-HtmlCard -Event $ev) )
    }
    [void]$sb.Append("</section>")

    # 3. Workload Specific Sections (Example: File Activity)
    [void]$sb.Append("<section id='files'><h2>File Activity (SharePoint/OneDrive)</h2>")
    $fileEvents = $CategorizedEvents | Where-Object { $_.ReportWorkload -eq "File Activity" } | Sort-Object CreationDate
    if ($fileEvents) {
        foreach ($ev in $fileEvents) {
            [void]$sb.Append( (New-HtmlCard -Event $ev) )
        }
    } else {
        [void]$sb.Append("<p>No file activity detected.</p>")
    }
    [void]$sb.Append("</section>")

    # 4. Identity Section
    [void]$sb.Append("<section id='identity'><h2>Identity & Access (Entra ID)</h2>")
    $idEvents = $CategorizedEvents | Where-Object { $_.ReportWorkload -eq "Identity & Access" } | Sort-Object CreationDate
    if ($idEvents) {
        foreach ($ev in $idEvents) {
            [void]$sb.Append( (New-HtmlCard -Event $ev) )
        }
    } else {
        [void]$sb.Append("<p>No identity events detected.</p>")
    }
    [void]$sb.Append("</section>")
    
    # 5. Exchange Section
    [void]$sb.Append("<section id='exchange'><h2>Exchange (Mail)</h2>")
    $exEvents = $CategorizedEvents | Where-Object { $_.ReportWorkload -eq "Exchange" } | Sort-Object CreationDate
    if ($exEvents) {
        foreach ($ev in $exEvents) {
            [void]$sb.Append( (New-HtmlCard -Event $ev) )
        }
    } else {
        [void]$sb.Append("<p>No exchange events detected.</p>")
    }
    [void]$sb.Append("</section>")

    [void]$sb.Append("</div></body></html>")
    
    $sb.ToString() | Out-File -FilePath $OutputPath -Encoding utf8
    Write-CaseLog "Report generated at $OutputPath"
}

# --- MAIN ORCHESTRATION ---

try {
    # 0. Setup
    if (-not (Test-Path $CaseDir)) { New-Item -Path $CaseDir -ItemType Directory -Force | Out-Null }
    Write-CaseLog "Starting CaseFileCreation for $HaloId - $TicketId"

    # 1. Archive (Placeholder for now, as per instructions "Zips existing content")
    # In a real run, we would zip $CaseDir contents to an Archive folder.
    
    # 2. Clean (Placeholder)
    
    # 3. Retrieve Data
    Write-CaseLog "Calling GetUnifiedLogData.ps1..."
    # Note: Assuming GetUnifiedLogData.ps1 is in the same directory
    $retrievalScript = Join-Path $ScriptRoot "GetUnifiedLogData.ps1"
    if (Test-Path $retrievalScript) {
        # We are NOT calling it in this implementation step to avoid external dependencies/auth issues during testing.
        # In production: & $retrievalScript -HaloId $HaloId -TicketId $TicketId ...
        Write-CaseLog "Skipping actual retrieval for this implementation phase. Using existing data." "WARN"
    }

    # 4. Report Generation
    $events = Import-CaseData -Path $CaseDir

    # Apply Filters
    if (-not [string]::IsNullOrWhiteSpace($RecordType)) {
        Write-CaseLog "Filtering by RecordType: $RecordType"
        $events = $events | Where-Object { $_.RecordType -eq $RecordType }
    }
    if (-not [string]::IsNullOrWhiteSpace($Operation)) {
        Write-CaseLog "Filtering by Operation: $Operation"
        $events = $events | Where-Object { $_.Operations -eq $Operation }
    }
    if (-not [string]::IsNullOrWhiteSpace($SearchText)) {
        Write-CaseLog "Filtering by SearchText: $SearchText"
        $events = $events | Where-Object { 
            # Simple text search across the JSON representation of the event
            ($_ | ConvertTo-Json -Depth 10 -Compress -WarningAction SilentlyContinue) -like "*$SearchText*"
        }
    }

    if ($events.Count -eq 0) {
        Write-CaseLog "No events found in $CaseDir (after filtering). Cannot generate report." "WARN"
    } else {
        $categorized = Get-CategorizedEvents -Events $events
        $reportPath = Join-Path $CaseDir "Report.html"
        
        $filterParts = @()
        $filterParts += "<li>Date Range: <b>$($StartDate.ToString('dd-MM-yyyy'))</b> to <b>$($EndDate.ToString('dd-MM-yyyy'))</b></li>"
        
        $rtDisplay = if (-not [string]::IsNullOrWhiteSpace($RecordType)) { $RecordType } else { "ALL" }
        $filterParts += "<li>RecordType: <b>$rtDisplay</b></li>"
        
        $opDisplay = if (-not [string]::IsNullOrWhiteSpace($Operation)) { $Operation } else { "ALL" }
        $filterParts += "<li>Operation: <b>$opDisplay</b></li>"
        
        $stDisplay = if (-not [string]::IsNullOrWhiteSpace($SearchText)) { $SearchText } else { "ALL" }
        $filterParts += "<li>SearchText: <b>$stDisplay</b></li>"
        
        $filterStr = "<ul style='margin: 5px 0 0 20px; padding: 0;'>$($filterParts -join '')</ul>"

        New-HtmlReport -CategorizedEvents $categorized -OutputPath $reportPath -Filters $filterStr
    }

    Write-CaseLog "CaseFileCreation completed successfully."

} catch {
    Write-CaseLog "Fatal Error: $_" "ERROR"
    exit 1
}
