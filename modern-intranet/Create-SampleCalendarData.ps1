<#
.SYNOPSIS
    Modern Intranet — Sample Calendar Data Generation
.DESCRIPTION
    Populates a SharePoint Calendar list with 25 realistic events for demonstration.
    Connects using Web Login (interactive browser).
.NOTES
    Requires: PnP.PowerShell module
#>

# ============================================================
# CONFIGURATION
# ============================================================
$ErrorActionPreference = "Stop"
$siteUrl = "https://devtenant0424.sharepoint.com/sites/DEVSITE" 
$listName = "CalendarEvents"

# ============================================================
# CONNECT TO SPO
# ============================================================
Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Modern Intranet — Sample Data Generator" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

try {
    Write-Host "Connecting to $siteUrl..." -ForegroundColor Yellow
    Connect-PnPOnline -Url $siteUrl -UseWebLogin
    Write-Host "Connected successfully!" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ============================================================
# GENERATE DATA
# ============================================================
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "Error: List '$listName' not found. Please create it first using Provision-Lists.ps1." -ForegroundColor Red
    exit 1
}

Write-Host "Generating 25 sample events for list '$listName'..." -ForegroundColor Cyan

$eventTitles = @(
    "Quarterly Town Hall", "Team Sync: Product Roadmap", "HR Training: Modern Workplace",
    "Customer Support Workshop", "Project Alpha: Kickoff", "Innovation Lab Brainstorming",
    "Monthly Budget Review", "Security Awareness Webinar", "Wellness Wednesday: Yoga",
    "IT Infrastructure Maintenance", "Design Review: Mobile App", "Marketing Campaign Launch",
    "Social Event: Pizza Friday", "Feedback Session with CEO", "Data Backup Operations"
)

$locations = @("Conference Room A", "Microsoft Teams", "Main Lobby", "Room 402", "External Venue")

$baseDate = Get-Date

for ($i = 0; $i -lt 25; $i++) {
    # Distribute events within -30 and +60 days from today
    $daysOffset = Get-Random -Minimum -15 -Maximum 45
    $hoursOffset = Get-Random -Minimum 8 -Maximum 17
    
    $startDate = $baseDate.AddDays($daysOffset)
    $startDate = Get-Date -Year $startDate.Year -Month $startDate.Month -Day $startDate.Day -Hour $hoursOffset -Minute 0 -Second 0
    
    $durationHours = Get-Random -Minimum 1 -Maximum 4
    $endDate = $startDate.AddHours($durationHours)
    
    $title = $eventTitles | Get-Random
    $location = $locations | Get-Random
    
    Write-Host "  Adding: $title ($($startDate.ToString('MMM dd, hh:mm tt')))" -ForegroundColor Gray
    
    Add-PnPListItem -List $listName -Values @{
        "Title" = $title;
        "EventDate" = $startDate;
        "EndDate" = $endDate;
        "Location" = $location;
    } | Out-Null
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  25 sample events created successfully!" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
