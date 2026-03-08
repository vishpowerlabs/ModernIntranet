<#
.SYNOPSIS
    Modern Intranet — SharePoint List Provisioning Script
.DESCRIPTION
    Creates SharePoint Online lists with correct column types for each
    Modern Intranet web part. Connects using Web Login (interactive browser).
    User selects which web part list(s) to create and provides custom list names.
.NOTES
    Requires: PnP.PowerShell module
    Install:  Install-Module -Name PnP.PowerShell -Scope CurrentUser
#>

# ============================================================
# CONFIGURATION
# ============================================================
$ErrorActionPreference = "Stop"

# ============================================================
# CONNECT TO SPO
# ============================================================
Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Modern Intranet — List Provisioning Tool" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

$siteUrl = "https://devtenant0424.sharepoint.com/sites/DEVSITE" 
Write-Host ""
Write-Host "Connecting to SharePoint Online (browser login)..." -ForegroundColor Yellow

try {
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
    Write-Host "Connected successfully!" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ============================================================
# HELPER FUNCTIONS
# ============================================================
function Add-ImageColumn {
    param(
        [string]$ListName,
        [string]$DisplayName,
        [string]$InternalName
    )
    $fieldXml = "<Field Type='Thumbnail' Name='$InternalName' StaticName='$InternalName' DisplayName='$DisplayName' ID='{$([guid]::NewGuid())}' />"
    Add-PnPFieldFromXml -List $ListName -FieldXml $fieldXml
    
    # Add to default view
    $view = Get-PnPView -List $ListName | Where-Object { $_.DefaultView -eq $true }
    $view.ViewFields.Add($InternalName)
    $view.Update()
    Invoke-PnPQuery
    
    Write-Host "    + $DisplayName (Image/Thumbnail)" -ForegroundColor Gray
}

function Add-TextField {
    param(
        [string]$ListName,
        [string]$DisplayName,
        [string]$InternalName,
        [switch]$Required
    )
    $params = @{
        List        = $ListName
        DisplayName = $DisplayName
        InternalName = $InternalName
        Type        = "Text"
        AddToDefaultView = $true
    }
    if ($Required) { $params.Required = $true }
    Add-PnPField @params | Out-Null
    Write-Host "    + $DisplayName (Text$(if($Required){' *'}))" -ForegroundColor Gray
}

function Add-NoteField {
    param(
        [string]$ListName,
        [string]$DisplayName,
        [string]$InternalName,
        [switch]$Required
    )
    $params = @{
        List        = $ListName
        DisplayName = $DisplayName
        InternalName = $InternalName
        Type        = "Note"
        AddToDefaultView = $true
    }
    if ($Required) { $params.Required = $true }
    Add-PnPField @params | Out-Null
    Write-Host "    + $DisplayName (Note/Multi-line$(if($Required){' *'}))" -ForegroundColor Gray
}

function Add-UrlField {
    param(
        [string]$ListName,
        [string]$DisplayName,
        [string]$InternalName
    )
    Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName -Type URL -AddToDefaultView | Out-Null
    Write-Host "    + $DisplayName (Hyperlink/URL)" -ForegroundColor Gray
}

function Add-BooleanField {
    param(
        [string]$ListName,
        [string]$DisplayName,
        [string]$InternalName,
        [string]$DefaultValue = "1"
    )
    $fieldXml = "<Field Type='Boolean' Name='$InternalName' StaticName='$InternalName' DisplayName='$DisplayName' ID='{$([guid]::NewGuid())}' ><Default>$DefaultValue</Default></Field>"
    Add-PnPFieldFromXml -List $ListName -FieldXml $fieldXml
    $view = Get-PnPView -List $ListName | Where-Object { $_.DefaultView -eq $true }
    $view.ViewFields.Add($InternalName)
    $view.Update()
    Invoke-PnPQuery
    Write-Host "    + $DisplayName (Yes/No, default=$DefaultValue)" -ForegroundColor Gray
}

function Add-DateTimeField {
    param(
        [string]$ListName,
        [string]$DisplayName,
        [string]$InternalName,
        [switch]$Required
    )
    $params = @{
        List        = $ListName
        DisplayName = $DisplayName
        InternalName = $InternalName
        Type        = "DateTime"
        AddToDefaultView = $true
    }
    if ($Required) { $params.Required = $true }
    Add-PnPField @params | Out-Null
    Write-Host "    + $DisplayName (DateTime$(if($Required){' *'}))" -ForegroundColor Gray
}

function Add-ChoiceField {
    param(
        [string]$ListName,
        [string]$DisplayName,
        [string]$InternalName,
        [string[]]$Choices,
        [string]$DefaultValue = ""
    )
    Add-PnPField -List $ListName -DisplayName $DisplayName -InternalName $InternalName -Type Choice -Choices $Choices -AddToDefaultView | Out-Null
    if ($DefaultValue -ne "") {
        Set-PnPField -List $ListName -Identity $InternalName -Values @{DefaultValue = $DefaultValue }
    }
    Write-Host "    + $DisplayName (Choice: $($Choices -join ', '))" -ForegroundColor Gray
}

function Add-NumberField {
    param(
        [string]$ListName,
        [string]$DisplayName,
        [string]$InternalName,
        [switch]$Required
    )
    $params = @{
        List        = $ListName
        DisplayName = $DisplayName
        InternalName = $InternalName
        Type        = "Number"
        AddToDefaultView = $true
    }
    if ($Required) { $params.Required = $true }
    Add-PnPField @params | Out-Null
    Write-Host "    + $DisplayName (Number$(if($Required){' *'}))" -ForegroundColor Gray
}

function Add-PersonField {
    param(
        [string]$ListName,
        [string]$DisplayName,
        [string]$InternalName,
        [switch]$Required
    )
    $params = @{
        List        = $ListName
        DisplayName = $DisplayName
        InternalName = $InternalName
        Type        = "User"
        AddToDefaultView = $true
    }
    if ($Required) { $params.Required = $true }
    Add-PnPField @params | Out-Null
    Write-Host "    + $DisplayName (Person/User$(if($Required){' *'}))" -ForegroundColor Gray
}

function Add-LookupField {
    param(
        [string]$ListName,
        [string]$DisplayName,
        [string]$InternalName,
        [string]$LookupListName,
        [string]$LookupField = "Title"
    )
    $lookupList = Get-PnPList -Identity $LookupListName
    $fieldXml = "<Field Type='Lookup' Name='$InternalName' StaticName='$InternalName' DisplayName='$DisplayName' ID='{$([guid]::NewGuid())}' List='{$($lookupList.Id)}' ShowField='$LookupField' />"
    Add-PnPFieldFromXml -List $ListName -FieldXml $fieldXml
    $view = Get-PnPView -List $ListName | Where-Object { $_.DefaultView -eq $true }
    $view.ViewFields.Add($InternalName)
    $view.Update()
    Invoke-PnPQuery
    Write-Host "    + $DisplayName (Lookup → $LookupListName.$LookupField)" -ForegroundColor Gray
}

function Create-ListIfNotExists {
    param(
        [string]$ListName,
        [string]$Description,
        [string]$Template = "GenericList"
    )
    $existingList = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($existingList) {
        Write-Host "  List '$ListName' already exists. Skipping creation." -ForegroundColor Yellow
        return $false
    }
    New-PnPList -Title $ListName -Template $Template -Url "Lists/$($ListName -replace ' ','')" | Out-Null
    Write-Host "  Created list: $ListName (Template: $Template)" -ForegroundColor Green
    return $true
}

# ============================================================
# LIST CREATION FUNCTIONS (one per web part)
# ============================================================

function Create-BannerSliderList {
    param([string]$ListName)
    Write-Host ""
    Write-Host "Creating Banner Slider list..." -ForegroundColor Cyan
    $created = Create-ListIfNotExists -ListName $ListName -Description "Banner slides for the intranet carousel"
    if (-not $created) { return }

    # Title column already exists by default
    Write-Host "    + Title (Text) — default column" -ForegroundColor Gray
    Add-TextField      -ListName $ListName -DisplayName "Description"  -InternalName "BannerDescription"
    Add-ImageColumn    -ListName $ListName -DisplayName "Banner Image" -InternalName "BannerImage"
    Add-BooleanField   -ListName $ListName -DisplayName "Active"       -InternalName "BannerActive" -DefaultValue "1"
    Add-TextField      -ListName $ListName -DisplayName "Button Text"  -InternalName "ButtonText"
    Add-UrlField       -ListName $ListName -DisplayName "Page Link"    -InternalName "PageLink"
    Add-NumberField    -ListName $ListName -DisplayName "Sort Order"   -InternalName "SortOrder"

    Write-Host "  Done! Banner Slider list created with 7 columns." -ForegroundColor Green
}

function Create-QuickLinksList {
    param([string]$ListName)
    Write-Host ""
    Write-Host "Creating Quick Links list..." -ForegroundColor Cyan
    $created = Create-ListIfNotExists -ListName $ListName -Description "Quick link tiles for intranet"
    if (-not $created) { return }

    Write-Host "    + Title (Text) — default column" -ForegroundColor Gray
    Add-UrlField       -ListName $ListName -DisplayName "Link URL"     -InternalName "LinkUrl"
    Add-TextField      -ListName $ListName -DisplayName "Icon Name"    -InternalName "IconName"
    Add-BooleanField   -ListName $ListName -DisplayName "Pinned"       -InternalName "Pinned"       -DefaultValue "0"
    Add-NumberField    -ListName $ListName -DisplayName "Sort Order"   -InternalName "SortOrder"

    Write-Host "  Done! Quick Links list created with 5 columns." -ForegroundColor Green
}

function Create-HighlightsList {
    param([string]$ListName)
    Write-Host ""
    Write-Host "Creating Highlights list..." -ForegroundColor Cyan
    $created = Create-ListIfNotExists -ListName $ListName -Description "Highlight cards for intranet"
    if (-not $created) { return }

    Write-Host "    + Title (Text) — default column" -ForegroundColor Gray
    Add-NoteField      -ListName $ListName -DisplayName "Description"     -InternalName "HighlightDescription" -Required
    Add-ImageColumn    -ListName $ListName -DisplayName "Banner Image"    -InternalName "HighlightImage"
    Add-UrlField       -ListName $ListName -DisplayName "Detail Page Link" -InternalName "DetailPageLink"
    Add-BooleanField   -ListName $ListName -DisplayName "Pinned"           -InternalName "Pinned"           -DefaultValue "0"

    Write-Host "  Done! Highlights list created with 5 columns." -ForegroundColor Green
}

function Create-EventsList {
    param([string]$ListName)
    Write-Host ""
    Write-Host "Creating Events list..." -ForegroundColor Cyan
    $created = Create-ListIfNotExists -ListName $ListName -Description "Events for intranet calendar and events web part"
    if (-not $created) { return }

    Write-Host "    + Title (Text) — default column" -ForegroundColor Gray
    Add-DateTimeField  -ListName $ListName -DisplayName "Event Date"   -InternalName "EventDate" -Required
    Add-ImageColumn    -ListName $ListName -DisplayName "Event Image"  -InternalName "EventImage"
    Add-UrlField       -ListName $ListName -DisplayName "Event Link"   -InternalName "EventLink"
    Add-TextField      -ListName $ListName -DisplayName "Location"     -InternalName "EventLocation"
    Add-BooleanField   -ListName $ListName -DisplayName "Active"       -InternalName "EventActive"   -DefaultValue "1"
    Add-BooleanField   -ListName $ListName -DisplayName "Pinned"       -InternalName "Pinned"        -DefaultValue "0"

    Write-Host "  Done! Events list created with 7 columns." -ForegroundColor Green
}

function Create-UpcomingMeetingList {
    param([string]$ListName)
    Write-Host ""
    Write-Host "Creating Upcoming Meeting list..." -ForegroundColor Cyan
    $created = Create-ListIfNotExists -ListName $ListName -Description "Meeting entries for upcoming meeting web part (SP List mode)"
    if (-not $created) { return }

    Write-Host "    + Title (Text) — default column (meeting subject)" -ForegroundColor Gray
    Add-DateTimeField  -ListName $ListName -DisplayName "Start Date"    -InternalName "MeetingStart" -Required
    Add-DateTimeField  -ListName $ListName -DisplayName "End Date"      -InternalName "MeetingEnd"
    Add-UrlField       -ListName $ListName -DisplayName "Meeting Link"  -InternalName "MeetingLink"
    Add-PersonField    -ListName $ListName -DisplayName "Organizer"     -InternalName "MeetingOrganizer"

    Write-Host "  Done! Upcoming Meeting list created with 5 columns." -ForegroundColor Green
}

function Create-RecentDocumentsLibrary {
    param([string]$ListName)
    Write-Host ""
    Write-Host "Note: Recent Documents web part can read from any existing Document Library." -ForegroundColor Yellow
    Write-Host "No custom list creation is needed — just point the web part to your library." -ForegroundColor Yellow
    Write-Host "Skipping." -ForegroundColor Yellow
}

function Create-CalendarList {
    param([string]$ListName)
    Write-Host ""
    Write-Host "Creating Modern Calendar list (Events Template)..." -ForegroundColor Cyan
    Write-Host "Note: This creates a standard SharePoint Calendar list (Template 106)." -ForegroundColor Yellow
    $created = Create-ListIfNotExists -ListName $ListName -Description "Calendar events for intranet" -Template "Events"
    if (-not $created) { return }

    # Built-in Calendar columns:
    # Title (Event Name)
    # EventDate (Start Time)
    # EndDate (End Time)
    # Location (Text)
    Write-Host "    + Title (Text) — built-in" -ForegroundColor Gray
    Write-Host "    + EventDate (Start Time) — built-in" -ForegroundColor Gray
    Write-Host "    + EndDate (End Time) — built-in" -ForegroundColor Gray
    Write-Host "    + Location (Text) — built-in" -ForegroundColor Gray

    Write-Host "  Done! Modern Calendar list created (Standard Template)." -ForegroundColor Green
}

function Create-EmployeeDirectoryList {
    param([string]$ListName)
    Write-Host ""
    Write-Host "Creating Employee Directory list..." -ForegroundColor Cyan
    Write-Host "Note: The Employee Directory web part can also use Graph API for Azure AD. This list is for SP List mode." -ForegroundColor Yellow
    $created = Create-ListIfNotExists -ListName $ListName -Description "Employee directory for intranet (SP List mode)"
    if (-not $created) { return }

    # Title column will be used as Employee Name
    Write-Host "    + Title (Text) — default column (Employee Name)" -ForegroundColor Gray
    Add-TextField      -ListName $ListName -DisplayName "Job Title"    -InternalName "EmpJobTitle" -Required
    Add-TextField      -ListName $ListName -DisplayName "Department"   -InternalName "EmpDepartment" -Required
    Add-TextField      -ListName $ListName -DisplayName "Location"     -InternalName "EmpLocation"
    Add-TextField      -ListName $ListName -DisplayName "Email"        -InternalName "EmpEmail" -Required
    Add-TextField      -ListName $ListName -DisplayName "Phone"        -InternalName "EmpPhone"
    Add-PersonField    -ListName $ListName -DisplayName "Manager"      -InternalName "EmpManager"
    Add-NoteField      -ListName $ListName -DisplayName "Projects"     -InternalName "EmpProjects"
    Add-NoteField      -ListName $ListName -DisplayName "About Me"     -InternalName "EmpAboutMe"
    Add-NoteField      -ListName $ListName -DisplayName "Interests"    -InternalName "EmpInterests"
    Add-NoteField      -ListName $ListName -DisplayName "Skills"       -InternalName "EmpSkills"

    Write-Host "  Done! Employee Directory list created with 11 columns." -ForegroundColor Green
}

function Create-PollLists {
    param([string]$PollsListName, [string]$VotesListName)
    Write-Host ""
    Write-Host "Creating Poll lists (2 lists required)..." -ForegroundColor Cyan

    # --- Polls List ---
    Write-Host ""
    Write-Host "  [1/2] Polls list..." -ForegroundColor White
    $created = Create-ListIfNotExists -ListName $PollsListName -Description "Poll questions for intranet"
    if ($created) {
        # Title column used as Question
        Write-Host "    + Title (Text) — default column (Poll Question)" -ForegroundColor Gray
        Add-NoteField      -ListName $PollsListName -DisplayName "Options"  -InternalName "PollOptions" -Required
        Write-Host "      (stores JSON array, e.g. [""Option A"",""Option B"",""Option C""])" -ForegroundColor DarkGray
        Add-ChoiceField    -ListName $PollsListName -DisplayName "Status"   -InternalName "PollStatus" -Choices "Active","Closed" -DefaultValue "Active"
        Write-Host "  Done! Polls list created with 3 columns." -ForegroundColor Green
    }

    # --- Votes List ---
    Write-Host ""
    Write-Host "  [2/2] Votes list..." -ForegroundColor White
    $created = Create-ListIfNotExists -ListName $VotesListName -Description "Individual vote records"
    if ($created) {
        # Title column can be hidden or used for reference
        Write-Host "    + Title (Text) — default column (can be auto-generated)" -ForegroundColor Gray
        Add-LookupField    -ListName $VotesListName -DisplayName "Poll"            -InternalName "PollId" -LookupListName $PollsListName -LookupField "Title"
        Add-PersonField    -ListName $VotesListName -DisplayName "Voter"           -InternalName "VoteUser" -Required
        Add-NumberField    -ListName $VotesListName -DisplayName "Selected Option" -InternalName "SelectedOption" -Required
        Write-Host "      (stores the index of the selected option, 0-based)" -ForegroundColor DarkGray
        Write-Host "  Done! Votes list created with 4 columns." -ForegroundColor Green
    }
}

function Create-ShoutoutLists {
    param([string]$ShoutoutsListName, [string]$LikesListName)
    Write-Host ""
    Write-Host "Creating Shoutout lists (2 lists required)..." -ForegroundColor Cyan

    # --- Shoutouts List ---
    Write-Host ""
    Write-Host "  [1/2] Shoutouts list..." -ForegroundColor White
    $created = Create-ListIfNotExists -ListName $ShoutoutsListName -Description "Peer recognition shoutouts"
    if ($created) {
        Write-Host "    + Title (Text) — default column (can store auto-generated label)" -ForegroundColor Gray
        Add-PersonField    -ListName $ShoutoutsListName -DisplayName "Sender"          -InternalName "ShoutSender" -Required
        Add-TextField      -ListName $ShoutoutsListName -DisplayName "Sender Email"    -InternalName "ShoutSenderEmail"
        Add-PersonField    -ListName $ShoutoutsListName -DisplayName "Recipient"       -InternalName "ShoutRecipient" -Required
        Add-TextField      -ListName $ShoutoutsListName -DisplayName "Recipient Email" -InternalName "ShoutRecipientEmail"
        Add-NoteField      -ListName $ShoutoutsListName -DisplayName "Message"         -InternalName "ShoutMessage" -Required
        Write-Host "  Done! Shoutouts list created with 6 columns." -ForegroundColor Green
    }

    # --- Likes List ---
    Write-Host ""
    Write-Host "  [2/2] Likes list..." -ForegroundColor White
    $created = Create-ListIfNotExists -ListName $LikesListName -Description "Like records for shoutouts"
    if ($created) {
        Write-Host "    + Title (Text) — default column (can be auto-generated)" -ForegroundColor Gray
        Add-LookupField    -ListName $LikesListName -DisplayName "Shoutout"  -InternalName "ShoutoutId" -LookupListName $ShoutoutsListName -LookupField "Title"
        Add-PersonField    -ListName $LikesListName -DisplayName "Liked By"  -InternalName "LikedByUser" -Required
        Write-Host "  Done! Likes list created with 3 columns." -ForegroundColor Green
    }
}

function Create-FaqList {
    param([string]$ListName)
    Write-Host ""
    Write-Host "Creating FAQ list..." -ForegroundColor Cyan
    $created = Create-ListIfNotExists -ListName $ListName -Description "Frequently asked questions for intranet"
    if (-not $created) { return }

    # Title column will be used for Question if preferred, but we'll add explicit ones
    Write-Host "    + Title (Text) — default column" -ForegroundColor Gray
    Add-TextField      -ListName $ListName -DisplayName "Question"     -InternalName "FaqQuestion" -Required
    Add-NoteField      -ListName $ListName -DisplayName "Answer"       -InternalName "FaqAnswer" -Required
    Add-TextField      -ListName $ListName -DisplayName "Category"     -InternalName "FaqCategory"
    Add-NumberField    -ListName $ListName -DisplayName "Sort Order"   -InternalName "SortOrder"

    Write-Host "  Done! FAQ list created with 5 columns." -ForegroundColor Green
}

# ============================================================
# MAIN MENU
# ============================================================
function Show-Menu {
    Write-Host ""
    Write-Host "=============================================" -ForegroundColor Cyan
    Write-Host "  Select a web part list to create:" -ForegroundColor Cyan
    Write-Host "=============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   1.  Banner Slider          (1 list: banners with image, title, CTA)" -ForegroundColor White
    Write-Host "   2.  Quick Links            (1 list: link tiles with icon)" -ForegroundColor White
    Write-Host "   3.  Highlights             (1 list: cards with image, description, link)" -ForegroundColor White
    Write-Host "   4.  Events                 (1 list: events with date, image, location)" -ForegroundColor White
    Write-Host "   5.  Upcoming Meeting       (1 list: meetings with date, link, organizer)" -ForegroundColor White
    Write-Host "   6.  Recent Documents       (no list needed — uses existing doc library)" -ForegroundColor DarkGray
    Write-Host "   7.  Modern Calendar        (1 list: events with start/end date, location)" -ForegroundColor White
    Write-Host "   8.  Employee Directory     (1 list: employees with dept, email, phone)" -ForegroundColor White
    Write-Host "  10.  Shoutouts              (2 lists: shoutouts + likes)" -ForegroundColor White
    Write-Host "  11.  FAQ                    (1 list: questions + answers)" -ForegroundColor White
    Write-Host ""
    Write-Host "  12.  Create ALL lists       (creates all of the above at once)" -ForegroundColor Yellow
    Write-Host "   0.  Exit" -ForegroundColor Red
    Write-Host ""
}

function Run-Selection {
    param([int]$Choice)

    switch ($Choice) {
        1 {
            $name = Read-Host "Enter list name for Banner Slider (default: Banners)"
            if ([string]::IsNullOrWhiteSpace($name)) { $name = "Banners" }
            Create-BannerSliderList -ListName $name
        }
        2 {
            $name = Read-Host "Enter list name for Quick Links (default: QuickLinks)"
            if ([string]::IsNullOrWhiteSpace($name)) { $name = "QuickLinks" }
            Create-QuickLinksList -ListName $name
        }
        3 {
            $name = Read-Host "Enter list name for Highlights (default: Highlights)"
            if ([string]::IsNullOrWhiteSpace($name)) { $name = "Highlights" }
            Create-HighlightsList -ListName $name
        }
        4 {
            $name = Read-Host "Enter list name for Events (default: Events)"
            if ([string]::IsNullOrWhiteSpace($name)) { $name = "Events" }
            Create-EventsList -ListName $name
        }
        5 {
            $name = Read-Host "Enter list name for Upcoming Meeting (default: Meetings)"
            if ([string]::IsNullOrWhiteSpace($name)) { $name = "Meetings" }
            Create-UpcomingMeetingList -ListName $name
        }
        6 {
            Create-RecentDocumentsLibrary -ListName ""
        }
        7 {
            $name = Read-Host "Enter list name for Modern Calendar (default: CalendarEvents)"
            if ([string]::IsNullOrWhiteSpace($name)) { $name = "CalendarEvents" }
            Create-CalendarList -ListName $name
        }
        8 {
            $name = Read-Host "Enter list name for Employee Directory (default: EmployeeDirectory)"
            if ([string]::IsNullOrWhiteSpace($name)) { $name = "EmployeeDirectory" }
            Create-EmployeeDirectoryList -ListName $name
        }
        9 {
            $pollsName = Read-Host "Enter list name for Polls (default: Polls)"
            if ([string]::IsNullOrWhiteSpace($pollsName)) { $pollsName = "Polls" }
            $votesName = Read-Host "Enter list name for Votes (default: PollVotes)"
            if ([string]::IsNullOrWhiteSpace($votesName)) { $votesName = "PollVotes" }
            Create-PollLists -PollsListName $pollsName -VotesListName $votesName
        }
        10 {
            $shoutoutsName = Read-Host "Enter list name for Shoutouts (default: Shoutouts)"
            if ([string]::IsNullOrWhiteSpace($shoutoutsName)) { $shoutoutsName = "Shoutouts" }
            $likesName = Read-Host "Enter list name for Likes (default: ShoutoutLikes)"
            if ([string]::IsNullOrWhiteSpace($likesName)) { $likesName = "ShoutoutLikes" }
            Create-ShoutoutLists -ShoutoutsListName $shoutoutsName -LikesListName $likesName
        }
        11 {
            $name = Read-Host "Enter list name for FAQ (default: FAQ)"
            if ([string]::IsNullOrWhiteSpace($name)) { $name = "FAQ" }
            Create-FaqList -ListName $name
        }
        12 {
            Write-Host ""
            Write-Host "Creating ALL lists with default names..." -ForegroundColor Yellow
            Write-Host "(Press Enter to accept defaults or type a custom name)" -ForegroundColor DarkGray
            Write-Host ""

            $n1 = Read-Host "Banner Slider list name (default: Banners)"
            if ([string]::IsNullOrWhiteSpace($n1)) { $n1 = "Banners" }
            Create-BannerSliderList -ListName $n1

            $n2 = Read-Host "Quick Links list name (default: QuickLinks)"
            if ([string]::IsNullOrWhiteSpace($n2)) { $n2 = "QuickLinks" }
            Create-QuickLinksList -ListName $n2

            $n3 = Read-Host "Highlights list name (default: Highlights)"
            if ([string]::IsNullOrWhiteSpace($n3)) { $n3 = "Highlights" }
            Create-HighlightsList -ListName $n3

            $n4 = Read-Host "Events list name (default: Events)"
            if ([string]::IsNullOrWhiteSpace($n4)) { $n4 = "Events" }
            Create-EventsList -ListName $n4

            $n5 = Read-Host "Meetings list name (default: Meetings)"
            if ([string]::IsNullOrWhiteSpace($n5)) { $n5 = "Meetings" }
            Create-UpcomingMeetingList -ListName $n5

            Create-RecentDocumentsLibrary -ListName ""

            $n7 = Read-Host "Modern Calendar list name (default: CalendarEvents)"
            if ([string]::IsNullOrWhiteSpace($n7)) { $n7 = "CalendarEvents" }
            Create-CalendarList -ListName $n7

            $n8 = Read-Host "Employee Directory list name (default: EmployeeDirectory)"
            if ([string]::IsNullOrWhiteSpace($n8)) { $n8 = "EmployeeDirectory" }
            Create-EmployeeDirectoryList -ListName $n8

            $n9a = Read-Host "Polls list name (default: Polls)"
            if ([string]::IsNullOrWhiteSpace($n9a)) { $n9a = "Polls" }
            $n9b = Read-Host "Poll Votes list name (default: PollVotes)"
            if ([string]::IsNullOrWhiteSpace($n9b)) { $n9b = "PollVotes" }
            Create-PollLists -PollsListName $n9a -VotesListName $n9b

            $n10a = Read-Host "Shoutouts list name (default: Shoutouts)"
            if ([string]::IsNullOrWhiteSpace($n10a)) { $n10a = "Shoutouts" }
            $n10b = Read-Host "Shoutout Likes list name (default: ShoutoutLikes)"
            if ([string]::IsNullOrWhiteSpace($n10b)) { $n10b = "ShoutoutLikes" }
            Create-ShoutoutLists -ShoutoutsListName $n10a -LikesListName $n10b

            $n11 = Read-Host "FAQ list name (default: FAQ)"
            if ([string]::IsNullOrWhiteSpace($n11)) { $n11 = "FAQ" }
            Create-FaqList -ListName $n11

            Write-Host ""
            Write-Host "=============================================" -ForegroundColor Green
            Write-Host "  All lists created successfully!" -ForegroundColor Green
            Write-Host "=============================================" -ForegroundColor Green
        }
        0 {
            Write-Host ""
            Write-Host "Disconnecting..." -ForegroundColor Yellow
            Disconnect-PnPOnline
            Write-Host "Done. Goodbye!" -ForegroundColor Green
            return $false
        }
        default {
            Write-Host "Invalid selection. Please enter a number between 0 and 11." -ForegroundColor Red
        }
    }
    return $true
}

# ============================================================
# MAIN LOOP
# ============================================================
$continue = $true
while ($continue) {
    Show-Menu
    $input = Read-Host "Enter your choice (0-12)"

    if ($input -match '^\d+$') {
        $choice = [int]$input
        try {
            $continue = Run-Selection -Choice $choice
        }
        catch {
            Write-Host ""
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Please check the list name and try again." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "Please enter a valid number." -ForegroundColor Red
    }
}
