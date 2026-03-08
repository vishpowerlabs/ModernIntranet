<#
.SYNOPSIS
Creates a sample SharePoint list for the Employee Directory web part and populates it with 20 mock records.

.DESCRIPTION
This script uses the PnP PowerShell module to connect to a SharePoint site, create a new Custom List named "Employee Directory", add necessary columns, and iteratively add sample employee data.

.EXAMPLE
.\Create-SampleEmployeeData.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Intranet"
#>

param (
    [string]$SiteUrl = "https://devtenant0424.sharepoint.com/sites/DEVSITE"
)

# Define the name of the list
$ListName = "Employee Directory5"

# Connect to the SharePoint Site
Write-Host "Connecting to $SiteUrl..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

# Check if list already exists
$list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue

if ($list) {
    Write-Host "List '$ListName' already exists. Deleting it to start fresh..." -ForegroundColor Yellow
    Remove-PnPList -Identity $ListName -Force
}

Write-Host "Creating list '$ListName'..." -ForegroundColor Green
$list = New-PnPList -Title $ListName -Template GenericList

Write-Host "Adding columns to the list..." -ForegroundColor Cyan
# Title defaults to Name
Set-PnPField -List $ListName -Identity "Title" -Values @{Title="Employee Name"}

# Add Custom Fields
Add-PnPField -List $ListName -Type Text -InternalName "JobTitle" -DisplayName "Job Title" | Out-Null
Add-PnPField -List $ListName -Type Text -InternalName "Department" -DisplayName "Department" | Out-Null
Add-PnPField -List $ListName -Type Text -InternalName "Location" -DisplayName "Location" | Out-Null
Add-PnPField -List $ListName -Type User -InternalName "Email" -DisplayName "Email" | Out-Null
Add-PnPField -List $ListName -Type Text -InternalName "Phone" -DisplayName "Phone Number" | Out-Null
Add-PnPField -List $ListName -Type User -InternalName "Manager" -DisplayName "Manager" | Out-Null
Add-PnPField -List $ListName -Type Thumbnail -InternalName "PhotoUrl" -DisplayName "Profile Photo" | Out-Null
Add-PnPField -List $ListName -Type Note -InternalName "Projects" -DisplayName "Projects" | Out-Null
Add-PnPField -List $ListName -Type Note -InternalName "AboutMe" -DisplayName "About Me" | Out-Null
Add-PnPField -List $ListName -Type Note -InternalName "Interests" -DisplayName "Interests" | Out-Null
Add-PnPField -List $ListName -Type Note -InternalName "Skills" -DisplayName "Skills" | Out-Null

# Arrays for random generation
$firstNames = @("Sanjay","Maria","Aiden","Sarah","David","Elena","Michael","Priya","Robert","Chloe","James","Yuki","Amanda","Carlos","John","Emily","Hassan","Oliver","Isabella","Liam","Emma","Noah","Olivia","William","Ava","James","Sophia","Oliver","Isabella","Benjamin","Mia","Elijah","Charlotte","Lucas","Amelia","Mason","Harper","Logan","Evelyn","Alexander")
$lastNames = @("Patel","Gonzalez","Chen","Jenkins","Smith","Rostova","Chang","Sharma","Wilson","Dubois","Taylor","Tanaka","Clark","Silva","Doe","White","Ali","Brown","Rossi","Davis","Johnson","Williams","Jones","Garcia","Miller","Davis","Rodriguez","Martinez","Hernandez","Lopez","Gonzalez","Wilson","Anderson","Thomas","Taylor","Moore","Jackson","Martin","Lee","Perez")
$jobTitles = @("Software Engineer", "Product Manager", "UX Designer", "Data Scientist", "Marketing Executive", "Financial Analyst", "Sales Representative", "QA Tester", "Systems Administrator", "Customer Support", "Operations Manager", "Business Analyst", "Technical Writer", "HR Specialist", "Accountant")
$departments = @("Engineering", "Product", "Design", "Data", "Marketing", "Finance", "Sales", "QA", "IT", "Operations", "Executive", "Human Resources")
$locations = @("London", "Madrid", "San Francisco", "New York", "Berlin", "Mumbai", "Paris", "Tokyo", "São Paulo", "Chicago", "Toronto", "Dubai", "Rome", "Sydney", "Singapore")
$realEmails = @("info@vishpowerlabs.com", "anu@vishpowerlabs.com", "AlexW@devtenant0424.onmicrosoft.com", "pradeep@vishpowerlabs.com", "vishnu@vishpowerlabs.com")
$sampleProjects = @("Project Alpha`nProject Beta", "Q3 Marketing Campaign", "Intranet Redesign`nSecurity Audit", "Data Migration 2025", "Cloud Native Initiative`nAPI Gateway")
$sampleAboutMes = @("I love driving innovation.", "Passionate about user experience and accessible design.", "Data-driven problem solver with 10 years experience.", "Always learning. Currently exploring AI.", "Avid runner and tech enthusiast.")
$sampleInterests = @("Reading`nCycling", "Photography`nTravel", "Baking`nBoard Games", "Machine Learning`nOpen Source", "Music`nGaming")
$sampleSkills = @("React`nTypeScript", "Project Management`nAgile", "Figma`nCSS", "Python`nSQL", "Public Speaking`nLeadership")

Write-Host "Pre-resolving users in Site Collection..." -ForegroundColor Cyan
$web = Get-PnPWeb
$context = Get-PnPContext
$resolvedUsers = @{}

foreach ($email in $realEmails) {
    try {
        $user = $web.EnsureUser($email)
        $context.Load($user)
        $context.ExecuteQuery()
        $resolvedUsers[$email] = $user.LoginName
        Write-Host "  Successfully ensured user: $email (Login: $($user.LoginName))" -ForegroundColor Green
    } catch {
        Write-Host "  [!] Failed to ensure user: $email. Error: $_" -ForegroundColor Red
    }
}

$employees = @()

# Temp image directory (Windows environment as requested)
$tempImagesPath = "C:\Temp\temp-avatars"

Write-Host "Generating 120 sample employees..." -ForegroundColor Cyan

for ($i = 1; $i -le 120; $i++) {
    $firstName = ($firstNames | Get-Random)
    $lastName = ($lastNames | Get-Random)
    $title = "$firstName $lastName"
    $phone = "+1 (555) $(Get-Random -Minimum 100 -Maximum 999)-$(Get-Random -Minimum 1000 -Maximum 9999)"
    $actualEmailStr = ($realEmails | Get-Random)
    $actualManagerStr = ($realEmails | Get-Random)
    
    # Assign a random local image from 1 to 10
    $imageIdx = $(Get-Random -Minimum 1 -Maximum 11)
    $localImagePath = Join-Path -Path $tempImagesPath -ChildPath "avatar$imageIdx.png"
    
    $employees += @{
        Title = $title
        JobTitle = ($jobTitles | Get-Random)
        Department = ($departments | Get-Random)
        Location = ($locations | Get-Random)
        Email = $resolvedUsers[$actualEmailStr]
        Manager = $resolvedUsers[$actualManagerStr]
        Phone = $phone
        LocalImagePath = $localImagePath
        Projects = ($sampleProjects | Get-Random)
        AboutMe = ($sampleAboutMes | Get-Random)
        Interests = ($sampleInterests | Get-Random)
        Skills = ($sampleSkills | Get-Random)
    }
}

Write-Host "Adding $($employees.Count) sample employees to the list..." -ForegroundColor Green

foreach ($employee in $employees) {
    Write-Host "Adding $($employee.Title)..." -ForegroundColor Gray

    if (-not $employee.Email -or -not $employee.Manager) {
        Write-Host "  [!] Skipping because user LoginName could not be found." -ForegroundColor Red
        continue
    }

    $itemValues = @{
        "Title" = $employee.Title
        "JobTitle" = $employee.JobTitle
        "Department" = $employee.Department
        "Location" = $employee.Location
        "Email" = $employee.Email
        "Phone" = $employee.Phone
        "Manager" = $employee.Manager
        "Projects" = $employee.Projects
        "AboutMe" = $employee.AboutMe
        "Interests" = $employee.Interests
        "Skills" = $employee.Skills
    }
    
    # First, create the item without the photo
    try {
        $newItem = Add-PnPListItem -List $ListName -Values $itemValues -ErrorAction Stop
        
        # Check if the photo file exists to upload to SiteAssets and attach
        if (Test-Path $employee.LocalImagePath) {
            Set-PnPImageListItemColumn -List $ListName -Identity $newItem.Id -Field "PhotoUrl" -Path $employee.LocalImagePath -ErrorAction Stop | Out-Null
        }
    } catch {
        Write-Host "  [!] Failed to add or upload image for $($employee.Title): $_" -ForegroundColor Yellow
    }
}

Write-Host "Sample data successfully created!" -ForegroundColor Green

Write-Host "Getting default view for '$ListName'..." -ForegroundColor Cyan
$view = Get-PnPView -List $ListName | Where-Object { $_.DefaultView -eq $true }

if ($view) {
    Write-Host "Updating view: $($view.Title)" -ForegroundColor Green

    try {
        Set-PnPView -List $ListName -Identity $view.Id -Fields "LinkTitle", "PhotoUrl", "JobTitle", "Department", "Location", "Email", "Phone", "Manager", "Projects", "AboutMe", "Interests", "Skills" -ErrorAction Stop
        Write-Host "Default view successfully updated with all columns!" -ForegroundColor Green
    } catch {
        Write-Host "  [!] Failed to update default view: $_" -ForegroundColor Red
    }
} else {
    Write-Host "Could not find the default view for the list." -ForegroundColor Red
}

Write-Host "You can now configure the Employee Directory web part to map to the '$ListName' list." -ForegroundColor Cyan

