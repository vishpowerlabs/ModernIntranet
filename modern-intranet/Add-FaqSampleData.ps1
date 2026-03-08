# ============================================================
# FAQ Sample Data Provisioning Script
# ============================================================
$ErrorActionPreference = "Stop"

$siteUrl = "https://devtenant0424.sharepoint.com/sites/DEVSITE" 
$listName = "FAQ"

Write-Host "Connecting to PnP Online..." -ForegroundColor Yellow
Connect-PnPOnline -Url $siteUrl -UseWebLogin

Write-Host "Checking if list '$listName' exists..." -ForegroundColor Yellow
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if (-not $list) {
    Write-Host "List '$listName' not found. Please run Provision-Lists.ps1 first." -ForegroundColor Red
    exit 1
}

$sampleData = @(
    @{
        Question = "How do I request a new laptop?";
        Answer = "You can request a new laptop through the <b>IT Helpdesk Portal</b>. Standard refresh cycle is every 3 years. For urgent requests, please obtain manager approval first.";
        Category = "IT Support";
        Order = 1
    },
    @{
        Question = "What is the policy for remote work?";
        Answer = "Contoso follows a <i>hybrid work model</i>. Employees are expected to be in the office 2-3 days per week. Please refer to the HR Portal for the full <b>Remote Work Policy</b>.";
        Category = "Human Resources";
        Order = 2
    },
    @{
        Question = "Where can I find my payslip?";
        Answer = "Payslips are available on the <b>Payroll Self-Service</b> portal. You can log in using your standard company credentials. New payslips are usually posted 2 days before payday.";
        Category = "Finance";
        Order = 3
    },
    @{
        Question = "How do I book a meeting room?";
        Answer = "Meeting rooms can be booked via <b>Outlook</b> or the <b>Meeting Room Finder</b> web part on the intranet home page. Simply select your preferred time and room size.";
        Category = "Operations";
        Order = 4
    },
    @{
        Question = "How do I update my internal directory profile?";
        Answer = "Navigate to the <i>Employee Directory</i> and click on <b>'Update My Profile'</b>. Changes to title or department must be initiated through HR.";
        Category = "IT Support";
        Order = 5
    },
    @{
        Question = "What are the company holidays for 2026?";
        Answer = "A full list of holidays is posted on the <b>HR Benefits</b> page. We observe all major public holidays plus two floating wellness days per year.";
        Category = "Human Resources";
        Order = 6
    },
    @{
        Question = "Who do I contact for travel bookings?";
        Answer = "All business travel must be booked through our partner agency, <b>Global Travel Pro</b>. Use the link in the <i>Quick Links</i> section to access the booking tool.";
        Category = "Finance";
        Order = 7
    },
    @{
        Question = "How do I reset my password?";
        Answer = "Visit <b>passwordreset.microsoftonline.com</b> to reset your password using Multi-Factor Authentication (MFA). If you are locked out, contact the IT helpdesk at ext. 555.";
        Category = "IT Support";
        Order = 8
    }
)

Write-Host "Adding sample data to '$listName'..." -ForegroundColor Cyan

foreach ($item in $sampleData) {
    Write-Host " - Adding: $($item.Question)" -ForegroundColor Gray
    Add-PnPListItem -List $listName -Values @{
        "Title"       = $item.Question;
        "FaqQuestion" = $item.Question;
        "FaqAnswer"   = $item.Answer;
        "FaqCategory" = $item.Category;
        "SortOrder"   = $item.Order
    } | Out-Null
}

Write-Host ""
Write-Host "Success! Added $($sampleData.Count) sample FAQ items." -ForegroundColor Green
