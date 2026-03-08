<#
.SYNOPSIS
Updates the default view of the Employee Directory list to include all custom columns.
#>

$ListName = "Employee Directory"

Write-Host "Getting default view for '$ListName'..." -ForegroundColor Cyan
$view = Get-PnPView -List $ListName | Where-Object { $_.DefaultView -eq $true }

if ($view) {
    Write-Host "Updating view: $($view.Title)" -ForegroundColor Green

    $fieldsToAdd = @("PhotoUrl", "Title", "JobTitle", "Department", "Location", "Email", "Phone", "Manager", "Projects", "AboutMe", "Interests", "Skills")

    foreach ($field in $fieldsToAdd) {
        Write-Host "  Adding column '$field' to view..." -ForegroundColor Gray
        Add-PnPViewField -List $ListName -Identity $view.Id -Field $field -ErrorAction SilentlyContinue
    }
    
    Write-Host "Default view successfully updated!" -ForegroundColor Green
} else {
    Write-Host "Could not find the default view for the list." -ForegroundColor Red
}
