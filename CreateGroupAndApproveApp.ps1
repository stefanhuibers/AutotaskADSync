# Variables
$groupName = "Baseline - App - Autotask Contacts Sync"
$appId = "939b40ad-5c52-4ad3-befe-467d3c83ce76"
$groupMembershipRule = "(user.assignedPlans -any (((assignedPlan.servicePlanId -eq `"9aaf7827-d63c-4b61-89c3-182f06f82e5c`") -or (assignedPlan.servicePlanId -eq `"efb87545-963c-4e0d-99df-69c6916d9eb0`") -or (assignedPlan.servicePlanId -eq `"4a82b400-a79f-41a4-b4e2-e94f5787b113`")) -and assignedPlan.capabilityStatus -eq `"Enabled`") )-and (user.accountEnabled -eq true) -and (user.userType -eq `"Member`") -and (user.givenName -ne null) -and (user.surname -ne null)"
$mailTo = "g.varekamp@xantion.nl"
$graphModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Groups", "Microsoft.Graph.Identity.DirectoryManagement")
$scopes = @("Group.ReadWrite.All", "Organization.Read.All")
$group = $null

function Install-GraphdModule {
    param(
        [string]$moduleName
    )
    $currentVersion = (Get-Module -ListAvailable -Name $moduleName).Version
    if (-not $currentVersion) {
        Write-Host "Module $moduleName not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $moduleName -Force -Scope CurrentUser
        Write-Host "Module $moduleName installed successfully." -ForegroundColor Green
    }
    else {
        $latestVersion = (Find-Module -Name $moduleName).Version
        if ($latestVersion -gt $currentVersion) {
            Write-Host "A newer version of the $moduleName module is available. Updating the module..." -ForegroundColor Yellow
            Install-Module -Name $moduleName -Force
            Write-Host "Module $moduleName updated successfully." -ForegroundColor Green
        }
    }
}

# Clear the console
Clear-Host

# Install the required modules
foreach ($module in $graphModules) {
    Install-GraphdModule -moduleName $module
}

# Connect to Microsoft Graph
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Connect-MgGraph -Scopes $scopes -NoWelcome
try {
    $tenantId = (Get-MgContext -ErrorAction Stop).TenantId
    $tenantDisplayName = (Get-MgOrganization -ErrorAction Stop).DisplayName
}
catch {
    Write-Host "Error connecting to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Check if the group already exists
$existingGroup = Get-MgGroup -Filter "displayName eq '$groupName'" -ErrorAction SilentlyContinue
if ($existingGroup) {
    Write-Host "Group '$groupName' already exists." -ForegroundColor Yellow
    $confirmation = Read-Host "Do you want want to remove the group '$groupName'? (y/n)"
    if ($confirmation -eq 'y') {
        try {
            Remove-MgGroup -GroupId $existingGroup.Id -Confirm:$false
            Write-Host "Group '$groupName' removed successfully." -ForegroundColor Green
        }
        catch {
            Write-Host "Error removing group '$groupName': $($_.Exception.Message)" -ForegroundColor Red
            exit 1
        }
    }
    else {
        $group = $existingGroup
        Write-Host "Group '$groupName' was not removed." -ForegroundColor Yellow
    }
}

# Create the group
if (-not $group) {
    $groupMailNickname = $groupName -replace "[^a-zA-Z0-9]", ""
    $params = @{
        displayName                   = $groupName
        mailEnabled                   = $false
        mailNickname                  = $groupMailNickname
        securityEnabled               = $true
        groupTypes                    = @("DynamicMembership")
        membershipRule                = $groupMembershipRule
        membershipRuleProcessingState = "On"
    }
    try {
        $group = New-MgGroup -BodyParameter $params
        Write-Host "Group '$groupName' created successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Error creating group '$groupName': $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Open the app consent URL
$appUrl = "https://login.microsoftonline.com/$tenantId/v2.0/adminconsent?client_id=$appId&scope=https://graph.microsoft.com/.default"
Start-Process $customAppURL
Write-Host "Please provide the required permissions to the application and press Enter to continue..." -ForegroundColor Yellow
Read-Host

# Create and open a new email using default mail client
$subject = "Active Directory (AD) Synchronization in Autotask voor klant $tenantDisplayName"
$body = @"
Beste applicatiebeheerder,

Voor de klant $tenantDisplayName kan Active Directory (AD) Synchronization in Autotask worden ingesteld.

Client ID: $appId
Tenant: $tenantId
Client Secret: <a href='https://xantion.eu.itglue.com/757545073688811/passwords/4017876740374726'>Klik hier om de client secret op te halen</a>
Group ID: $($group.Id)
"@
$mailtoUrl = "mailto:$mailTo`?subject=$([uri]::EscapeDataString($subject))&body=$([uri]::EscapeDataString($body))"
Write-Host "Opening email client..." -ForegroundColor Yellow
Start-Process $mailtoUrl

# Disconnect from Microsoft Graph
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Write-Host "Script completed successfully." -ForegroundColor Green
Read-Host "Press Enter to exit..."
exit 0
