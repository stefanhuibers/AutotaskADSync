# Variables
$GroupName             = "Baseline - App - Autotask Contacts Sync"
$GroupMembershipRule   = "(user.assignedPlans -any (((assignedPlan.servicePlanId -eq `"9aaf7827-d63c-4b61-89c3-182f06f82e5c`") -or (assignedPlan.servicePlanId -eq `"efb87545-963c-4e0d-99df-69c6916d9eb0`") -or (assignedPlan.servicePlanId -eq `"4a82b400-a79f-41a4-b4e2-e94f5787b113`")) -and assignedPlan.capabilityStatus -eq `"Enabled`") )-and (user.accountEnabled -eq true) -and (user.userType -eq `"Member`") -and (user.givenName -ne null) -and (user.surname -ne null)"
$AppId                 = "939b40ad-5c52-4ad3-befe-467d3c83ce76"
$MailTo                = "g.varekamp@xantion.nl"
$MicrosoftGraphModules = @("Microsoft.Graph.Groups")
$MicrosoftGraphScopes  = @("Group.ReadWrite.All", "Organization.Read.All")
$script:WriteLog       = $false

function Write-Log {
    param (
        [string]$Message
    )
    if ($script:WriteLog) {
        Write-Host $Message
    }
}

# Clear the console
Clear-Host

# Check installed modules
foreach ($Name in $MicrosoftGraphModules) {
    try {
        $Installed = Get-Module -Name $Name -ListAvailable -ErrorAction Stop
        if ($Installed) {
            Write-Log "Module $Name $($Installed.Version) is already installed." -ForegroundColor Green
        } else {
            # Check if NuGet is installed as a package provider
            if (!(Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                Write-Log "NuGet package provider not found. Installing..." -ForegroundColor Yellow
                try {
                    Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
                    Write-Log "NuGet package provider installed successfully." -ForegroundColor Green
                } catch {
                    Write-Host "Error installing NuGet package provider: $($_.Exception.Message)" -ForegroundColor Red
                    exit 1
                }
            }
            # Install module from PSGallery
            Write-Log "Module $Name not found. Installing..." -ForegroundColor Yellow
            try {
                Install-Module -Name $Name -Scope CurrentUser -Repository PSGallery -Force
                $Installed = Get-Module -Name $Name -ListAvailable -ErrorAction Stop
                Write-Log "Module $Name $($Installed.Version) installed successfully." -ForegroundColor Green
            } catch {
                Write-Host "Error installing module $Name`: $($_.Exception.Message)" -ForegroundColor Red
                exit 1
            }
        }
    } catch {
        Write-Host "Error checking module $Name`: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# Import the modules
foreach ($Name in $MicrosoftGraphModules) {
    try {
        Import-Module -Name $Name -ErrorAction Stop
        Write-Log "Module $Name imported successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error importing module $Name`: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# Connect to Microsoft Graph
if (Get-MgContext -ErrorAction SilentlyContinue) {
    Disconnect-MgGraph | Out-Null
}
Connect-MgGraph -Scopes $MicrosoftGraphScopes -NoWelcome

# Get tenant information
try {
    $TenantId = Get-MgContext -ErrorAction Stop | Select-Object -ExpandProperty TenantId
}
catch {
    Write-Host "Error connecting to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Check if the group already exists
$Group = $null
$Group = Get-MgGroup -Filter "displayName eq '$groupName'" -ErrorAction SilentlyContinue
if ($Group) {
    Write-Host "Group '$groupName' already exists." -ForegroundColor Yellow
    $confirmation = Read-Host "Do you want want to remove the current group '$groupName'? (y/n)"
    if ($confirmation -eq 'y') {
        try {
            Remove-MgGroup -GroupId $Group.Id -Confirm:$false -ErrorAction Stop
            Write-Host "Group '$groupName' removed successfully." -ForegroundColor Green
            $Group = $null
        }
        catch {
            Write-Host "Error removing group '$groupName': $($_.Exception.Message)" -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "Group '$groupName' was not removed." -ForegroundColor Yellow
    }
}

# Create the group
if (-not $Group) {
    $GroupMailNickname = $groupName -replace "[^a-zA-Z0-9]", ""
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
        $group = New-MgGroup -BodyParameter $params -ErrorAction Stop
        Write-Host "Group '$groupName' created successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Error creating group '$groupName': $($_.Exception.Message)" -ForegroundColor Red
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Read-Host "Press Enter to exit..."
        exit 1
    }
}

# Open the app consent URL
$appUrl = "https://login.microsoftonline.com/$tenantId/v2.0/adminconsent?client_id=$appId&scope=https://graph.microsoft.com/.default"
Start-Process $appUrl
Write-Host "Please provide the required permissions to the application and press Enter to continue..." -ForegroundColor Yellow
Read-Host

# Get company name from user input
$CompanyName = Read-Host "Company name"
if ($CompanyName -eq "") {
    $CompanyName = "<Company Name>"
}

# Create and open a new email using default mail client
$subject = "Active Directory (AD) Synchronization in Autotask voor klant `"$CompanyName`""
$body = @"
Beste applicatiebeheerder,

Voor de klant "$CompanyName" kan Active Directory (AD) Synchronization in Autotask worden ingesteld.

Client ID:
$AppId

Tenant ID:
$TenantId

Client Secret:
https://xantion.eu.itglue.com/4010208547029165/passwords/4013394209947788

Group ID:
$($Group.Id)

"@
$mailtoUrl = "mailto:$mailTo`?subject=$([uri]::EscapeDataString($subject))&body=$([uri]::EscapeDataString($body))"
Write-Host "Opening email client..." -ForegroundColor Yellow
Start-Process $mailtoUrl

# Disconnect from Microsoft Graph
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
Write-Host "Script completed successfully." -ForegroundColor Green
Read-Host "Press Enter to exit..."
exit 0