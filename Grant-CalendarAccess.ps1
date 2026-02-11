#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Interactive script to grant calendar permissions in Exchange Online.
.DESCRIPTION
    Authenticates, prompts for Target/Reviewer/Access Level, applies permissions, and reports status.
#>

# --- 1. Authentication ---
Write-Host "Checking Exchange Online connection..." -ForegroundColor Cyan
try {
    # Quick check if we can reach a cmdlet
    Get-Mailbox -Identity "discovery" -ErrorAction Stop | Out-Null
    Write-Host "Already connected." -ForegroundColor Green
}
catch {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ShowProgress $true -ErrorAction Stop
        Write-Host "Connected successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Exchange Online. Please check your internet or credentials."
        exit
    }
}

# --- 2. Gather Inputs ---
Write-Host "`n--- Calendar Permission Wizard ---" -ForegroundColor Cyan

# Get Target (The Director/Owner)
do {
    $TargetUser = Read-Host "Enter the EMAIL of the Calendar OWNER (e.g. Director)"
    if ([string]::IsNullOrWhiteSpace($TargetUser)) { Write-Warning "Email is required." }
} until (-not [string]::IsNullOrWhiteSpace($TargetUser))

# Get Reviewer (The EA/User)
do {
    $ReviewerUser = Read-Host "Enter the EMAIL of the person getting ACCESS (e.g. EA)"
    if ([string]::IsNullOrWhiteSpace($ReviewerUser)) { Write-Warning "Email is required." }
} until (-not [string]::IsNullOrWhiteSpace($ReviewerUser))

# Validate Users Exist
Write-Host "Validating users..." -NoNewline
try {
    $TargetMailbox = Get-Mailbox -Identity $TargetUser -ErrorAction Stop
    $ReviewerMailbox = Get-Mailbox -Identity $ReviewerUser -ErrorAction Stop
    Write-Host " OK." -ForegroundColor Green
}
catch {
    Write-Host "`n[ERROR] One of the users could not be found. Please check spelling." -ForegroundColor Red
    exit
}

# Get Permissions Level
Write-Host "`nChoose Access Level:" -ForegroundColor Yellow
Write-Host "1. Editor           (Create, Read, Edit, Delete Own & Others' items - Standard for EAs)"
Write-Host "2. Reviewer         (Read-Only: Full Details)"
Write-Host "3. AvailabilityOnly (Free/Busy time only)"
Write-Host "4. LimitedDetails   (Subject & Location only)"

$Selection = Read-Host "Select (1-4)"
switch ($Selection) {
    "1" { $AccessRights = "Editor" }
    "2" { $AccessRights = "Reviewer" }
    "3" { $AccessRights = "AvailabilityOnly" }
    "4" { $AccessRights = "LimitedDetails" }
    Default { 
        Write-Warning "Invalid selection. Defaulting to Reviewer."
        $AccessRights = "Reviewer" 
    }
}

# --- 3. Apply Permissions ---
$CalendarFolder = "$($TargetUser):\Calendar"

Write-Host "`nApplying '$AccessRights' access for $ReviewerUser on $($TargetUser)'s Calendar..." -ForegroundColor Cyan

try {
    # Attempt to ADD permission
    Add-MailboxFolderPermission -Identity $CalendarFolder -User $ReviewerUser -AccessRights $AccessRights -ErrorAction Stop
    Write-Host "[SUCCESS] Permission Added." -ForegroundColor Green
}
catch {
    # If it fails, check if permission exists and needs UPDATE
    if ($_.Exception.Message -match "An existing permission entry was found") {
        Write-Host "User already has permission. Updating to '$AccessRights'..." -ForegroundColor Yellow
        try {
            Set-MailboxFolderPermission -Identity $CalendarFolder -User $ReviewerUser -AccessRights $AccessRights -ErrorAction Stop
            Write-Host "[SUCCESS] Permission Updated." -ForegroundColor Green
        }
        catch {
            Write-Host "[ERROR] Update failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "[ERROR] Could not apply permission: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# --- 4. Final Review ---
Write-Host "`n--- Current Calendar Permissions for $TargetUser ---" -ForegroundColor Cyan
try {
    Get-MailboxFolderPermission -Identity $CalendarFolder | 
        Select-Object User, AccessRights, SharingPermissionFlags |
        Format-Table -AutoSize
}
catch {
    Write-Error "Could not retrieve final permissions report."
}

Write-Host "Note: Changes may take up to 60 minutes to appear in Outlook Desktop (Instant in OWA)." -ForegroundColor Gray
Pause
