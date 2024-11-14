# -------------------------------
# Office 365 User Deprovisioning Script
# -------------------------------

# Import necessary modules
Import-Module AzureAD
Import-Module ExchangeOnlineManagement

# Function to log messages
function Write-Log {
    param (
        [string]$Message,
        [string]$Path = "C:\Logs\O365_Deprovisioning.log"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    Write-Output $logMessage
    Add-Content -Path $Path -Value $logMessage
}

# Ensure log directory exists
$logDirectory = "C:\Logs"
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

$logFile = "$logDirectory\O365_Deprovisioning.log"

# Connect to Azure AD
try {
    Write-Log "Connecting to Azure AD..."
    Connect-AzureAD -ErrorAction Stop
    Write-Log "Successfully connected to Azure AD."
} catch {
    Write-Log "ERROR: Failed to connect to Azure AD. $_"
    exit
}

# Connect to Exchange Online
try {
    Write-Log "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ErrorAction Stop
    Write-Log "Successfully connected to Exchange Online."
} catch {
    Write-Log "ERROR: Failed to connect to Exchange Online. $_"
    exit
}

# Define the user prefix
$userPrefix = "xEM - "

# Retrieve all users with the specified prefix in their DisplayName
try {
    Write-Log "Retrieving users with prefix '$userPrefix'..."
    $users = Get-AzureADUser -All $true | Where-Object { $_.DisplayName -like "$userPrefix*" }
    Write-Log "Found $($users.Count) user(s) with the prefix '$userPrefix'."
} catch {
    Write-Log "ERROR: Failed to retrieve users. $_"
    exit
}

if ($users.Count -eq 0) {
    Write-Log "No users found with the prefix '$userPrefix'. Exiting script."
    exit
}

# Iterate through each user and perform the required actions
foreach ($user in $users) {
    Write-Log "Processing user: $($user.DisplayName) ($($user.UserPrincipalName))"

    # -------------------------------
    # Step 1: Remove User from All Groups
    # -------------------------------
    try {
        Write-Log "Retrieving group memberships for $($user.UserPrincipalName)..."
        $groups = Get-AzureADUserMembership -ObjectId $user.ObjectId | Where-Object { $_.ObjectType -eq "Group" }
        if ($groups.Count -eq 0) {
            Write-Log "User $($user.UserPrincipalName) is not a member of any groups."
        } else {
            foreach ($group in $groups) {
                try {
                    Write-Log "Removing $($user.UserPrincipalName) from group: $($group.DisplayName)..."
                    Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $user.ObjectId -ErrorAction Stop
                    Write-Log "Successfully removed from group: $($group.DisplayName)."
                } catch {
                    Write-Log "WARNING: Failed to remove from group: $($group.DisplayName). $_"
                }
            }
        }
    } catch {
        Write-Log "ERROR: Unable to retrieve or remove group memberships for $($user.UserPrincipalName). $_"
    }

    # -------------------------------
    # Step 2: Remove All Licenses
    # -------------------------------
    try {
        Write-Log "Retrieving assigned licenses for $($user.UserPrincipalName)..."
        $licensedUser = Get-AzureADUser -ObjectId $user.ObjectId | Select-Object -ExpandProperty AssignedLicenses
        if ($licensedUser.Count -eq 0) {
            Write-Log "User $($user.UserPrincipalName) has no assigned licenses."
        } else {
            foreach ($license in $licensedUser) {
                try {
                    Write-Log "Removing license: $($license.SkuId) from $($user.UserPrincipalName)..."
                    Set-AzureADUserLicense -ObjectId $user.ObjectId -AssignedLicenses @{ RemoveLicenses = @($license.SkuId) } -ErrorAction Stop
                    Write-Log "Successfully removed license: $($license.SkuId)."
                } catch {
                    Write-Log "WARNING: Failed to remove license: $($license.SkuId). $_"
                }
            }
        }
    } catch {
        Write-Log "ERROR: Unable to retrieve or remove licenses for $($user.UserPrincipalName). $_"
    }

    # -------------------------------
    # Step 3: Convert User Mailbox to Shared Mailbox
    # -------------------------------
    try {
        Write-Log "Converting mailbox of $($user.UserPrincipalName) to a shared mailbox..."
        Set-Mailbox -Identity $user.UserPrincipalName -Type Shared -ErrorAction Stop
        Write-Log "Successfully converted to shared mailbox."
    } catch {
        Write-Log "ERROR: Failed to convert mailbox to shared. $_"
    }

    Write-Log "Completed processing for user: $($user.UserPrincipalName).`n"
}

# Disconnect from services
try {
    Write-Log "Disconnecting from Azure AD..."
    Disconnect-AzureAD -Confirm:$false
    Write-Log "Disconnected from Azure AD."

    Write-Log "Disconnecting from Exchange Online..."
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Log "Disconnected from Exchange Online."
} catch {
    Write-Log "WARNING: Failed to disconnect from services. $_"
}

Write-Log "Script execution completed."
