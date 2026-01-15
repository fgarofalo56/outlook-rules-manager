# Using an Existing Azure AD App Registration

This guide explains how to configure the Outlook Rules Manager to work with an **existing** Azure AD app registration. Use this when:

- Your IT department restricts creating new app registrations
- You have an existing app that already has the required permissions
- You're reusing an app registration across multiple projects

## Quick Start

If you have an existing app registration, use this command to validate and configure it:

```powershell
.\src\Register-OutlookRulesApp.ps1 -UseExisting -ClientId "your-client-id" -TenantId "your-tenant-id"
```

This will:
1. Validate all required settings
2. Report any issues
3. Save the configuration to `.env` if validation passes

To automatically fix common issues:

```powershell
.\src\Register-OutlookRulesApp.ps1 -UseExisting -ClientId "your-client-id" -TenantId "your-tenant-id" -AutoFix
```

## Required App Settings

Your existing app registration **must** have these settings configured:

### 1. Public Client Flow (Required)

Device code authentication requires public client flow to be enabled.

**Where to configure:**
- Azure Portal > App registrations > [Your App] > Authentication
- Under "Advanced settings", set **Allow public client flows** to **Yes**

### 2. API Permissions (Required)

The app needs these delegated (not application) permissions:

| Permission | Type | Purpose |
|------------|------|---------|
| `Mail.ReadWrite` | Delegated | Create/manage mail folders, read/update rules |
| `User.Read` | Delegated | Read basic user profile |

**Where to configure:**
- Azure Portal > App registrations > [Your App] > API permissions
- Click "Add a permission" > Microsoft Graph > Delegated permissions
- Search for and add `Mail.ReadWrite` and `User.Read`

**Note:** Admin consent is NOT required for these permissions - users consent for their own mailbox only.

### 3. Redirect URI (Recommended)

For device code flow, add this redirect URI:

```
https://login.microsoftonline.com/common/oauth2/nativeclient
```

**Where to configure:**
- Azure Portal > App registrations > [Your App] > Authentication
- Under "Platform configurations", add a "Mobile and desktop applications" platform
- Add the redirect URI above

### 4. Supported Account Types (Check Configuration)

The app should support the account types you need:

| Setting | Description |
|---------|-------------|
| Single tenant | Only accounts in your organization |
| Multitenant | Accounts in any Azure AD organization |
| Multitenant + personal | Any Azure AD org + personal Microsoft accounts |

**Where to check:**
- Azure Portal > App registrations > [Your App] > Overview
- Look at "Supported account types"

For most use cases, "Accounts in any organizational directory" (multitenant) works best.

## Optional: App Roles for Multi-User Authorization

If multiple people will use this app and you want role-based access control, add these app roles:

### OutlookRules.Admin Role
- **Display name:** Administrator
- **Allowed member types:** Users
- **Value:** `OutlookRules.Admin`
- **Description:** Administrators can manage authorized users and perform all mailbox operations.
- **ID:** `f8b8c3d1-9a2b-4c5e-8f7d-6a1b2c3d4e5f`

### OutlookRules.User Role
- **Display name:** User
- **Allowed member types:** Users
- **Value:** `OutlookRules.User`
- **Description:** Users can perform mailbox operations on their own mailbox only.
- **ID:** `a1b2c3d4-5e6f-7a8b-9c0d-1e2f3a4b5c6d`

**Where to configure:**
- Azure Portal > App registrations > [Your App] > App roles
- Create role > Fill in details > Save

**Note:** App roles are optional. Without them, anyone who can authenticate can use the app.

## Manual Configuration

If you prefer to configure manually instead of using the validation script:

### Step 1: Find Your App IDs

1. Go to Azure Portal > Microsoft Entra ID > App registrations
2. Find your app and click on it
3. Copy the **Application (client) ID**
4. Copy the **Directory (tenant) ID**

### Step 2: Create .env File

Create a `.env` file in the project root:

```powershell
$ClientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$TenantId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

Replace the x's with your actual IDs.

### Step 3: Test Connection

```powershell
.\src\Connect-OutlookRulesApp.ps1
```

## Troubleshooting

### Error: "Application not found"

**Cause:** Client ID is wrong or app exists in different tenant.

**Fix:**
1. Verify the Client ID in Azure Portal
2. Ensure you're using the correct Tenant ID
3. Run: `.\src\Test-ExistingAppRegistration.ps1 -ClientId "..." -TenantId "..."`

### Error: "Public client flow not enabled"

**Cause:** Device code auth requires public client flow.

**Fix:**
1. Azure Portal > App registrations > [Your App] > Authentication
2. Set "Allow public client flows" to Yes
3. Or run with `-AutoFix` parameter

### Error: "Redirect URI not configured"

**Cause:** Missing redirect URI for authentication.

**Fix:**
1. Azure Portal > App registrations > [Your App] > Authentication
2. Add platform: Mobile and desktop applications
3. Add URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`

### Error: "Insufficient permissions"

**Cause:** Missing API permissions or consent not granted.

**Fix:**
1. Azure Portal > App registrations > [Your App] > API permissions
2. Ensure Mail.ReadWrite and User.Read are listed
3. If needed, click "Grant admin consent" (or consent on first login)

### Error: "User not assigned to application"

**Cause:** "User Assignment Required" is enabled but you're not assigned.

**Fix:**
1. Contact your administrator to assign you to the app
2. Or use: `.\src\Manage-AppAuthorization.ps1 -Operation AddUser -UserPrincipalName you@company.com`

## Validation Script Details

The `Test-ExistingAppRegistration.ps1` script checks:

| Check | Required | Description |
|-------|----------|-------------|
| Public Client | Yes | Device code flow requires this |
| Mail.ReadWrite | Yes | Core functionality |
| User.Read | Yes | User profile access |
| Redirect URI | Recommended | For interactive auth |
| App Roles | Optional | For multi-user authorization |
| Service Principal | Auto-created | Created on first sign-in |

### Validation Script Parameters

| Parameter | Description |
|-----------|-------------|
| `-ClientId` | (Required) Application client ID |
| `-TenantId` | (Required) Azure AD tenant ID |
| `-AutoFix` | Attempt to fix issues automatically |
| `-SaveConfig` | Save configuration to .env on success |
| `-ConfigProfile` | Profile name for multi-account setups |

### Examples

```powershell
# Just validate (don't save)
.\src\Test-ExistingAppRegistration.ps1 -ClientId "abc" -TenantId "xyz"

# Validate and save configuration
.\src\Test-ExistingAppRegistration.ps1 -ClientId "abc" -TenantId "xyz" -SaveConfig

# Validate, fix issues, and save
.\src\Test-ExistingAppRegistration.ps1 -ClientId "abc" -TenantId "xyz" -AutoFix -SaveConfig

# For multi-account setup
.\src\Test-ExistingAppRegistration.ps1 -ClientId "abc" -TenantId "xyz" -SaveConfig -ConfigProfile work
```

## Multi-Account Setup

If you manage multiple email accounts (personal + work), create separate configurations:

```powershell
# Personal account
.\src\Register-OutlookRulesApp.ps1 -UseExisting -ClientId "personal-client-id" -TenantId "personal-tenant-id" -ConfigProfile personal

# Work account
.\src\Register-OutlookRulesApp.ps1 -UseExisting -ClientId "work-client-id" -TenantId "work-tenant-id" -ConfigProfile work
```

This creates:
- `.env.personal` for personal account
- `.env.work` for work account

Connect to each:
```powershell
.\src\Connect-OutlookRulesApp.ps1 -ConfigProfile personal
.\src\Connect-OutlookRulesApp.ps1 -ConfigProfile work
```

## Security Considerations

When using an existing app registration:

1. **Verify permissions are minimal** - Only Mail.ReadWrite and User.Read are needed
2. **Check for unexpected permissions** - Extra permissions may be a security risk
3. **Review "User Assignment Required"** - Enable if you want to control who can use the app
4. **Audit app usage** - Check sign-in logs periodically

## Getting Help

If you encounter issues:

1. Run the validation script for diagnostics:
   ```powershell
   .\src\Test-ExistingAppRegistration.ps1 -ClientId "..." -TenantId "..."
   ```

2. Check connection with verbose output:
   ```powershell
   .\src\Connect-OutlookRulesApp.ps1
   ```

3. Review Azure AD sign-in logs for authentication errors

4. Contact your IT administrator if tenant policies are blocking access
