<#
.SYNOPSIS
    IT Operations Center - Comprehensive GUI-based management tool for IT infrastructure and services

.DESCRIPTION
    A comprehensive PowerShell GUI tool for managing Exchange Online mailbox permissions, calendar permissions,
    automatic replies (Out of Office), message tracking, Active Directory group memberships, and Microsoft 365
    authentication methods. Features include:
    - Optional Exchange Online connection (connect when needed)
    - Visual connection status indicator with Connect/Disconnect controls
    - Mailbox permissions (Full Access & Send As)
    - Calendar permissions (7 levels)
    - Automatic Replies (OOF) with rich text editor
    - Message Trace / Tracking for email delivery troubleshooting
    - AD group member viewing and management
    - Microsoft 365 Temporary Access Pass generation
    - Excel export capabilities for all data
    - Double-click any user/group to view their AD properties
    - GUID resolution for AD groups in permission lists

.AUTHOR
    Created by: Craig Werts
    Department: Desktop Engineering

.VERSION HISTORY
    Version 1.0.0 - 10-06-25
    - Initial release
    - Basic mailbox and calendar permission management
    - AD group member viewing
    
    Version 2.0 - 10-08-25
    - Added Excel export functionality for all three modules
    - Added GUID-to-name resolution for AD groups in permissions
    - Added double-click functionality to view AD properties (ADUC-style)
    - AD Properties window shows 4 tabs: General, Contact Info, Organization, Account
    - Copy email to clipboard feature in AD Properties window
    
    Version 2.6 - 10-09-25
    - Improved error handling and logging
    - Added progress indicators for group member loading
    - Added "Copy Emails" button for group members (Outlook format)
    - Enhanced UI with status indicators
    - AutoMapping disabled by default for Full Access permissions
    - Added "Add_KeyDown" event handler to Mailbox and Calendar windows
    
    Version 2.6.1 - 10-09-25
    - Added Company Logo and dynamic version text
    
    Version 2.7.0 - 10-13-25
    - Added Automatic Replies (Out of Office) management module
    - Implemented rich text editor with formatting toolbar (Bold, Italic, Underline)
    - HTML-enabled message editor - no HTML knowledge required
    - Support for internal and external automatic reply messages
    - Scheduled automatic replies with date/time picker
    - Three reply states: Disabled, Enabled, Scheduled
    - Automatic HTML conversion from formatted text
    - Visual status indicators with color coding
    - External audience controls (All/Contacts only)
    
    Version 2.8.0 - 10-13-25
    - Implemented optional Exchange Online connection
    - GUI now launches immediately without requiring EXO connection
    - Added visual connection status indicator (Red/Green)
    - Connect/Disconnect buttons in GUI
    - Connection check before opening each module
    - Auto-restore GUI after successful authentication
    - Prompts user to connect when accessing modules while disconnected
    - Improved flexibility for users who don't need immediate EXO access
    
    Version 2.8.1 - 10-13-25
    - Fixed bug in mailbox permissions loading when only one delegate exists
    - Fixed bug in calendar permissions loading when only one delegate exists
    - Improved array handling in Get-CombinedMailboxPermissions function
    - Improved array handling in calendar permission retrieval
    - Added null/empty value checking for permission entries
    - Enhanced error logging for permission retrieval
    - Functions now properly return IEnumerable collections for DataGrid binding

    Version 2.9.0 - 10-16-25
    - Modified AD Group Members to use Active Directory module instead of Exchange Online
    - Now supports all AD group types (Security Groups, Distribution Groups, etc.)
    - No longer requires Exchange Online connection for AD Group Members feature
    - Added Group Scope display (DomainLocal, Global, Universal)
    - Enhanced member type support (Users, Groups, Computers, Contacts)
    - Added SAM Account Name to Excel export
    - Requires Active Directory PowerShell module (RSAT)

    Version 3.2.0 - 10-23-25 
    - Added Message Trace / Tracking module for email delivery troubleshooting
    - Search by sender, recipient, subject, message ID, or status
    - Flexible date range options (24 hours, 7 days, or custom up to 10 days)
    - View detailed message trace events and delivery timeline
    - Support for up to 5,000 results per search
    - Export message trace results to Excel
    - Copy Message ID to clipboard for further investigation
    - Status filtering (Delivered, Failed, Pending, Quarantined, FilteredAsSpam)
    - Real-time search progress indicators
    - Full message delivery event tracking with timestamps
    
    Version 3.3.0 - 10-27-25
    - Added IP Network Scanner module for network infrastructure management
    - Scan IP address ranges to discover active devices
    - **Parallel scanning with 50 concurrent threads for fast performance**
    - Test connectivity using Test-Connection (ping) with 1-second timeout
    - Automatic hostname resolution for online devices
    - MAC address detection via ARP table lookup
    - Response time measurement in milliseconds
    - Real-time scan progress with online/offline counters
    - Export scan results to Excel with full device details
    - Support for large IP ranges with progress indicators
    - Color-coded status display (Green=Online, Red=Offline)
    - Typical scan speed: ~50 IPs per second (vs ~1 IP/sec sequential)

        Version 3.4.0 - 10-28-25
    - Added Export Active Users Report module
    - Retrieves all enabled AD user accounts from Active and Consultants OUs
    - Automatically filters out test accounts (containing "test" or "t-" prefix)
    - Exports Name, SamAccountName, Email, Department, Title, and Office (desk location)
    - Results automatically sorted by Name in ascending order
    - User-selectable save location (file save dialog)
    - Excel export with auto-formatting, filters, and frozen header row
    - Requires Active Directory PowerShell module (RSAT)

    Version 3.5.0 - 10-28-25
    - Added Intune Mobile Devices module for MDM device management
    - Retrieves all mobile devices (iOS, iPadOS, Android) from Microsoft Intune
    - Dashboard with real-time statistics (device counts by OS, compliance status)
    - Comprehensive device information (name, user, IMEI, serial, model, OS, storage)
    - Excel export with conditional formatting (compliant=green, non-compliant=red, jailbroken=orange)
    - IMEI and MEID formatted as numbers (no decimals) in Excel
    - Automatic Microsoft Graph API connection and consent management
    - Requires Microsoft.Graph.Authentication and Microsoft.Graph.DeviceManagement modules
    - Device enrollment dates, last sync times, and compliance states
    - Sortable grid with filtering capabilities
    - Added Export Active Users Report module
    - Retrieves all enabled AD user accounts from Active and Consultants OUs
    - Automatically filters out test accounts (containing "test" or "t-" prefix)
    - Exports Name, SamAccountName, Email, Department, Title, and Office (desk location)
    - Results automatically sorted by Name in ascending order
    - User-selectable save location (file save dialog)
    - Excel export with auto-formatting, filters, and frozen header row
    - Requires Active Directory PowerShell module (RSAT)
    
    Version 4.0.0 - 10-29-25
    - MAJOR UPDATE: Complete UI Redesign and Management Options Reorganization
    - Header redesign: Logo and title combined in single blue bar (no more white space)
    - Saved ~75px of vertical space with cleaner, more professional look
    - Management Options completely reorganized from 7 to 6 logical categories
    - New "Exchange Online" category consolidates all Exchange features (7 items)
    - Moved Calendar Permissions from "Calendar & Resources" to "Exchange Online"
    - Moved Message Trace/Tracking from "Compliance & Security" to "Exchange Online"
    - Renamed "Intune & SCCM" to "Device Management" for clarity
    - "Active Directory" category now contains pure AD functions (3 items)
    - Improved feature organization by service type for intuitive navigation
    - Code cleanup: reduced from 6,880 to 6,858 lines (22 lines removed)
    - All functionality preserved - 100% backward compatible
    - Zero breaking changes - drop-in replacement
    - Enhanced scalability for future feature additions
    - More professional and polished interface
    
    Version 4.1.0 - 10-29-25
    - Enhanced Intune Mobile Devices module with advanced filtering and search
    - Added device type checkboxes to filter by iOS/iPadOS, Android, Windows, and macOS
    - Filters are now applied automatically on Load/Refresh (removed separate Apply Filter button)
    - Updated statistics dashboard to show counts for all four device types (iOS, Android, Windows, macOS)
    - Retrieves ALL managed devices from Intune (not just mobile devices)
    - Added real-time search box to quickly find devices by name, user, model, serial number, IMEI, or OS
    - Search filters instantly as you type without making new API calls
    - Clear button to quickly reset search and show all devices
    - Search box automatically clears when loading/refreshing devices
    - Improved device filtering logic to support all Windows versions (Windows 10, Windows 11, etc.)
    - Enhanced error handling with better null-checking throughout event handlers
    - Simplified UI workflow for better user experience
    - Refresh button now properly reloads devices with current filter selections
    - Device count updates dynamically based on search and filter criteria
    - Removed unnecessary Write-Log calls from Intune window for better stability

    Version 4.2.0 - 11-10-25
    - Added new Microsoft 365 section to Management Options
    - Implemented Generate Temporary Access Pass (TAP) feature
    - TAP generation with configurable lifetime (10-43200 minutes)
    - Support for one-time use or multi-use TAPs
    - Automatic Microsoft Graph connection with UserAuthenticationMethod.ReadWrite.All scope
    - User-friendly TAP generation window with clear instructions
    - One-click copy to clipboard functionality
    - Real-time TAP details display (user, expiry time, usage type)
    - Security warning about TAP being shown only once
    - Support for username or UPN lookup
    - Enter key support for quick TAP generation
    - Professional UI with emoji icons and color-coded sections
    - Comprehensive error handling and logging


    Version 4.3.0 - 11-12-25
    - Added Employee Conversion module to Active Directory section
    - Convert employees between status types (e.g., Consultant to Full Time Employee)
    - Template-based conversion with surgical group management
    - Only removes groups associated with "from template" (preserves custom groups)
    - Adds all groups from "to template"
    - Automatically moves user to new template's OU
    - Real-time preview of changes before applying
    - Comprehensive logging of all conversion actions
    - Built-in validation and error handling
    
    Version 4.4.0 - 11-14-25
    - Added Locked Out Users Management module to Active Directory section
    - Real-time search and filtering of locked out user accounts
    - Multi-select unlock functionality with bulk operations
    - Displays lockout time, bad logon count, and user details
    - Excel export capabilities with formatted reports
    - Comprehensive error handling and success/failure tracking
    - Automatic refresh after unlock operations
    - Professional red-themed UI matching security context
    - Search across display name, username, email, department, and title
    - Status bar with live count updates
    - Requires Active Directory PowerShell module (RSAT)
    
    Version 4.5.0 - 11-14-25
    - Updated header information and version tracking
    - Enhanced documentation with complete feature descriptions
    - Improved version history organization and clarity
    - Code refinements and stability improvements
    - Total codebase: 8,443 lines of PowerShell
    - 4 major feature categories: Exchange Online, Microsoft 365, Active Directory, Device Management
    - 11 active modules with 3 placeholder features for future development

    Version 4.7.1 - 03-17-26 (Current)
    - Fixed Azure.Core.dll conflict between ExchangeOnlineManagement and Microsoft Graph modules
    - ExchangeOnlineManagement is now loaded on-demand when connecting to Exchange Online
    - Previously loaded at startup unconditionally, preventing Graph/Intune modules from loading
    - Intune Mobile Devices module now works in a fresh session without connecting to EXO first


.NOTES
    File Name      : IT-Operations-Center.ps1
    Prerequisites  : 
    - ExchangeOnlineManagement module (for Exchange Online features)
    - Microsoft.Graph.Authentication and Microsoft.Graph.Users modules (for M365 TAP generation)
    - Active Directory PowerShell module/RSAT (for AD features)
    
.USAGE
    Simply run the script. It will:
    1. Check for required modules and install ImportExcel if needed
    2. Connect to Exchange Online (browser authentication)
    3. Launch the GUI management tool
    
    From the GUI you can:
    - Manage Mailboxes: Add/edit/remove Full Access and Send As permissions
    - Calendar Permissions: Add/edit/remove calendar delegation permissions
    - Automatic Replies: Configure Out of Office messages with rich text formatting
    - Message Trace: Track and troubleshoot email delivery issues
    - Microsoft 365: Generate Temporary Access Passes for passwordless authentication onboarding
    - AD Group Members: View group members, export to Excel, copy email addresses
    - Double-click any user/group name in permission lists to view their AD properties

.FEATURES
    Connection Management:
    - Optional Exchange Online connection (launch GUI without connecting)
    - Visual status indicator (Red=Disconnected, Green=Connected)
    - Connect/Disconnect buttons in GUI
    - Console-based authentication with auto-restore GUI
    - Connection validation before accessing modules
    - Reconnect capability if session expires
    
    Mailbox Permissions:
    - Add Full Access and/or Send As permissions
    - View and edit existing permissions
    - Remove permissions selectively
    - Export permissions to Excel
    - GUID resolution for AD groups
    
    Calendar Permissions:
    - 7 permission levels (AvailabilityOnly to Owner)
    - Add, edit, and remove calendar access
    - Export permissions to Excel
    
    Automatic Replies (Out of Office):
    - Rich text editor with formatting toolbar (Bold, Italic, Underline)
    - Create formatted messages without HTML knowledge
    - Automatic HTML conversion and rendering
    - Internal and external message support
    - Three states: Disabled, Enabled (always on), Scheduled (date/time range)
    - Date/time picker for scheduled replies
    - External audience options (All senders or Contacts only)
    - Color-coded status display (Gray=Disabled, Green=Enabled, Orange=Scheduled)
    - Clear formatting button to remove all text formatting
    
    Message Trace / Tracking:
    - Search by sender, recipient, subject, message ID, or delivery status
    - Flexible date ranges (last 24 hours, last 7 days, custom up to 10 days)
    - Status filtering (All, Delivered, Failed, Pending, Quarantined, FilteredAsSpam)
    - Configurable result limits (100, 1000, 5000 messages)
    - Detailed message trace events with full delivery timeline
    - View complete transport events and routing information
    - Export search results to Excel with full message details
    - Copy Message ID to clipboard for further investigation
    - Real-time search progress indicators
    - Enter key support for quick searches
    - Note: Distribution group searches show emails sent TO the group address
      (external emails to groups are delivered to individual members and may not appear)
    
    Microsoft 365 - Temporary Access Pass:
    - Generate time-limited passcodes for passwordless authentication onboarding
    - Configurable lifetime (10 to 43,200 minutes / 30 days)
    - One-time use or multi-use TAP options
    - Automatic Microsoft Graph connection with required permissions
    - User lookup by UPN or username
    - One-click copy to clipboard
    - Security warning that TAP is only displayed once
    - Real-time details (user, expiry time, usage type)
    - Use TAP to onboard Windows Hello, Microsoft Authenticator, or FIDO2 keys
    - Enter key support for quick generation
    
    AD Group Members:
    - Search by group name or email
    - View all members with enriched details (title, department)
    - Copy all member emails to clipboard (Outlook format)
    - Export to Excel with full member details

    Employee Conversion:
    - Template-based employee status conversion (e.g., Consultant to FTE)
    - Select "from template" (current status) and "to template" (new status)
    - Surgical group management: only removes groups from "from template"
    - Preserves custom groups not associated with either template
    - Adds all groups from "to template"
    - Automatically moves user to new template's OU
    - Real-time preview of all changes before applying
    - Shows groups to remove, add, and preserve
    - Comprehensive logging of all conversion actions
    - Built-in validation and confirmation dialogs
    - Requires Active Directory PowerShell module (RSAT)
    
    AD Properties Viewer:
    - Double-click any user/group in any grid to view properties
    - General tab: Name, email, title, department, office, company
    - Contact Information tab: Phone, mobile, address
    - Organization tab: Manager, direct reports, group memberships
    - Account tab: UPN, SAM, DN, GUID, creation/modification dates

    IP Network Scanner:
    - Scan custom IP address ranges (start IP to end IP)
    - Real-time connectivity testing using Test-Connection
    - Automatic hostname resolution for discovered devices
    - MAC address lookup via ARP table
    - Response time measurement in milliseconds
    - Live scan progress with online/offline counters
    - Color-coded status display (Green=Online, Red=Offline)
    - Export complete scan results to Excel
    - Support for large IP ranges with progress tracking
    - Warning prompts for scans over 1,000 addresses
    
    Export Active Users Report:
    - One-click export of all enabled AD user accounts
    - Retrieves users from Active and Consultants OUs
    - Automatic filtering of test accounts (containing "test" or "t-" prefix)
    - Exports: Name, SamAccountName, Email, Department, Title, Office (desk location)
    - Results automatically sorted by Name
    - User-selectable save location with file save dialog
    - Excel export with professional formatting, filters, and frozen headers
    - Timestamped filenames for easy organization
    - Requires Active Directory PowerShell module (RSAT)v

.PROXY CONFIGURATION
    The script includes proxy configuration for corporate environments.
    Update the proxy URL in lines 209-211 if needed for your organization.

#>

# Update Script version
$ScriptVersion = "4.7.1"

# Load logo from file path
# Embedded logo (Base64 encoded)
$logoBase64 = @"
iVBORw0KGgoAAAANSUhEUgAABrgAAAGQCAYAAAD1DmBrAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAyZpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2
tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3Jl
IDkuMC1jMDAwIDc5LjE3MWMyN2ZhYiwgMjAyMi8wOC8xNi0yMjozNTo0MSAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi
1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25z
LmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZVJlZiMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYm
UgUGhvdG9zaG9wIDI0LjEgKFdpbmRvd3MpIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOjg2MkEyNDlCOTUyQzExRUQ4M0Q3QUM1NkQxMjY3MkNCIiB4bXBNTTpEb2N1bWVudElEPSJ4
bXAuZGlkOjg2MkEyNDlDOTUyQzExRUQ4M0Q3QUM1NkQxMjY3MkNCIj4gPHhtcE1NOkRlcml2ZWRGcm9tIHN0UmVmOmluc3RhbmNlSUQ9InhtcC5paWQ6ODYyQTI0OTk5NTJDMTFFRDgzRD
dBQzU2RDEyNjcyQ0IiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6ODYyQTI0OUE5NTJDMTFFRDgzRDdBQzU2RDEyNjcyQ0IiLz4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmRmOlJERj4g
PC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz4Urbr9AADF0ElEQVR42uy9CaAkVX3v/6sLAyMMCKIM+yIgIOu4PEQTjYm+7M8lZtHERH1xQ9wx7klcEjXR6Is+E0VxXoyaqC+J/l
+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
TkSuQmCC
"@

# Function to load logo from Base64 string
function Get-LogoFromBase64 {
    param([string]$Base64String)
    
    try {
        # Convert Base64 to byte array
        $imageBytes = [Convert]::FromBase64String($Base64String)
        
        # Create a memory stream from the bytes
        $memoryStream = New-Object System.IO.MemoryStream
        $memoryStream.Write($imageBytes, 0, $imageBytes.Length)
        $memoryStream.Position = 0
        
        # Create BitmapImage from stream
        $logoImage = New-Object System.Windows.Media.Imaging.BitmapImage
        $logoImage.BeginInit()
        $logoImage.StreamSource = $memoryStream
        $logoImage.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $logoImage.EndInit()
        $logoImage.Freeze()
        
        return $logoImage
    } catch {
        Write-Log "Could not load embedded logo: $($_.Exception.Message)"
        return $null
    }
}

#Requires -Modules ExchangeOnlineManagement

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host "Restarting in STA mode..." -ForegroundColor Yellow
    Start-Process powershell.exe -ArgumentList "-STA", "-NoProfile", "-File", "`"$PSCommandPath`"" -Wait
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12  

[system.net.webrequest]::defaultwebproxy = new-object system.net.webproxy('http://proxy.gellerco.com:8080')
[system.net.webrequest]::defaultwebproxy.credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[system.net.webrequest]::defaultwebproxy.BypassProxyOnLocal = $true


# ExchangeOnlineManagement is loaded on-demand when the user connects to Exchange Online
# to avoid Azure.Core.dll version conflicts with Microsoft Graph modules (used by Intune)

# Check if ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
       Write-Host "ImportExcel module installed successfully!" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not install ImportExcel module. Excel export feature will not be available." -ForegroundColor Yellow
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Check current connection status
$existingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue

if ($null -ne $existingConnection -and $existingConnection.State -eq 'Connected') {
#    Write-Host "Existing Exchange Online connection detected!" -ForegroundColor Green
#    Write-Host "Connected as: $($existingConnection.UserPrincipalName)" -ForegroundColor Cyan
#    Write-Host "Connection State: $($existingConnection.State)" -ForegroundColor Cyan
#    Write-Host ""
} else {
#    Write-Host "No Exchange Online connection detected." -ForegroundColor Yellow
#    Write-Host "You can connect later from the GUI if needed." -ForegroundColor Yellow
#    Write-Host ""
}

# Write-Host "Launching GUI..." -ForegroundColor Green
Start-Sleep -Seconds 1

$syncHash = [hashtable]::Synchronized(@{})

function Show-ADPropertiesWindow {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Identity,
        
        [Parameter(Mandatory=$false)]
        [System.Windows.Window]$Owner
    )
    
    try {
        Write-Log "Loading AD properties for: $Identity"
        
        $recipient = Get-Recipient -Identity $Identity -ErrorAction Stop
        
        $isGroup = $false
        $detailedInfo = $null
        
        if ($recipient.RecipientType -like "*Group*") {
            $isGroup = $true
            try {
                $detailedInfo = Get-DistributionGroup -Identity $Identity -ErrorAction Stop
            } catch {
                $detailedInfo = Get-Group -Identity $Identity -ErrorAction Stop
            }
        } else {
            try {
                $detailedInfo = Get-User -Identity $Identity -ErrorAction Stop
            } catch {
                $detailedInfo = Get-Mailbox -Identity $Identity -ErrorAction SilentlyContinue
            }
        }
        
        [xml]$PropertiesXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Properties - $($recipient.DisplayName)" 
        Height="600" 
        Width="550" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <Border Grid.Row="0" Background="#233A4A" Padding="20,15">
            <StackPanel>
                <TextBlock Text="$($recipient.DisplayName)" FontSize="18" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="$($recipient.RecipientType)" FontSize="12" Foreground="#B0BEC5" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <TabControl Grid.Row="1" Margin="10">
            <TabItem Header="General">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <TextBlock Text="Display Name:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="DisplayNameBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Email Address:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="EmailBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock x:Name="TitleLabel" Text="Title:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <TextBox x:Name="TitleBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5" Visibility="Collapsed"/>
                        
                        <TextBlock x:Name="DepartmentLabel" Text="Department:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <TextBox x:Name="DepartmentBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5" Visibility="Collapsed"/>
                        
                        <TextBlock x:Name="OfficeLabel" Text="Office:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <TextBox x:Name="OfficeBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5" Visibility="Collapsed"/>
                        
                        <TextBlock x:Name="CompanyLabel" Text="Company:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <TextBox x:Name="CompanyBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5" Visibility="Collapsed"/>
                        
                        <TextBlock Text="Recipient Type:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="RecipientTypeBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock x:Name="ManagedByLabel" Text="Managed By:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <TextBox x:Name="ManagedByBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5" Visibility="Collapsed"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="Contact Information" x:Name="ContactTab">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <TextBlock Text="Phone:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="PhoneBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Mobile:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="MobileBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Fax:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="FaxBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Street Address:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="StreetBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="City:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="CityBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="State/Province:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="StateBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="ZIP/Postal Code:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="PostalBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Country:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="CountryBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="Organization">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <TextBlock Text="Manager:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="ManagerBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock x:Name="DirectReportsLabel" Text="Direct Reports:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <Border x:Name="DirectReportsBorder" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15" MaxHeight="150" Visibility="Collapsed">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <ItemsControl x:Name="DirectReportsList" Padding="5"/>
                            </ScrollViewer>
                        </Border>
                        
                        <TextBlock x:Name="MemberOfLabel" Text="Member Of (Groups):" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <Border x:Name="MemberOfBorder" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15" MaxHeight="200" Visibility="Collapsed">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <ItemsControl x:Name="MemberOfList" Padding="5"/>
                            </ScrollViewer>
                        </Border>
                        
                        <TextBlock x:Name="MembersLabel" Text="Group Members:" FontWeight="Bold" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <TextBlock x:Name="MembersCount" FontSize="11" Foreground="#666" Margin="0,0,0,5" Visibility="Collapsed"/>
                        <Border x:Name="MembersBorder" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15" MaxHeight="200" Visibility="Collapsed">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <ItemsControl x:Name="MembersList" Padding="5"/>
                            </ScrollViewer>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
            
            <TabItem Header="Account">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15">
                        <TextBlock Text="User Principal Name (UPN):" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="UPNBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="SAM Account Name:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="SAMBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Distinguished Name:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="DNBox" IsReadOnly="True" TextWrapping="Wrap" Margin="0,0,0,15" Padding="5" MinHeight="60"/>
                        
                        <TextBlock Text="Object GUID:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="GUIDBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Created:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="WhenCreatedBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                        
                        <TextBlock Text="Modified:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="WhenChangedBox" IsReadOnly="True" Margin="0,0,0,15" Padding="5"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>
        
        <Border Grid.Row="2" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="0,1,0,0" Padding="15">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="CopyEmailBtn" Content="Copy Email" Width="100" Height="30" Margin="0,0,10,0" Background="#6c757d" Foreground="White"/>
                <Button x:Name="CloseBtn" Content="Close" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
        
        $propReader = New-Object System.Xml.XmlNodeReader $PropertiesXAML
        $PropWindow = [Windows.Markup.XamlReader]::Load($propReader)
        if ($Owner) { $PropWindow.Owner = $Owner }
        
        $DisplayNameBox = $PropWindow.FindName("DisplayNameBox")
        $EmailBox = $PropWindow.FindName("EmailBox")
        $TitleLabel = $PropWindow.FindName("TitleLabel")
        $TitleBox = $PropWindow.FindName("TitleBox")
        $DepartmentLabel = $PropWindow.FindName("DepartmentLabel")
        $DepartmentBox = $PropWindow.FindName("DepartmentBox")
        $OfficeLabel = $PropWindow.FindName("OfficeLabel")
        $OfficeBox = $PropWindow.FindName("OfficeBox")
        $CompanyLabel = $PropWindow.FindName("CompanyLabel")
        $CompanyBox = $PropWindow.FindName("CompanyBox")
        $RecipientTypeBox = $PropWindow.FindName("RecipientTypeBox")
        $ManagedByLabel = $PropWindow.FindName("ManagedByLabel")
        $ManagedByBox = $PropWindow.FindName("ManagedByBox")
        
        $ContactTab = $PropWindow.FindName("ContactTab")
        $PhoneBox = $PropWindow.FindName("PhoneBox")
        $MobileBox = $PropWindow.FindName("MobileBox")
        $FaxBox = $PropWindow.FindName("FaxBox")
        $StreetBox = $PropWindow.FindName("StreetBox")
        $CityBox = $PropWindow.FindName("CityBox")
        $StateBox = $PropWindow.FindName("StateBox")
        $PostalBox = $PropWindow.FindName("PostalBox")
        $CountryBox = $PropWindow.FindName("CountryBox")
        
        $ManagerBox = $PropWindow.FindName("ManagerBox")
        $DirectReportsLabel = $PropWindow.FindName("DirectReportsLabel")
        $DirectReportsBorder = $PropWindow.FindName("DirectReportsBorder")
        $DirectReportsList = $PropWindow.FindName("DirectReportsList")
        $MemberOfLabel = $PropWindow.FindName("MemberOfLabel")
        $MemberOfBorder = $PropWindow.FindName("MemberOfBorder")
        $MemberOfList = $PropWindow.FindName("MemberOfList")
        $MembersLabel = $PropWindow.FindName("MembersLabel")
        $MembersCount = $PropWindow.FindName("MembersCount")
        $MembersBorder = $PropWindow.FindName("MembersBorder")
        $MembersList = $PropWindow.FindName("MembersList")
        
        $UPNBox = $PropWindow.FindName("UPNBox")
        $SAMBox = $PropWindow.FindName("SAMBox")
        $DNBox = $PropWindow.FindName("DNBox")
        $GUIDBox = $PropWindow.FindName("GUIDBox")
        $WhenCreatedBox = $PropWindow.FindName("WhenCreatedBox")
        $WhenChangedBox = $PropWindow.FindName("WhenChangedBox")
        
        $CopyEmailBtn = $PropWindow.FindName("CopyEmailBtn")
        $CloseBtn = $PropWindow.FindName("CloseBtn")
        
        $DisplayNameBox.Text = if ($recipient.DisplayName) { $recipient.DisplayName } else { "" }
        $EmailBox.Text = if ($recipient.PrimarySmtpAddress) { $recipient.PrimarySmtpAddress.ToString() } else { "N/A" }
        $RecipientTypeBox.Text = $recipient.RecipientType
        
        if (-not $isGroup) {
            if ($detailedInfo.Title) {
                $TitleLabel.Visibility = [System.Windows.Visibility]::Visible
                $TitleBox.Visibility = [System.Windows.Visibility]::Visible
                $TitleBox.Text = $detailedInfo.Title
            }
            if ($detailedInfo.Department) {
                $DepartmentLabel.Visibility = [System.Windows.Visibility]::Visible
                $DepartmentBox.Visibility = [System.Windows.Visibility]::Visible
                $DepartmentBox.Text = $detailedInfo.Department
            }
            if ($detailedInfo.Office) {
                $OfficeLabel.Visibility = [System.Windows.Visibility]::Visible
                $OfficeBox.Visibility = [System.Windows.Visibility]::Visible
                $OfficeBox.Text = $detailedInfo.Office
            }
            if ($detailedInfo.Company) {
                $CompanyLabel.Visibility = [System.Windows.Visibility]::Visible
                $CompanyBox.Visibility = [System.Windows.Visibility]::Visible
                $CompanyBox.Text = $detailedInfo.Company
            }
            
            $PhoneBox.Text = if ($detailedInfo.Phone) { $detailedInfo.Phone } else { "" }
            $MobileBox.Text = if ($detailedInfo.MobilePhone) { $detailedInfo.MobilePhone } else { "" }
            $FaxBox.Text = if ($detailedInfo.Fax) { $detailedInfo.Fax } else { "" }
            $StreetBox.Text = if ($detailedInfo.StreetAddress) { $detailedInfo.StreetAddress } else { "" }
            $CityBox.Text = if ($detailedInfo.City) { $detailedInfo.City } else { "" }
            $StateBox.Text = if ($detailedInfo.StateOrProvince) { $detailedInfo.StateOrProvince } else { "" }
            $PostalBox.Text = if ($detailedInfo.PostalCode) { $detailedInfo.PostalCode } else { "" }
            $CountryBox.Text = if ($detailedInfo.CountryOrRegion) { $detailedInfo.CountryOrRegion } else { "" }
            
            if ($detailedInfo.Manager) {
                try {
                    $mgr = Get-Recipient -Identity $detailedInfo.Manager -ErrorAction SilentlyContinue
                    $ManagerBox.Text = if ($mgr) { $mgr.DisplayName } else { $detailedInfo.Manager }
                } catch {
                    $ManagerBox.Text = $detailedInfo.Manager
                }
            }
            
            if ($detailedInfo.DirectReports -and $detailedInfo.DirectReports.Count -gt 0) {
				$DirectReportsLabel.Visibility = [System.Windows.Visibility]::Visible
				$DirectReportsBorder.Visibility = [System.Windows.Visibility]::Visible
				
				# Get and sort direct reports
				$sortedReports = @()
				foreach ($dr in $detailedInfo.DirectReports) {
					try {
						$drRecip = Get-Recipient -Identity $dr -ErrorAction SilentlyContinue
						$sortedReports += if ($drRecip) { $drRecip.DisplayName } else { $dr }
					} catch {
						$sortedReports += $dr
					}
				}
				
				# Sort alphabetically
				$sortedReports = $sortedReports | Sort-Object
				
				foreach ($name in $sortedReports) {
					$tb = New-Object System.Windows.Controls.TextBlock
					$tb.Text = "Ã¢â‚¬Â¢ $name"
					$tb.Margin = "0,2,0,2"
					$DirectReportsList.Items.Add($tb) | Out-Null
				}
			}
            
            try {
				$groups = Get-Recipient -Identity $Identity -ErrorAction Stop | Select-Object -ExpandProperty DistinguishedName | ForEach-Object {
					Get-Recipient -Filter "Members -eq '$_'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup -ErrorAction SilentlyContinue
				}
				
				if ($groups -and $groups.Count -gt 0) {
					$MemberOfLabel.Visibility = [System.Windows.Visibility]::Visible
					$MemberOfBorder.Visibility = [System.Windows.Visibility]::Visible
					
					# Sort groups alphabetically
					$sortedGroups = $groups | Sort-Object DisplayName
					
					foreach ($grp in $sortedGroups) {
						$tb = New-Object System.Windows.Controls.TextBlock
						$tb.Text = "Ã¢â‚¬Â¢ $($grp.DisplayName)"
						$tb.Margin = "0,2,0,2"
						$MemberOfList.Items.Add($tb) | Out-Null
					}
				}
			} catch {}
            
        } else {
            $ContactTab.Visibility = [System.Windows.Visibility]::Collapsed
            
            if ($detailedInfo.ManagedBy) {
                $ManagedByLabel.Visibility = [System.Windows.Visibility]::Visible
                $ManagedByBox.Visibility = [System.Windows.Visibility]::Visible
                try {
                    $mgr = Get-Recipient -Identity $detailedInfo.ManagedBy[0] -ErrorAction SilentlyContinue
                    $ManagedByBox.Text = if ($mgr) { $mgr.DisplayName } else { $detailedInfo.ManagedBy[0] }
                } catch {
                    $ManagedByBox.Text = $detailedInfo.ManagedBy[0]
                }
            }
            
            try {
				$members = Get-DistributionGroupMember -Identity $Identity -ErrorAction Stop
				if ($members -and $members.Count -gt 0) {
					$MembersLabel.Visibility = [System.Windows.Visibility]::Visible
					$MembersCount.Visibility = [System.Windows.Visibility]::Visible
					$MembersBorder.Visibility = [System.Windows.Visibility]::Visible
					$MembersCount.Text = "Total: $($members.Count) members"
					
					# Sort members alphabetically
					$sortedMembers = $members | Sort-Object DisplayName | Select-Object -First 50
					
					foreach ($mem in $sortedMembers) {
						$tb = New-Object System.Windows.Controls.TextBlock
						$tb.Text = "Ã¢â‚¬Â¢ $($mem.DisplayName)"
						$tb.Margin = "0,2,0,2"
						$MembersList.Items.Add($tb) | Out-Null
					}
					
					if ($members.Count -gt 50) {
						$tb = New-Object System.Windows.Controls.TextBlock
						$tb.Text = "... and $($members.Count - 50) more"
						$tb.Margin = "0,2,0,2"
						$tb.FontStyle = [System.Windows.FontStyles]::Italic
						$tb.Foreground = [System.Windows.Media.Brushes]::Gray
						$MembersList.Items.Add($tb) | Out-Null
					}
				}
			} catch {}
        }
        
        $UPNBox.Text = if ($recipient.UserPrincipalName) { $recipient.UserPrincipalName } else { "" }
        $SAMBox.Text = if ($detailedInfo.SamAccountName) { $detailedInfo.SamAccountName } else { "" }
        $DNBox.Text = if ($recipient.DistinguishedName) { $recipient.DistinguishedName } else { "" }
        $GUIDBox.Text = if ($recipient.Guid) { $recipient.Guid.ToString() } else { "" }
        $WhenCreatedBox.Text = if ($detailedInfo.WhenCreated) { $detailedInfo.WhenCreated.ToString() } else { "" }
        $WhenChangedBox.Text = if ($detailedInfo.WhenChanged) { $detailedInfo.WhenChanged.ToString() } else { "" }
        
        $CopyEmailBtn.Add_Click({
            if ($EmailBox.Text -and $EmailBox.Text -ne "N/A") {
                [System.Windows.Forms.Clipboard]::SetText($EmailBox.Text)
                [System.Windows.MessageBox]::Show("Email address copied to clipboard!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        })
        
        $CloseBtn.Add_Click({ $PropWindow.Close() })
        
        $PropWindow.ShowDialog() | Out-Null
        
    } catch {
        Write-Log "Error loading properties: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Error loading properties:`n`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="IT Operations Center" 
        Height="750" 
        Width="850" 
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanMinimize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="100"/>
        </Grid.RowDefinitions>
        
        <!-- Header Section with Logo -->
        <Border Grid.Row="0" Background="#233A4A" Padding="15">
            <DockPanel>
                <Image x:Name="CompanyLogo" 
                       DockPanel.Dock="Left"
                       Width="200" 
                       Height="50" 
                       Margin="0,0,15,0"
                       Stretch="Uniform"
                       VerticalAlignment="Center"/>
                <StackPanel DockPanel.Dock="Left" VerticalAlignment="Center">
                    <TextBlock Text="IT Operations Center" 
                              FontSize="20" 
                              FontWeight="Bold" 
                              Foreground="White"/>
                </StackPanel>
                <TextBlock x:Name="VersionText"
                          Text="v4.2.0"
                          FontSize="11"
                          Foreground="#B0BEC5"
                          VerticalAlignment="Bottom"
                          HorizontalAlignment="Right"
                          DockPanel.Dock="Right"/>
            </DockPanel>
        </Border>
        
        
        <!-- Main Content Area -->
        <Border Grid.Row="1" Background="WhiteSmoke" Padding="20">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <!-- Connection Status -->
                <GroupBox Grid.Row="0" Header="Connection Status" Padding="10" Margin="0,0,0,15">
                    <StackPanel>
                        <DockPanel Margin="0,0,0,10">
                            <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
                                <Ellipse x:Name="ConnectionStatusIndicator" 
                                        Width="12" 
                                        Height="12" 
                                        Fill="Red" 
                                        Margin="0,0,10,0"
                                        VerticalAlignment="Center"/>
                                <TextBlock x:Name="ConnectionStatusText" 
                                        Text="Not Connected" 
                                        FontWeight="Bold"
                                        VerticalAlignment="Center"
                                        Foreground="Red"/>
                            </StackPanel>
                        </DockPanel>
                        
                        <TextBlock x:Name="ConnectionInfoText" 
                                Text="Click 'Connect' to authenticate to Exchange Online" 
                                TextWrapping="Wrap" 
                                Margin="0,0,0,15"
                                FontSize="11"
                                Foreground="#666"/>
                        
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                            <Button x:Name="ConnectButton" 
                                Content="Connect to Exchange Online" 
                                Width="180" 
                                Height="35" 
                                Margin="5"
                                Background="#28a745"
                                Foreground="White"
                                FontWeight="Bold"
                                Cursor="Hand"/>
                            <Button x:Name="DisconnectButton" 
                                Content="Disconnect" 
                                Width="120" 
                                Height="35" 
                                Margin="5"
                                Background="#dc3545"
                                Foreground="White"
                                FontWeight="Bold"
                                Cursor="Hand"
                                IsEnabled="False"/>
                        </StackPanel>
                    </StackPanel>
                </GroupBox>
                
<!-- REORGANIZED MANAGEMENT OPTIONS -->
<!-- This replaces the Management Options section starting around line 1269 -->

                <GroupBox Grid.Row="1" Header="Management Options" 
                        x:Name="ManagementGroup" 
                        Padding="10">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel>
                            
                            <!-- Exchange Online Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#007bff" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Exchange Online" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="MailboxButton" 
                                        Content="Mailbox Permissions (Full Access &amp; Send As)" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="CalendarButton" 
                                        Content="Calendar Permissions" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="AutoRepliesButton" 
                                        Content="Automatic Replies (Out of Office)" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="MessageTraceButton" 
                                        Content="Message Trace / Tracking" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="SendOnBehalfButton" 
                                        Content="Send on Behalf Permissions" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                    <Button x:Name="ForwardingButton" 
                                        Content="Email Forwarding Management" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                    <Button x:Name="ResourceMailboxButton" 
                                        Content="Resource Mailbox Management" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                </StackPanel>
                            </Expander>
                            
                            <!-- Microsoft 365 Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#5e72e4" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Microsoft 365" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="GenerateTAPButton" 
                                        Content="Generate Temporary Access Pass" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                </StackPanel>
                            </Expander>
                            
                            <!-- Active Directory Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#28a745" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Active Directory" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="GroupMembersButton" 
                                        Content="AD Group Members Viewer" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="ExportActiveUsersButton" 
                                        Content="Export Active Users Report" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="EmployeeConversionButton" 
                                        Content="Employee Conversion" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="LockedOutUsersButton" 
                                        Content="Locked Out Users" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="DistributionGroupButton" 
                                        Content="Distribution List Management" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                </StackPanel>
                            </Expander>
                            
                            <!-- Reports & Analytics Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#17a2b8" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Reports &amp; Analytics" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="MailboxStatsButton" 
                                        Content="Mailbox Size &amp; Quota Report" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="PermissionAuditButton" 
                                        Content="Permission Audit Report" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                </StackPanel>
                            </Expander>
                            
                            <!-- Device Management Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#FF9800" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Device Management" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="IntuneMobileButton" 
                                        Content="Intune Mobile Devices" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                    <Button x:Name="SCCMDevicesButton" 
                                        Content="SCCM Device Management" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                    <Button x:Name="IntuneComplianceButton" 
                                        Content="Compliance Policy Reports" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                </StackPanel>
                            </Expander>
                            
                            <!-- Network & Infrastructure Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#6f42c1" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Network &amp; Infrastructure" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="IPScannerButton" 
                                        Content="IP Network Scanner" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#f8f9fa"
                                        BorderBrush="#dee2e6"/>
                                </StackPanel>
                            </Expander>
                            
                            <!-- Compliance & Security Section -->
                            <Expander IsExpanded="False" Margin="0,0,0,5">
                                <Expander.Header>
                                    <StackPanel Orientation="Horizontal">
                                        <Ellipse Width="12" Height="12" Fill="#dc3545" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                        <TextBlock Text="Compliance &amp; Security" FontWeight="Bold" FontSize="13"/>
                                    </StackPanel>
                                </Expander.Header>
                                <StackPanel Margin="20,5,0,5">
                                    <Button x:Name="LitigationHoldButton" 
                                        Content="Litigation Hold Management" 
                                        Height="32" 
                                        Margin="0,5"
                                        HorizontalAlignment="Stretch"
                                        HorizontalContentAlignment="Left"
                                        Padding="10,0"
                                        Background="#e9ecef"
                                        BorderBrush="#dee2e6"
                                        Foreground="#6c757d"
                                        IsEnabled="False"
                                        ToolTip="Coming in future version"/>
                                </StackPanel>
                            </Expander>

                        </StackPanel>
                    </ScrollViewer>
                </GroupBox>
            </Grid>
        </Border>
        
        <!-- Status/Log Section -->
        <Border Grid.Row="2" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="0,1,0,0">
            <DockPanel Margin="10">
                <TextBlock Text="Activity Log:" 
                          DockPanel.Dock="Top" 
                          FontWeight="Bold" 
                          Margin="0,0,0,5"/>
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <TextBox x:Name="LogBox" 
                            IsReadOnly="True" 
                            Background="Transparent" 
                            BorderThickness="0"
                            TextWrapping="Wrap"/>
                </ScrollViewer>
            </DockPanel>
        </Border>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($reader)

$syncHash.Window = $Window

$syncHash.ConnectionStatusIndicator = $Window.FindName("ConnectionStatusIndicator")
$syncHash.ConnectionStatusText = $Window.FindName("ConnectionStatusText")
$syncHash.ConnectionInfoText = $Window.FindName("ConnectionInfoText")
$syncHash.ConnectButton = $Window.FindName("ConnectButton")
$syncHash.DisconnectButton = $Window.FindName("DisconnectButton")
$syncHash.StatusText = $Window.FindName("StatusText")
$syncHash.LogBox = $Window.FindName("LogBox")
$syncHash.ManagementGroup = $Window.FindName("ManagementGroup")
$syncHash.MailboxButton = $Window.FindName("MailboxButton")
$syncHash.CalendarButton = $Window.FindName("CalendarButton")
$syncHash.GroupMembersButton = $Window.FindName("GroupMembersButton")
$syncHash.VersionText = $Window.FindName("VersionText")
$syncHash.AutoRepliesButton = $Window.FindName("AutoRepliesButton")

# Microsoft 365
$syncHash.GenerateTAPButton = $Window.FindName("GenerateTAPButton")

# Mailbox Management - Future
$syncHash.SendOnBehalfButton = $Window.FindName("SendOnBehalfButton")
$syncHash.ForwardingButton = $Window.FindName("ForwardingButton")

# Calendar & Resources - Future
$syncHash.ResourceMailboxButton = $Window.FindName("ResourceMailboxButton")

# Groups & Distribution - Future
$syncHash.ExportActiveUsersButton = $Window.FindName("ExportActiveUsersButton")
$syncHash.EmployeeConversionButton = $Window.FindName("EmployeeConversionButton")
$syncHash.LockedOutUsersButton = $Window.FindName("LockedOutUsersButton")
$syncHash.DistributionGroupButton = $Window.FindName("DistributionGroupButton")

# Compliance & Security - Future
$syncHash.MessageTraceButton = $Window.FindName("MessageTraceButton")
$syncHash.LitigationHoldButton = $Window.FindName("LitigationHoldButton")

# Reports & Analytics - Future
$syncHash.MailboxStatsButton = $Window.FindName("MailboxStatsButton")
$syncHash.PermissionAuditButton = $Window.FindName("PermissionAuditButton")

# Network & Infrastructure
$syncHash.IPScannerButton = $Window.FindName("IPScannerButton")
$syncHash.IntuneMobileButton = $Window.FindName("IntuneMobileButton")
$syncHash.SCCMDevicesButton = $Window.FindName("SCCMDevicesButton")

$syncHash.IntuneComplianceButton = $Window.FindName("IntuneComplianceButton")



# Set dynamic version text
$syncHash.Window.Title = "IT Operations Center v$ScriptVersion"
$syncHash.VersionText.Text = "v$ScriptVersion"

function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logEntry = "[$timestamp] $Message`r`n"
    $syncHash.Window.Dispatcher.Invoke([action]{
        $syncHash.LogBox.AppendText($logEntry)
    })
}

# Function to update connection status in GUI
function Update-ConnectionStatus {
    $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    
    if ($null -ne $connInfo -and $connInfo.State -eq 'Connected') {
        $syncHash.ConnectionStatusIndicator.Fill = [System.Windows.Media.Brushes]::Green
        $syncHash.ConnectionStatusText.Text = "Connected"
        $syncHash.ConnectionStatusText.Foreground = [System.Windows.Media.Brushes]::Green
        $syncHash.ConnectionInfoText.Text = "Connected as: $($connInfo.UserPrincipalName)"
        $syncHash.ConnectButton.IsEnabled = $false
        $syncHash.DisconnectButton.IsEnabled = $true
        Write-Log "Connected to Exchange Online as $($connInfo.UserPrincipalName)"
    } else {
        $syncHash.ConnectionStatusIndicator.Fill = [System.Windows.Media.Brushes]::Red
        $syncHash.ConnectionStatusText.Text = "Not Connected"
        $syncHash.ConnectionStatusText.Foreground = [System.Windows.Media.Brushes]::Red
        $syncHash.ConnectionInfoText.Text = "Click 'Connect' to authenticate to Exchange Online"
        $syncHash.ConnectButton.IsEnabled = $true
        $syncHash.DisconnectButton.IsEnabled = $false
        Write-Log "Not connected to Exchange Online"
    }
}

# Load embedded logo
try {
    $logoImage = Get-LogoFromBase64 -Base64String $logoBase64
    if ($null -ne $logoImage) {
        $CompanyLogo = $Window.FindName("CompanyLogo")
        $CompanyLogo.Source = $logoImage
        Write-Log "Company logo loaded from embedded data"
    }
} catch {
    Write-Log "Could not load embedded company logo: $($_.Exception.Message)"
}

# Update connection status on load
Update-ConnectionStatus

$syncHash.ConnectButton.Add_Click({
    Write-Log "Initiating Exchange Online connection..."
    
    # Minimize the GUI window
    $syncHash.Window.WindowState = [System.Windows.WindowState]::Minimized
    
    <#
    Show console message
    $host.UI.RawUI.ForegroundColor = "Cyan"
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "CONNECTING TO EXCHANGE ONLINE" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "A browser window will open for authentication..." -ForegroundColor Yellow
    Write-Host "Please complete the authentication process." -ForegroundColor Yellow
    Write-Host ""
    #>

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
#       Write-Host ""
       Write-Host "Successfully connected to Exchange Online!" -ForegroundColor Green
#        Write-Host ""
#        Write-Host "Returning to GUI..." -ForegroundColor Green

        # Restore the GUI window
        $syncHash.Window.WindowState = [System.Windows.WindowState]::Normal
        $syncHash.Window.Activate()
        
        # Update connection status
        Update-ConnectionStatus
        
        [System.Windows.MessageBox]::Show(
            "Successfully connected to Exchange Online!",
            "Connected",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
        
    } catch {
#        Write-Host ""
        Write-Host "Connection failed: $($_.Exception.Message)" -ForegroundColor Red
#        Write-Host ""
#        Write-Host "Press ENTER to return to the GUI..." -ForegroundColor Yellow
        $null = Read-Host
        
        # Restore the GUI window
        $syncHash.Window.WindowState = [System.Windows.WindowState]::Normal
        $syncHash.Window.Activate()
        
        Write-Log "Connection failed: $($_.Exception.Message)"
        
        [System.Windows.MessageBox]::Show(
            "Failed to connect to Exchange Online:`n`n$($_.Exception.Message)`n`nPlease try again.",
            "Connection Failed",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
})

$syncHash.DisconnectButton.Add_Click({
    Write-Log "Disconnecting from Exchange Online..."
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Log "Successfully disconnected from Exchange Online"
        
        Update-ConnectionStatus
        
        [System.Windows.MessageBox]::Show(
            "Disconnected from Exchange Online.",
            "Disconnected",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
        
    } catch {
        Write-Log "Disconnect error: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show(
            "Error disconnecting: $($_.Exception.Message)",
            "Disconnect Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
    }
})

$syncHash.MailboxButton.Add_Click({
    # Check if connected
    $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
        $result = [System.Windows.MessageBox]::Show(
            "You are not connected to Exchange Online.`n`nWould you like to connect now?",
            "Connection Required",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $syncHash.ConnectButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            
            # Wait a moment for connection to establish, then recheck
            Start-Sleep -Seconds 2
            $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
            
            if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
                Write-Log "Connection was not established. Please try again."
                return
            }
        } else {
            return
        }
    }
    
    Write-Log "Opening Mailbox Permissions window..."
    
    function Resolve-UserDisplayName {
        param($Identity)
        
        $identityStr = $Identity.ToString()
        
        if ($identityStr -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
            try {
                $group = Get-DistributionGroup -Identity $identityStr -ErrorAction SilentlyContinue
                if ($group) { return $group.DisplayName }
                
                $group = Get-Group -Identity $identityStr -ErrorAction SilentlyContinue
                if ($group) { return $group.DisplayName }
                
                $recipient = Get-Recipient -Identity $identityStr -ErrorAction SilentlyContinue
                if ($recipient) { return $recipient.DisplayName }
            } catch {
                Write-Log "Could not resolve GUID: $identityStr"
            }
        }
        
        return $identityStr
    }
    
    function Get-CombinedMailboxPermissions {
    param($MailboxIdentity)
    
    try {
        Write-Log "Retrieving Full Access permissions for $MailboxIdentity"
        $fullAccessPerms = @(Get-MailboxPermission -Identity $MailboxIdentity -ErrorAction Stop | 
            Where-Object {$_.User -notlike "NT AUTHORITY\*" -and $_.User -notlike "S-1-5-*" -and $_.IsInherited -eq $false -and $_.AccessRights -contains "FullAccess"})
        
        Write-Log "Retrieving Send As permissions for $MailboxIdentity"
        $sendAsPerms = @(Get-RecipientPermission -Identity $MailboxIdentity -ErrorAction Stop | 
            Where-Object {$_.Trustee -notlike "NT AUTHORITY\*" -and $_.Trustee -notlike "S-1-5-*" -and $_.AccessRights -contains "SendAs"})
        
        $allUsers = @{}
        
        # Process Full Access permissions
        if ($fullAccessPerms.Count -gt 0) {
            foreach ($perm in $fullAccessPerms) {
                try {
                    $userKey = $perm.User.ToString()
                    
                    # Skip if null or empty
                    if ([string]::IsNullOrWhiteSpace($userKey)) {
                        Write-Log "Warning: Skipping null/empty Full Access user"
                        continue
                    }
                    
                    $displayName = Resolve-UserDisplayName -Identity $userKey
                    
                    if (-not $allUsers.ContainsKey($userKey)) {
                        $allUsers[$userKey] = [PSCustomObject]@{
                            User = $displayName
                            UserIdentity = $userKey
                            HasFullAccess = $true
                            HasSendAs = $false
                        }
                    } else {
                        $allUsers[$userKey].HasFullAccess = $true
                    }
                } catch {
                    Write-Log "Warning: Could not process Full Access permission for user: $($_.Exception.Message)"
                    continue
                }
            }
        }
        
        # Process Send As permissions
        if ($sendAsPerms.Count -gt 0) {
            foreach ($perm in $sendAsPerms) {
                try {
                    $userKey = $perm.Trustee.ToString()
                    
                    # Skip if null or empty
                    if ([string]::IsNullOrWhiteSpace($userKey)) {
                        Write-Log "Warning: Skipping null/empty Send As trustee"
                        continue
                    }
                    
                    $displayName = Resolve-UserDisplayName -Identity $userKey
                    
                    if (-not $allUsers.ContainsKey($userKey)) {
                        $allUsers[$userKey] = [PSCustomObject]@{
                            User = $displayName
                            UserIdentity = $userKey
                            HasFullAccess = $false
                            HasSendAs = $true
                        }
                    } else {
                        $allUsers[$userKey].HasSendAs = $true
                    }
                } catch {
                    Write-Log "Warning: Could not process Send As permission for trustee: $($_.Exception.Message)"
                    continue
                }
            }
        }
        
        # Return sorted results, or empty array if no permissions
        if ($allUsers.Count -gt 0) {
            $result = @($allUsers.Values | Sort-Object User)
            return ,$result  # Comma forces PowerShell to return as array
        } else {
            Write-Log "No delegated permissions found for $MailboxIdentity"
            return @()
        }
        
    } catch {
        Write-Log "Error in Get-CombinedMailboxPermissions: $($_.Exception.Message)"
        throw
    }
}
    
    [xml]$MailboxXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Mailbox Permissions Management" 
        Height="600" 
        Width="700" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <TabControl Margin="10">
            <TabItem Header="Add Permission">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="MbxAddMailboxBox" Height="25" Padding="5"/>
                    </StackPanel>
                    
                    <StackPanel Grid.Row="1" Margin="0,0,0,15">
                        <TextBlock Text="User Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="MbxAddUserBox" Height="25" Padding="5"/>
                    </StackPanel>
                    
                    <StackPanel Grid.Row="2" Margin="0,0,0,15">
                        <TextBlock Text="Access Rights:" FontWeight="Bold" Margin="0,0,0,10"/>
                        <CheckBox x:Name="MbxFullAccessCheck" Content="Full Access" Margin="0,0,0,8" FontSize="13"/>
                        <CheckBox x:Name="MbxSendAsCheck" Content="Send As" Margin="0,0,0,8" FontSize="13"/>
                        <TextBlock Text="(Select one or both)" FontStyle="Italic" FontSize="11" Foreground="Gray" Margin="0,5,0,0"/>
                    </StackPanel>
                    
                    <Border Grid.Row="3" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="1" Padding="10" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Margin="0,0,0,8">
                                <Bold>Full Access:</Bold> User can open and view the mailbox contents.
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11">
                                <Bold>Send As:</Bold> User can send emails as if they were the mailbox owner.
                            </TextBlock>
                        </StackPanel>
                    </Border>
                    
                    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="MbxAddButton" Content="Add Permission" Width="120" Height="30" Margin="0,0,10,0" Background="#28a745" Foreground="White" FontWeight="Bold"/>
                        <Button x:Name="MbxAddCancelButton" Content="Cancel" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="View/Edit">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <DockPanel>
                            <Button x:Name="MbxLoadPermissionsButton" Content="Load" Width="80" Height="25" Margin="10,0,0,0" DockPanel.Dock="Right" Background="#007bff" Foreground="White" FontWeight="Bold"/>
                            <TextBox x:Name="MbxViewMailboxBox" Height="25" Padding="5"/>
                        </DockPanel>
                    </StackPanel>
                    
                    <Border Grid.Row="1" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15">
                        <DataGrid x:Name="MbxPermissionsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="User" Binding="{Binding User}" Width="2*"/>
                                <DataGridCheckBoxColumn Header="Full Access" Binding="{Binding HasFullAccess}" Width="*"/>
                                <DataGridCheckBoxColumn Header="Send As" Binding="{Binding HasSendAs}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Border>
                    
                    <GroupBox Grid.Row="2" Header="Edit Selected" Padding="10" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock x:Name="MbxEditUserLabel" Text="No user selected" Margin="0,0,0,10"/>
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                <CheckBox x:Name="MbxEditFullAccessCheck" Content="Full Access" Margin="0,0,20,0" IsEnabled="False"/>
                                <CheckBox x:Name="MbxEditSendAsCheck" Content="Send As" IsEnabled="False"/>
                            </StackPanel>
                            <Button x:Name="MbxUpdateButton" Content="Update Permissions" Width="140" Height="25" HorizontalAlignment="Left" Background="#ffc107" Foreground="Black" FontWeight="Bold" IsEnabled="False"/>
                        </StackPanel>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="MbxExportToExcelButton" Content="Export to Excel" Width="120" Height="30" Margin="0,0,10,0" Background="#17a2b8" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="MbxViewCloseButton" Content="Close" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Remove">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <DockPanel>
                            <Button x:Name="MbxLoadRemovePermissionsButton" Content="Load" Width="80" Height="25" Margin="10,0,0,0" DockPanel.Dock="Right" Background="#007bff" Foreground="White" FontWeight="Bold"/>
                            <TextBox x:Name="MbxRemoveMailboxBox" Height="25" Padding="5"/>
                        </DockPanel>
                    </StackPanel>
                    
                    <Border Grid.Row="1" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15">
                        <DataGrid x:Name="MbxRemovePermissionsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="User" Binding="{Binding User}" Width="2*"/>
                                <DataGridCheckBoxColumn Header="Full Access" Binding="{Binding HasFullAccess}" Width="*"/>
                                <DataGridCheckBoxColumn Header="Send As" Binding="{Binding HasSendAs}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Border>
                    
                    <StackPanel Grid.Row="2">
                        <TextBlock Text="Select which permissions to remove:" FontWeight="Bold" Margin="0,0,0,10"/>
                        <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                            <CheckBox x:Name="MbxRemoveFullAccessCheck" Content="Remove Full Access" Margin="0,0,20,0" IsEnabled="False"/>
                            <CheckBox x:Name="MbxRemoveSendAsCheck" Content="Remove Send As" IsEnabled="False"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button x:Name="MbxRemoveButton" Content="Remove Selected" Width="130" Height="30" Margin="0,0,10,0" Background="#dc3545" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                            <Button x:Name="MbxRemoveCloseButton" Content="Close" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
                        </StackPanel>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@
    
    $mbxReader = New-Object System.Xml.XmlNodeReader $MailboxXAML
    $MbxWindow = [Windows.Markup.XamlReader]::Load($mbxReader)
    $MbxWindow.Owner = $syncHash.Window
    
    $MbxAddMailboxBox = $MbxWindow.FindName("MbxAddMailboxBox")
    $MbxAddUserBox = $MbxWindow.FindName("MbxAddUserBox")
    $MbxFullAccessCheck = $MbxWindow.FindName("MbxFullAccessCheck")
    $MbxSendAsCheck = $MbxWindow.FindName("MbxSendAsCheck")
    $MbxAddButton = $MbxWindow.FindName("MbxAddButton")
    $MbxAddCancelButton = $MbxWindow.FindName("MbxAddCancelButton")
    
    $MbxViewMailboxBox = $MbxWindow.FindName("MbxViewMailboxBox")
    $MbxLoadPermissionsButton = $MbxWindow.FindName("MbxLoadPermissionsButton")
    $MbxPermissionsGrid = $MbxWindow.FindName("MbxPermissionsGrid")
    $MbxEditUserLabel = $MbxWindow.FindName("MbxEditUserLabel")
    $MbxEditFullAccessCheck = $MbxWindow.FindName("MbxEditFullAccessCheck")
    $MbxEditSendAsCheck = $MbxWindow.FindName("MbxEditSendAsCheck")
    $MbxUpdateButton = $MbxWindow.FindName("MbxUpdateButton")
    $MbxExportToExcelButton = $MbxWindow.FindName("MbxExportToExcelButton")
    $MbxViewCloseButton = $MbxWindow.FindName("MbxViewCloseButton")
    
    $MbxRemoveMailboxBox = $MbxWindow.FindName("MbxRemoveMailboxBox")
    $MbxLoadRemovePermissionsButton = $MbxWindow.FindName("MbxLoadRemovePermissionsButton")
    $MbxRemovePermissionsGrid = $MbxWindow.FindName("MbxRemovePermissionsGrid")
    $MbxRemoveFullAccessCheck = $MbxWindow.FindName("MbxRemoveFullAccessCheck")
    $MbxRemoveSendAsCheck = $MbxWindow.FindName("MbxRemoveSendAsCheck")
    $MbxRemoveButton = $MbxWindow.FindName("MbxRemoveButton")
    $MbxRemoveCloseButton = $MbxWindow.FindName("MbxRemoveCloseButton")
    
    $MbxPermissionsGrid.Add_MouseDoubleClick({
        if ($MbxPermissionsGrid.SelectedItem) {
            $selectedUser = $MbxPermissionsGrid.SelectedItem.UserIdentity
            Show-ADPropertiesWindow -Identity $selectedUser -Owner $MbxWindow
        }
    })
    
    $MbxRemovePermissionsGrid.Add_MouseDoubleClick({
        if ($MbxRemovePermissionsGrid.SelectedItem) {
            $selectedUser = $MbxRemovePermissionsGrid.SelectedItem.UserIdentity
            Show-ADPropertiesWindow -Identity $selectedUser -Owner $MbxWindow
        }
    })
    
    $MbxAddCancelButton.Add_Click({ $MbxWindow.Close() })
    
    $MbxAddButton.Add_Click({
        $mailbox = $MbxAddMailboxBox.Text.Trim()
        $user = $MbxAddUserBox.Text.Trim()
        $addFullAccess = $MbxFullAccessCheck.IsChecked
        $addSendAs = $MbxSendAsCheck.IsChecked
        
        if ([string]::IsNullOrWhiteSpace($mailbox) -or [string]::IsNullOrWhiteSpace($user)) {
            [System.Windows.MessageBox]::Show("Please enter both mailbox and user email addresses", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not $addFullAccess -and -not $addSendAs) {
            [System.Windows.MessageBox]::Show("Please select at least one permission type", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $MbxAddButton.IsEnabled = $false
            $successMessages = @()
            
            if ($addFullAccess) {
                Write-Log "Adding Full Access for $user on $mailbox"
                Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping $false -ErrorAction Stop
                $successMessages += "Full Access"
            }
            
            if ($addSendAs) {
                Write-Log "Adding Send As for $user on $mailbox"
                Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                $successMessages += "Send As"
            }
            
            Write-Log "Successfully added permissions: $($successMessages -join ', ')"
            [System.Windows.MessageBox]::Show("Permissions added successfully!`n`n$($successMessages -join ', ')", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            
            $MbxAddMailboxBox.Clear()
            $MbxAddUserBox.Clear()
            $MbxFullAccessCheck.IsChecked = $false
            $MbxSendAsCheck.IsChecked = $false
            
        } catch {
            Write-Log "Error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } finally {
            $MbxAddButton.IsEnabled = $true
        }
    })
    
    $MbxLoadPermissionsButton.Add_Click({
        $mailbox = $MbxViewMailboxBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($mailbox)) { return }
        
        try {
            $MbxLoadPermissionsButton.IsEnabled = $false
            Write-Log "Loading permissions for $mailbox"
            
            $perms = Get-CombinedMailboxPermissions -MailboxIdentity $mailbox
            $MbxPermissionsGrid.ItemsSource = $perms
            
            if ($perms -and $perms.Count -gt 0) {
                $MbxExportToExcelButton.IsEnabled = $true
            } else {
                $MbxExportToExcelButton.IsEnabled = $false
            }
            
            Write-Log "Loaded $($perms.Count) user permissions"
        } catch {
            Write-Log "Error loading permissions: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } finally {
            $MbxLoadPermissionsButton.IsEnabled = $true
        }
    })
    
    $MbxPermissionsGrid.Add_SelectionChanged({
        if ($MbxPermissionsGrid.SelectedItem) {
            $sel = $MbxPermissionsGrid.SelectedItem
            $MbxEditUserLabel.Text = "Editing: $($sel.User)"
            $MbxEditFullAccessCheck.IsEnabled = $true
            $MbxEditSendAsCheck.IsEnabled = $true
            $MbxEditFullAccessCheck.IsChecked = $sel.HasFullAccess
            $MbxEditSendAsCheck.IsChecked = $sel.HasSendAs
            $MbxUpdateButton.IsEnabled = $true
        }
    })
    
    $MbxUpdateButton.Add_Click({
        $mailbox = $MbxViewMailboxBox.Text.Trim()
        $sel = $MbxPermissionsGrid.SelectedItem
        $newFullAccess = $MbxEditFullAccessCheck.IsChecked
        $newSendAs = $MbxEditSendAsCheck.IsChecked
        
        if ($null -eq $sel) { return }
        
        if (-not $newFullAccess -and -not $newSendAs) {
            [System.Windows.MessageBox]::Show("At least one permission must be selected. Use the Remove tab to remove all permissions.", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $result = [System.Windows.MessageBox]::Show("Update permissions for $($sel.User)?", "Confirm", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                $MbxUpdateButton.IsEnabled = $false
                
                if ($sel.HasFullAccess -and -not $newFullAccess) {
                    Write-Log "Removing Full Access for $($sel.User)"
                    Remove-MailboxPermission -Identity $mailbox -User $sel.UserIdentity -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
                } elseif (-not $sel.HasFullAccess -and $newFullAccess) {
                    Write-Log "Adding Full Access for $($sel.User)"
                    Add-MailboxPermission -Identity $mailbox -User $sel.UserIdentity -AccessRights FullAccess -InheritanceType All -AutoMapping $false -ErrorAction Stop
                }
                
                if ($sel.HasSendAs -and -not $newSendAs) {
                    Write-Log "Removing Send As for $($sel.User)"
                    Remove-RecipientPermission -Identity $mailbox -Trustee $sel.UserIdentity -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                } elseif (-not $sel.HasSendAs -and $newSendAs) {
                    Write-Log "Adding Send As for $($sel.User)"
                    Add-RecipientPermission -Identity $mailbox -Trustee $sel.UserIdentity -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                }
                
                Write-Log "Successfully updated permissions"
                [System.Windows.MessageBox]::Show("Permissions updated successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                $MbxLoadPermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                
            } catch {
                Write-Log "Error: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                $MbxUpdateButton.IsEnabled = $true
            }
        }
    })
	
	$MbxExportToExcelButton.Add_Click({
        $mailbox = $MbxViewMailboxBox.Text.Trim()
        $permissions = $MbxPermissionsGrid.ItemsSource
        
        if ($null -eq $permissions -or $permissions.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No permissions to export. Please load permissions first.", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed. Cannot export to Excel.`n`nPlease run: Install-Module -Name ImportExcel", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Mailbox Permissions Report"
            $saveDialog.FileName = "Mailbox_Permissions_$($mailbox.Replace('@','_'))_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $MbxExportToExcelButton.IsEnabled = $false
                Write-Log "Exporting permissions to Excel: $excelPath"
                
                $exportData = @()
                foreach ($perm in $permissions) {
                    $exportData += [PSCustomObject]@{
                        'Mailbox' = $mailbox
                        'Delegate User' = $perm.User
                        'Full Access' = if ($perm.HasFullAccess) { "Yes" } else { "No" }
                        'Send As' = if ($perm.HasSendAs) { "Yes" } else { "No" }
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "Mailbox Permissions" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "MailboxPermissions"
                
                Write-Log "Successfully exported $($exportData.Count) permissions to Excel"
                
                [System.Windows.MessageBox]::Show(
                    "Mailbox permissions exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
            } else {
                Write-Log "Export cancelled by user"
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error exporting to Excel:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $MbxExportToExcelButton.IsEnabled = $true
        }
    })
    
    $MbxViewCloseButton.Add_Click({ $MbxWindow.Close() })
    
    $MbxLoadRemovePermissionsButton.Add_Click({
        $mailbox = $MbxRemoveMailboxBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($mailbox)) { return }
        
        try {
            $MbxLoadRemovePermissionsButton.IsEnabled = $false
            Write-Log "Loading permissions for $mailbox"
            
            $perms = Get-CombinedMailboxPermissions -MailboxIdentity $mailbox
            $MbxRemovePermissionsGrid.ItemsSource = $perms
            
            Write-Log "Loaded $($perms.Count) user permissions"
        } catch {
            Write-Log "Error loading permissions: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } finally {
            $MbxLoadRemovePermissionsButton.IsEnabled = $true
        }
    })
    
    $MbxRemovePermissionsGrid.Add_SelectionChanged({
        if ($MbxRemovePermissionsGrid.SelectedItem) {
            $sel = $MbxRemovePermissionsGrid.SelectedItem
            $MbxRemoveFullAccessCheck.IsEnabled = $sel.HasFullAccess
            $MbxRemoveSendAsCheck.IsEnabled = $sel.HasSendAs
            $MbxRemoveFullAccessCheck.IsChecked = $false
            $MbxRemoveSendAsCheck.IsChecked = $false
            $MbxRemoveButton.IsEnabled = $true
        } else {
            $MbxRemoveFullAccessCheck.IsEnabled = $false
            $MbxRemoveSendAsCheck.IsEnabled = $false
            $MbxRemoveButton.IsEnabled = $false
        }
    })
    
    $MbxRemoveButton.Add_Click({
        $mailbox = $MbxRemoveMailboxBox.Text.Trim()
        $sel = $MbxRemovePermissionsGrid.SelectedItem
        $removeFullAccess = $MbxRemoveFullAccessCheck.IsChecked
        $removeSendAs = $MbxRemoveSendAsCheck.IsChecked
        
        if ($null -eq $sel) { return }
        
        if (-not $removeFullAccess -and -not $removeSendAs) {
            [System.Windows.MessageBox]::Show("Please select at least one permission type to remove", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $permsToRemove = @()
        if ($removeFullAccess) { $permsToRemove += "Full Access" }
        if ($removeSendAs) { $permsToRemove += "Send As" }
        
        $result = [System.Windows.MessageBox]::Show(
            "Remove the following permissions for $($sel.User)?`n`n$($permsToRemove -join ', ')",
            "Confirm",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                $MbxRemoveButton.IsEnabled = $false
                $removedPerms = @()
                
                if ($removeFullAccess) {
                    Write-Log "Removing Full Access for $($sel.User)"
                    Remove-MailboxPermission -Identity $mailbox -User $sel.UserIdentity -AccessRights FullAccess -Confirm:$false -ErrorAction Stop
                    $removedPerms += "Full Access"
                }
                
                if ($removeSendAs) {
                    Write-Log "Removing Send As for $($sel.User)"
                    Remove-RecipientPermission -Identity $mailbox -Trustee $sel.UserIdentity -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                    $removedPerms += "Send As"
                }
                
                Write-Log "Successfully removed: $($removedPerms -join ', ')"
                [System.Windows.MessageBox]::Show("Permissions removed successfully!`n`n$($removedPerms -join ', ')", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                $MbxLoadRemovePermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                
            } catch {
                Write-Log "Error: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                $MbxRemoveButton.IsEnabled = $true
            }
        }
    })
    
    $MbxRemoveCloseButton.Add_Click({ $MbxWindow.Close() })
	
	# Add Enter key support for Mailbox Add tab
	$MbxAddMailboxBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$MbxAddUserBox.Focus()
		}
	})

	$MbxAddUserBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$MbxAddButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
		}
	})

	# Add Enter key support for Mailbox View/Edit tab
	$MbxViewMailboxBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$MbxLoadPermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
		}
	})

	# Add Enter key support for Mailbox Remove tab
	$MbxRemoveMailboxBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$MbxLoadRemovePermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
		}
	})
    
    $MbxWindow.ShowDialog() | Out-Null
})

$syncHash.CalendarButton.Add_Click({
    # Check if connected
    $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
        $result = [System.Windows.MessageBox]::Show(
            "You are not connected to Exchange Online.`n`nWould you like to connect now?",
            "Connection Required",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $syncHash.ConnectButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            
            # Wait a moment for connection to establish, then recheck
            Start-Sleep -Seconds 2
            $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
            
            if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
                Write-Log "Connection was not established. Please try again."
                return
            }
        } else {
            return
        }
    }
    
    Write-Log "Opening Calendar Permissions window..."
    
    [xml]$CalendarXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Calendar Permissions Management" 
        Height="550" 
        Width="650" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <TabControl Margin="10">
            <TabItem Header="Add Permission">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="AddMailboxBox" Height="25" Padding="5"/>
                    </StackPanel>
                    
                    <StackPanel Grid.Row="1" Margin="0,0,0,15">
                        <TextBlock Text="User Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <TextBox x:Name="AddUserBox" Height="25" Padding="5"/>
                    </StackPanel>
                    
                    <StackPanel Grid.Row="2" Margin="0,0,0,15">
                        <TextBlock Text="Access Rights:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <ComboBox x:Name="AddPermissionCombo" Height="25">
                            <ComboBoxItem Content="AvailabilityOnly" Tag="AvailabilityOnly"/>
                            <ComboBoxItem Content="LimitedDetails" Tag="LimitedDetails"/>
                            <ComboBoxItem Content="Reviewer" Tag="Reviewer"/>
                            <ComboBoxItem Content="Contributor" Tag="Contributor"/>
                            <ComboBoxItem Content="Author" Tag="Author"/>
                            <ComboBoxItem Content="Editor" Tag="Editor"/>
                            <ComboBoxItem Content="Owner" Tag="Owner"/>
                        </ComboBox>
                    </StackPanel>
                    
                    <Border Grid.Row="3" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="1" Padding="10" Margin="0,0,0,15">
                        <TextBlock TextWrapping="Wrap" FontSize="11" Text="Select the appropriate permission level for calendar access."/>
                    </Border>
                    
                    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="AddButton" Content="Add Permission" Width="120" Height="30" Margin="0,0,10,0" Background="#28a745" Foreground="White" FontWeight="Bold"/>
                        <Button x:Name="AddCancelButton" Content="Cancel" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="View/Edit">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <DockPanel>
                            <Button x:Name="LoadPermissionsButton" Content="Load" Width="80" Height="25" Margin="10,0,0,0" DockPanel.Dock="Right" Background="#007bff" Foreground="White" FontWeight="Bold"/>
                            <TextBox x:Name="ViewMailboxBox" Height="25" Padding="5"/>
                        </DockPanel>
                    </StackPanel>
                    
                    <Border Grid.Row="1" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15">
                        <DataGrid x:Name="PermissionsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="User" Binding="{Binding User}" Width="*"/>
                                <DataGridTextColumn Header="Access Rights" Binding="{Binding AccessRights}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Border>
                    
                    <GroupBox Grid.Row="2" Header="Edit Selected" Padding="10" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock x:Name="EditUserLabel" Text="No permission selected" Margin="0,0,0,10"/>
                            <DockPanel>
                                <Button x:Name="UpdateButton" Content="Update" Width="80" Height="25" Margin="10,0,0,0" DockPanel.Dock="Right" Background="#ffc107" Foreground="Black" FontWeight="Bold" IsEnabled="False"/>
                                <ComboBox x:Name="EditPermissionCombo" Height="25" IsEnabled="False">
                                    <ComboBoxItem Content="AvailabilityOnly" Tag="AvailabilityOnly"/>
                                    <ComboBoxItem Content="LimitedDetails" Tag="LimitedDetails"/>
                                    <ComboBoxItem Content="Reviewer" Tag="Reviewer"/>
                                    <ComboBoxItem Content="Contributor" Tag="Contributor"/>
                                    <ComboBoxItem Content="Author" Tag="Author"/>
                                    <ComboBoxItem Content="Editor" Tag="Editor"/>
                                    <ComboBoxItem Content="Owner" Tag="Owner"/>
                                </ComboBox>
                            </DockPanel>
                        </StackPanel>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="ExportToExcelButton" Content="Export to Excel" Width="120" Height="30" Margin="0,0,10,0" Background="#17a2b8" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="ViewCloseButton" Content="Close" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Remove">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="Mailbox Email:" FontWeight="Bold" Margin="0,0,0,5"/>
                        <DockPanel>
                            <Button x:Name="LoadRemovePermissionsButton" Content="Load" Width="80" Height="25" Margin="10,0,0,0" DockPanel.Dock="Right" Background="#007bff" Foreground="White" FontWeight="Bold"/>
                            <TextBox x:Name="RemoveMailboxBox" Height="25" Padding="5"/>
                        </DockPanel>
                    </StackPanel>
                    
                    <Border Grid.Row="1" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15">
                        <DataGrid x:Name="RemovePermissionsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="User" Binding="{Binding User}" Width="*"/>
                                <DataGridTextColumn Header="Access Rights" Binding="{Binding AccessRights}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Border>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="RemoveButton" Content="Remove Selected" Width="130" Height="30" Margin="0,0,10,0" Background="#dc3545" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="RemoveCloseButton" Content="Close" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@
    
    $calReader = New-Object System.Xml.XmlNodeReader $CalendarXAML
    $CalWindow = [Windows.Markup.XamlReader]::Load($calReader)
    $CalWindow.Owner = $syncHash.Window
    
    $AddMailboxBox = $CalWindow.FindName("AddMailboxBox")
    $AddUserBox = $CalWindow.FindName("AddUserBox")
    $AddPermissionCombo = $CalWindow.FindName("AddPermissionCombo")
    $AddButton = $CalWindow.FindName("AddButton")
    $AddCancelButton = $CalWindow.FindName("AddCancelButton")
    
    $ViewMailboxBox = $CalWindow.FindName("ViewMailboxBox")
    $LoadPermissionsButton = $CalWindow.FindName("LoadPermissionsButton")
    $PermissionsGrid = $CalWindow.FindName("PermissionsGrid")
    $EditUserLabel = $CalWindow.FindName("EditUserLabel")
    $EditPermissionCombo = $CalWindow.FindName("EditPermissionCombo")
    $UpdateButton = $CalWindow.FindName("UpdateButton")
    $ExportToExcelButton = $CalWindow.FindName("ExportToExcelButton")
    $ViewCloseButton = $CalWindow.FindName("ViewCloseButton")
    
    $RemoveMailboxBox = $CalWindow.FindName("RemoveMailboxBox")
    $LoadRemovePermissionsButton = $CalWindow.FindName("LoadRemovePermissionsButton")
    $RemovePermissionsGrid = $CalWindow.FindName("RemovePermissionsGrid")
    $RemoveButton = $CalWindow.FindName("RemoveButton")
    $RemoveCloseButton = $CalWindow.FindName("RemoveCloseButton")
    
    $AddPermissionCombo.SelectedIndex = 2
    
    $PermissionsGrid.Add_MouseDoubleClick({
        if ($PermissionsGrid.SelectedItem) {
            $selectedUser = $PermissionsGrid.SelectedItem.User
            Show-ADPropertiesWindow -Identity $selectedUser -Owner $CalWindow
        }
    })
    
    $RemovePermissionsGrid.Add_MouseDoubleClick({
        if ($RemovePermissionsGrid.SelectedItem) {
            $selectedUser = $RemovePermissionsGrid.SelectedItem.User
            Show-ADPropertiesWindow -Identity $selectedUser -Owner $CalWindow
        }
    })
    
    $AddCancelButton.Add_Click({ $CalWindow.Close() })
    
    $AddButton.Add_Click({
        $mailbox = $AddMailboxBox.Text.Trim()
        $user = $AddUserBox.Text.Trim()
        $perm = $AddPermissionCombo.SelectedItem.Tag
        
        if ([string]::IsNullOrWhiteSpace($mailbox) -or [string]::IsNullOrWhiteSpace($user) -or $null -eq $perm) {
            [System.Windows.MessageBox]::Show("Please fill all fields", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $AddButton.IsEnabled = $false
            Write-Log "Adding $perm for $user on ${mailbox}:\Calendar"
            Add-MailboxFolderPermission -Identity "${mailbox}:\Calendar" -User $user -AccessRights $perm -ErrorAction Stop
            Write-Log "Successfully added permission"
            [System.Windows.MessageBox]::Show("Permission added successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            $AddMailboxBox.Clear(); $AddUserBox.Clear(); $AddPermissionCombo.SelectedIndex = 2
        } catch {
            Write-Log "Error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } finally {
            $AddButton.IsEnabled = $true
        }
    })
    
    $LoadPermissionsButton.Add_Click({
        $mailbox = $ViewMailboxBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($mailbox)) { return }
        
        try {
            $LoadPermissionsButton.IsEnabled = $false
            Write-Log "Loading permissions for $mailbox"
            
            # Get all permissions first
            $allPerms = Get-MailboxFolderPermission -Identity "${mailbox}:\Calendar" -ErrorAction Stop
            Write-Log "DEBUG: Retrieved $($allPerms.Count) total permissions"
            
            # Filter and format properly
            $perms = @()
            foreach ($perm in $allPerms) {
                $userName = $perm.User.DisplayName
                
                # Skip Default and Anonymous
                if ($userName -eq "Default" -or $userName -eq "Anonymous") {
                    Write-Log "DEBUG: Skipping $userName"
                    continue
                }
                
                Write-Log "DEBUG: Adding user: $userName with rights: $($perm.AccessRights)"
                
                $perms += [PSCustomObject]@{
                    User = $userName
                    AccessRights = ($perm.AccessRights -join ", ")
                }
            }
            
            Write-Log "Found $($perms.Count) delegated permissions (after filtering)"
            
            # Bind to grid
            if ($perms.Count -gt 0) {
                $PermissionsGrid.ItemsSource = $perms
                $ExportToExcelButton.IsEnabled = $true
            } else {
                $PermissionsGrid.ItemsSource = $null
                $ExportToExcelButton.IsEnabled = $false
                [System.Windows.MessageBox]::Show("No delegated calendar permissions found (only Default/Anonymous).", "No Permissions", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
            
        } catch {
            Write-Log "Error loading permissions: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } finally {
            $LoadPermissionsButton.IsEnabled = $true
        }
    })
    
    $PermissionsGrid.Add_SelectionChanged({
        if ($PermissionsGrid.SelectedItem) {
            $sel = $PermissionsGrid.SelectedItem
            $EditUserLabel.Text = "Editing: $($sel.User)"
            $EditPermissionCombo.IsEnabled = $true
            $UpdateButton.IsEnabled = $true
            $currentRight = ($sel.AccessRights -split ",")[0].Trim()
            for ($i = 0; $i -lt $EditPermissionCombo.Items.Count; $i++) {
                if ($EditPermissionCombo.Items[$i].Tag -eq $currentRight) {
                    $EditPermissionCombo.SelectedIndex = $i
                    break
                }
            }
        }
    })
    
    $UpdateButton.Add_Click({
        $mailbox = $ViewMailboxBox.Text.Trim()
        $sel = $PermissionsGrid.SelectedItem
        $newPerm = $EditPermissionCombo.SelectedItem.Tag
        
        if ($null -eq $sel -or $null -eq $newPerm) { return }
        
        $result = [System.Windows.MessageBox]::Show("Update permission for $($sel.User) to $newPerm?", "Confirm", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                $UpdateButton.IsEnabled = $false
                Write-Log "Updating permission for $($sel.User) to $newPerm"
                Set-MailboxFolderPermission -Identity "${mailbox}:\Calendar" -User $sel.User -AccessRights $newPerm -ErrorAction Stop
                Write-Log "Successfully updated"
                [System.Windows.MessageBox]::Show("Permission updated!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                $LoadPermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            } catch {
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                $UpdateButton.IsEnabled = $true
            }
        }
    })
	
	$ExportToExcelButton.Add_Click({
        $mailbox = $ViewMailboxBox.Text.Trim()
        $permissions = $PermissionsGrid.ItemsSource
        
        if ($null -eq $permissions -or $permissions.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No permissions to export. Please load permissions first.", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed. Cannot export to Excel.`n`nPlease run: Install-Module -Name ImportExcel", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Calendar Permissions Report"
            $saveDialog.FileName = "Calendar_Permissions_$($mailbox.Replace('@','_'))_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $ExportToExcelButton.IsEnabled = $false
                Write-Log "Exporting permissions to Excel: $excelPath"
                
                $exportData = @()
                foreach ($perm in $permissions) {
                    $exportData += [PSCustomObject]@{
                        'Mailbox' = $mailbox
                        'Delegate User' = $perm.User
                        'Access Rights' = $perm.AccessRights
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "Calendar Permissions" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "CalendarPermissions"
                
                Write-Log "Successfully exported $($exportData.Count) permissions to Excel"
                
                [System.Windows.MessageBox]::Show(
                    "Calendar permissions exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
                
            } else {
                Write-Log "Export cancelled by user"
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error exporting to Excel:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $ExportToExcelButton.IsEnabled = $true
        }
    })
    
    $ViewCloseButton.Add_Click({ $CalWindow.Close() })
    
    $LoadRemovePermissionsButton.Add_Click({
        $mailbox = $RemoveMailboxBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($mailbox)) { return }
        
        try {
            $LoadRemovePermissionsButton.IsEnabled = $false
            Write-Log "Loading permissions for $mailbox"
            
            # Get all permissions first
            $allPerms = Get-MailboxFolderPermission -Identity "${mailbox}:\Calendar" -ErrorAction Stop
            Write-Log "DEBUG: Retrieved $($allPerms.Count) total permissions"
            
            # Filter and format properly
            $perms = @()
            foreach ($perm in $allPerms) {
                $userName = $perm.User.DisplayName
                
                # Skip Default and Anonymous
                if ($userName -eq "Default" -or $userName -eq "Anonymous") {
                    Write-Log "DEBUG: Skipping $userName"
                    continue
                }
                
                Write-Log "DEBUG: Adding user: $userName with rights: $($perm.AccessRights)"
                
                $perms += [PSCustomObject]@{
                    User = $userName
                    AccessRights = ($perm.AccessRights -join ", ")
                }
            }
            
            Write-Log "Found $($perms.Count) delegated permissions (after filtering)"
            
            # Bind to grid
            if ($perms.Count -gt 0) {
                $RemovePermissionsGrid.ItemsSource = $perms
            } else {
                $RemovePermissionsGrid.ItemsSource = $null
                [System.Windows.MessageBox]::Show("No delegated calendar permissions found (only Default/Anonymous).", "No Permissions", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
            
        } catch {
            Write-Log "Error loading permissions: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } finally {
            $LoadRemovePermissionsButton.IsEnabled = $true
        }
    })

    $RemovePermissionsGrid.Add_SelectionChanged({
        $RemoveButton.IsEnabled = ($null -ne $RemovePermissionsGrid.SelectedItem)
    })
    
    $RemoveButton.Add_Click({
        $mailbox = $RemoveMailboxBox.Text.Trim()
        $sel = $RemovePermissionsGrid.SelectedItem
        if ($null -eq $sel) { return }
        
        $result = [System.Windows.MessageBox]::Show("Remove permission for $($sel.User)?", "Confirm", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                $RemoveButton.IsEnabled = $false
                Write-Log "Removing permission for $($sel.User)"
                Remove-MailboxFolderPermission -Identity "${mailbox}:\Calendar" -User $sel.User -Confirm:$false -ErrorAction Stop
                Write-Log "Successfully removed"
                [System.Windows.MessageBox]::Show("Permission removed!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                $LoadRemovePermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            } catch {
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                $RemoveButton.IsEnabled = $true
            }
        }
    })
    
    $RemoveCloseButton.Add_Click({ $CalWindow.Close() })
	
	# Add Enter key support for Calendar Add tab
	$AddMailboxBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$AddUserBox.Focus()
		}
	})

	$AddUserBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$AddButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
		}
	})

	# Add Enter key support for Calendar View/Edit tab
	$ViewMailboxBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$LoadPermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
		}
	})

	# Add Enter key support for Calendar Remove tab
	$RemoveMailboxBox.Add_KeyDown({
		param($sender, $e)
		if ($e.Key -eq [System.Windows.Input.Key]::Return) {
			$LoadRemovePermissionsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
		}
	})
    
    $CalWindow.ShowDialog() | Out-Null
})

$syncHash.GroupMembersButton.Add_Click({
    Write-Log "Opening AD Group Members window..."
    
    # Check if Active Directory module is available
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.MessageBox]::Show(
            "Active Directory PowerShell module is not installed.`n`nThis feature requires the RSAT Active Directory module.`n`nPlease install it from:`nSettings > Apps > Optional Features > Add RSAT: Active Directory Domain Services",
            "Module Required",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        Write-Log "AD Group Members requires ActiveDirectory module"
        return
    }
    
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        [System.Windows.MessageBox]::Show(
            "Failed to load Active Directory module:`n`n$($_.Exception.Message)",
            "Module Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }
    
    [xml]$GroupMembersXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AD Group Members Management" 
        Height="550" 
        Width="700" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <GroupBox Grid.Row="0" Header="Group Information" Margin="15,15,15,10" Padding="15">
            <StackPanel>
                <TextBlock Text="Group Name, Email, or SAM Account Name:" FontWeight="Bold" Margin="0,0,0,5"/>
                <DockPanel Margin="0,0,0,10">
                    <Button x:Name="LoadGroupButton" 
                           Content="Load Members" 
                           Width="120" 
                           Height="30" 
                           Margin="10,0,0,0" 
                           DockPanel.Dock="Right" 
                           Background="#007bff" 
                           Foreground="White" 
                           FontWeight="Bold"/>
                    <TextBox x:Name="GroupIdentityBox" 
                            Height="30" 
                            Padding="5"
                            VerticalContentAlignment="Center"/>
                </DockPanel>
                
                <StackPanel x:Name="GroupInfoPanel" Visibility="Collapsed" Margin="0,10,0,0">
                    <Border Background="#e7f3ff" BorderBrush="#007bff" BorderThickness="1" Padding="10" CornerRadius="3">
                        <StackPanel>
                            <TextBlock x:Name="GroupNameText" FontWeight="Bold" Margin="0,0,0,5"/>
                            <TextBlock x:Name="GroupTypeText" FontSize="11" Foreground="#666" Margin="0,0,0,3"/>
                            <TextBlock x:Name="GroupEmailText" FontSize="11" Foreground="#666" Margin="0,0,0,3"/>
                            <TextBlock x:Name="GroupScopeText" FontSize="11" Foreground="#666" Margin="0,0,0,3"/>
                            <TextBlock x:Name="GroupMemberCountText" FontSize="11" Foreground="#666"/>
                        </StackPanel>
                    </Border>
                </StackPanel>
            </StackPanel>
        </GroupBox>
        
        <GroupBox Grid.Row="1" Header="Group Members" Margin="15,0,15,10" Padding="10">
            <Border BorderBrush="#dee2e6" BorderThickness="1">
                <DataGrid x:Name="MembersGrid" 
                         AutoGenerateColumns="False" 
                         IsReadOnly="True" 
                         SelectionMode="Extended"
                         AlternatingRowBackground="#f8f9fa">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="2*"/>
                        <DataGridTextColumn Header="Email Address" Binding="{Binding Email}" Width="2*"/>
                        <DataGridTextColumn Header="Object Type" Binding="{Binding ObjectClass}" Width="*"/>
                        <DataGridTextColumn Header="Title" Binding="{Binding Title}" Width="1.5*"/>
                        <DataGridTextColumn Header="Department" Binding="{Binding Department}" Width="1.5*"/>
                    </DataGrid.Columns>
                </DataGrid>
            </Border>
        </GroupBox>
        
        <Border Grid.Row="2" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="0,1,0,0" Padding="15">
            <DockPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" DockPanel.Dock="Left">
                    <TextBlock x:Name="StatusText" 
                              Text="Enter a group name or email to begin" 
                              VerticalAlignment="Center"
                              FontSize="11"
                              Foreground="#666"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="CopyEmailsButton" 
                           Content="Copy Emails" 
                           Width="110" 
                           Height="32" 
                           Margin="0,0,10,0" 
                           Background="#6c757d" 
                           Foreground="White" 
                           FontWeight="Bold"
                           IsEnabled="False"/>
                    <Button x:Name="ExportToExcelButton" 
                           Content="Export to Excel" 
                           Width="130" 
                           Height="32" 
                           Margin="0,0,10,0" 
                           Background="#28a745" 
                           Foreground="White" 
                           FontWeight="Bold"
                           IsEnabled="False"/>
                    <Button x:Name="CloseButton" 
                           Content="Close" 
                           Width="80" 
                           Height="32" 
                           Background="#6c757d" 
                           Foreground="White"/>
                </StackPanel>
            </DockPanel>
        </Border>
    </Grid>
</Window>
"@
    
    $grpReader = New-Object System.Xml.XmlNodeReader $GroupMembersXAML
    $GrpWindow = [Windows.Markup.XamlReader]::Load($grpReader)
    $GrpWindow.Owner = $syncHash.Window
    
    $GroupIdentityBox = $GrpWindow.FindName("GroupIdentityBox")
    $LoadGroupButton = $GrpWindow.FindName("LoadGroupButton")
    $GroupInfoPanel = $GrpWindow.FindName("GroupInfoPanel")
    $GroupNameText = $GrpWindow.FindName("GroupNameText")
    $GroupTypeText = $GrpWindow.FindName("GroupTypeText")
    $GroupEmailText = $GrpWindow.FindName("GroupEmailText")
    $GroupScopeText = $GrpWindow.FindName("GroupScopeText")
    $GroupMemberCountText = $GrpWindow.FindName("GroupMemberCountText")
    $MembersGrid = $GrpWindow.FindName("MembersGrid")
    $StatusText = $GrpWindow.FindName("StatusText")
    $CopyEmailsButton = $GrpWindow.FindName("CopyEmailsButton")
    $ExportToExcelButton = $GrpWindow.FindName("ExportToExcelButton")
    $CloseButton = $GrpWindow.FindName("CloseButton")
    
    $script:currentGroupInfo = $null
    $script:currentMembers = $null
    
    $MembersGrid.Add_MouseDoubleClick({
        if ($MembersGrid.SelectedItem) {
            $selectedMember = $MembersGrid.SelectedItem.Identity
            Show-ADPropertiesWindow -Identity $selectedMember -Owner $GrpWindow
        }
    })
    
    $LoadGroupButton.Add_Click({
        $groupIdentity = $GroupIdentityBox.Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($groupIdentity)) {
            [System.Windows.MessageBox]::Show("Please enter a group name, email, or SAM account name", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $LoadGroupButton.IsEnabled = $false
            $StatusText.Text = "Loading group information..."
            Write-Log "Loading AD group: $groupIdentity"
            
            # Try to find the group using various methods
            $group = $null
            
            # Try by Identity (works for DN, GUID, SAM, etc.)
            try {
                $group = Get-ADGroup -Identity $groupIdentity -Properties * -ErrorAction Stop
                Write-Log "Found group by identity: $($group.Name)"
            } catch {
                # Try by email address
                try {
                    $group = Get-ADGroup -Filter "mail -eq '$groupIdentity'" -Properties * -ErrorAction Stop
                    Write-Log "Found group by email: $($group.Name)"
                } catch {
                    # Try by display name
                    try {
                        $group = Get-ADGroup -Filter "DisplayName -eq '$groupIdentity'" -Properties * -ErrorAction Stop
                        Write-Log "Found group by display name: $($group.Name)"
                    } catch {
                        throw "Group not found. Please verify the group name, email, or SAM account name."
                    }
                }
            }
            
            if ($null -eq $group) {
                throw "Group not found"
            }
            
            $StatusText.Text = "Loading members..."
            Write-Log "Retrieving members for: $($group.Name)"
            
            # Get group members using AD cmdlets
            $members = Get-ADGroupMember -Identity $group.DistinguishedName -ErrorAction Stop
            
            $enrichedMembers = @()
            $processedCount = 0
            $totalCount = $members.Count
            
            foreach ($member in $members) {
                $processedCount++
                $StatusText.Text = "Processing member $processedCount of $totalCount..."
                
                try {
                    $displayName = ""
                    $email = ""
                    $title = ""
                    $department = ""
                    $objectClass = $member.objectClass
                    
                    # Get detailed information based on object type
                    if ($member.objectClass -eq "user") {
                        try {
                            $adUser = Get-ADUser -Identity $member.DistinguishedName -Properties DisplayName, EmailAddress, Title, Department, mail -ErrorAction Stop
                            $displayName = if ($adUser.DisplayName) { $adUser.DisplayName } else { $adUser.Name }
                            $email = if ($adUser.EmailAddress) { $adUser.EmailAddress } elseif ($adUser.mail) { $adUser.mail } else { "" }
                            $title = if ($adUser.Title) { $adUser.Title } else { "" }
                            $department = if ($adUser.Department) { $adUser.Department } else { "" }
                        } catch {
                            Write-Log "Warning: Could not get full details for user $($member.Name)"
                            $displayName = $member.Name
                        }
                    } elseif ($member.objectClass -eq "group") {
                        try {
                            $adGroup = Get-ADGroup -Identity $member.DistinguishedName -Properties DisplayName, mail -ErrorAction Stop
                            $displayName = if ($adGroup.DisplayName) { $adGroup.DisplayName } else { $adGroup.Name }
                            $email = if ($adGroup.mail) { $adGroup.mail } else { "" }
                            $objectClass = "Group"
                        } catch {
                            Write-Log "Warning: Could not get full details for group $($member.Name)"
                            $displayName = $member.Name
                        }
                    } elseif ($member.objectClass -eq "computer") {
                        $displayName = $member.Name
                        $objectClass = "Computer"
                    } elseif ($member.objectClass -eq "contact") {
                        try {
                            $adContact = Get-ADObject -Identity $member.DistinguishedName -Properties DisplayName, mail -ErrorAction Stop
                            $displayName = if ($adContact.DisplayName) { $adContact.DisplayName } else { $member.Name }
                            $email = if ($adContact.mail) { $adContact.mail } else { "" }
                            $objectClass = "Contact"
                        } catch {
                            $displayName = $member.Name
                            $objectClass = "Contact"
                        }
                    } else {
                        $displayName = $member.Name
                    }
                    
                    $memberObj = [PSCustomObject]@{
                        DisplayName = $displayName
                        Email = $email
                        ObjectClass = $objectClass
                        Title = $title
                        Department = $department
                        Identity = $member.DistinguishedName
                        SAMAccountName = $member.SamAccountName
                    }
                    
                    $enrichedMembers += $memberObj
                } catch {
                    Write-Log "Warning: Error processing member $($member.Name): $($_.Exception.Message)"
                    $enrichedMembers += [PSCustomObject]@{
                        DisplayName = $member.Name
                        Email = ""
                        ObjectClass = $member.objectClass
                        Title = ""
                        Department = ""
                        Identity = $member.DistinguishedName
                        SAMAccountName = $member.SamAccountName
                    }
                }
            }
            
            # Sort by display name
            $enrichedMembers = $enrichedMembers | Sort-Object DisplayName
            
            $MembersGrid.ItemsSource = $enrichedMembers
            
            # Determine group type
            $groupType = "Unknown"
            if ($group.GroupCategory -eq "Security") {
                $groupType = "Security Group"
            } elseif ($group.GroupCategory -eq "Distribution") {
                $groupType = "Distribution Group"
            }
            
            # Update info panel
            $GroupNameText.Text = "Group: $($group.Name)"
            $GroupTypeText.Text = "Category: $groupType"
            $GroupScopeText.Text = "Scope: $($group.GroupScope)"
            $GroupEmailText.Text = "Email: $(if ($group.mail) { $group.mail } else { 'N/A' })"
            $GroupMemberCountText.Text = "Total Members: $($enrichedMembers.Count)"
            $GroupInfoPanel.Visibility = [System.Windows.Visibility]::Visible
            
            $script:currentGroupInfo = $group
            $script:currentMembers = $enrichedMembers
            
            $ExportToExcelButton.IsEnabled = $true
            $CopyEmailsButton.IsEnabled = $true
            
            $StatusText.Text = "Loaded $($enrichedMembers.Count) members successfully"
            Write-Log "Successfully loaded $($enrichedMembers.Count) members from $($group.Name)"
            
        } catch {
            Write-Log "Error loading group: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error loading group:`n`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            $StatusText.Text = "Error loading group"
            $GroupInfoPanel.Visibility = [System.Windows.Visibility]::Collapsed
            $MembersGrid.ItemsSource = $null
            $ExportToExcelButton.IsEnabled = $false
            $CopyEmailsButton.IsEnabled = $false
        } finally {
            $LoadGroupButton.IsEnabled = $true
        }
    })
    
    $CopyEmailsButton.Add_Click({
        if ($null -eq $script:currentMembers -or $script:currentMembers.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No members to copy", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $emailAddresses = $script:currentMembers | 
                Where-Object { -not [string]::IsNullOrWhiteSpace($_.Email) } | 
                Select-Object -ExpandProperty Email
            
            if ($emailAddresses.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No email addresses found to copy", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            $emailList = $emailAddresses -join "; "
            [System.Windows.Forms.Clipboard]::SetText($emailList)
            
            Write-Log "Copied $($emailAddresses.Count) email addresses to clipboard"
            [System.Windows.MessageBox]::Show("Copied $($emailAddresses.Count) email addresses to clipboard!`n`nYou can now paste them into an email.", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            
        } catch {
            Write-Log "Error copying emails: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error copying emails:`n`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })
    
    $ExportToExcelButton.Add_Click({
        if ($null -eq $script:currentMembers -or $script:currentMembers.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No members to export", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed. Cannot export to Excel.`n`nPlease run: Install-Module -Name ImportExcel", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Group Members Report"
            $groupNameSafe = $script:currentGroupInfo.Name -replace '[\\/:*?"<>|]', '_'
            $saveDialog.FileName = "AD_Group_Members_${groupNameSafe}_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $ExportToExcelButton.IsEnabled = $false
                $StatusText.Text = "Exporting to Excel..."
                Write-Log "Exporting group members to Excel: $excelPath"
                
                $exportData = @()
                foreach ($member in $script:currentMembers) {
                    $exportData += [PSCustomObject]@{
                        'Group Name' = $script:currentGroupInfo.Name
                        'Group Email' = if ($script:currentGroupInfo.mail) { $script:currentGroupInfo.mail } else { "N/A" }
                        'Group Category' = $script:currentGroupInfo.GroupCategory
                        'Group Scope' = $script:currentGroupInfo.GroupScope
                        'Member Display Name' = $member.DisplayName
                        'Member Email' = $member.Email
                        'Member SAM Account' = $member.SAMAccountName
                        'Object Type' = $member.ObjectClass
                        'Title' = $member.Title
                        'Department' = $member.Department
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "AD Group Members" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "ADGroupMembers"
                
                Write-Log "Successfully exported $($exportData.Count) members to Excel"
                $StatusText.Text = "Export completed successfully"
                
                [System.Windows.MessageBox]::Show(
                    "Group members exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
                
            } else {
                Write-Log "Export cancelled by user"
                $StatusText.Text = "Export cancelled"
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error exporting to Excel:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            $StatusText.Text = "Export failed"
        } finally {
            $ExportToExcelButton.IsEnabled = $true
        }
    })
    
    $CloseButton.Add_Click({ $GrpWindow.Close() })
    
    $GroupIdentityBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $LoadGroupButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $GrpWindow.ShowDialog() | Out-Null
})

$syncHash.ExportActiveUsersButton.Add_Click({
    Write-Log "Opening Export Active Users Report"
    
    # Check for Active Directory module
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.MessageBox]::Show(
            "Active Directory module is not available.`n`nThis feature requires the Active Directory PowerShell module (RSAT).",
            "Module Required",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        Write-Log "Export Active Users failed: ActiveDirectory module not available"
        return
    }
    
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        [System.Windows.MessageBox]::Show(
            "Failed to load Active Directory module.`n`n$($_.Exception.Message)",
            "Module Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        Write-Log "Export Active Users failed: Could not import ActiveDirectory module"
        return
    }
    
    # Check for ImportExcel module
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        [System.Windows.MessageBox]::Show(
            "ImportExcel module is not installed.`n`nThis feature requires the ImportExcel module for Excel export.",
            "Module Required",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        Write-Log "Export Active Users failed: ImportExcel module not installed"
        return
    }
    
    Write-Log "Starting Active Users export..."
    
    # Show progress window
    $progressResult = [System.Windows.MessageBox]::Show(
        "This will retrieve all enabled user accounts from Active and Consultants OUs.`n`nThis may take a few moments. Continue?",
        "Export Active Users",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($progressResult -eq [System.Windows.MessageBoxResult]::No) {
        Write-Log "Export Active Users cancelled by user"
        return
    }
    
    try {
        # Define target OUs - modify these to match your organization's OU structure
        $OUList = @(
            "OU=Active,OU=Accts,DC=gellerco,DC=net", 
            "OU=Consultants,OU=Accts,DC=gellerco,DC=net",
            "OU=Vendor,OU=Accts,DC=gellerco,DC=net",
            "OU=Interns,OU=Accts,DC=gellerco,DC=net"
        )
        
        Write-Log "Retrieving users from OUs..."
        
        # Collect enabled users from all OUs
        $activeUsers = @()
        foreach ($OU in $OUList) {
            try {
                $users = Get-ADUser -Filter {Enabled -eq $True} -SearchBase $OU -Properties Name, SamAccountName, Enabled, EmailAddress, Department, Title, Office -ErrorAction Stop
                if ($users) {
                    $activeUsers += $users
                }
                Write-Log "Retrieved $($users.Count) users from $OU"
            } catch {
                Write-Log "Warning: Could not retrieve users from $OU - $($_.Exception.Message)"
            }
        }
        
        if ($activeUsers.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "No enabled users found in the specified OUs.",
                "No Data",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            Write-Log "No active users found"
            return
        }
        
        Write-Log "Found $($activeUsers.Count) total enabled users"
        
        # Exclude users with "test" or "t-" in SamAccountName
        $filteredUsers = $activeUsers | Where-Object {
            $_.SamAccountName -notmatch '(?i)test|^t-'
        }
        
        Write-Log "Filtered to $($filteredUsers.Count) users (excluding test accounts)"
        
        # Select properties to export and sort by Name in ascending order
        $selectedProperties = $filteredUsers | 
            Select-Object Name, SamAccountName, Enabled, EmailAddress, Department, Title, Office | 
            Sort-Object -Property Name
        
        # Prompt user for save location
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $saveDialog.Title = "Save Active Users Report"
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $saveDialog.FileName = "AD-ActiveAccounts_$timestamp.xlsx"
        
        if ($saveDialog.ShowDialog()) {
            $excelPath = $saveDialog.FileName
            
            Write-Log "Exporting $($selectedProperties.Count) users to Excel: $excelPath"
            
            # Import ImportExcel module
            Import-Module ImportExcel -ErrorAction Stop
            
            # Export to Excel
            $selectedProperties | Export-Excel -Path $excelPath `
                -AutoSize `
                -AutoFilter `
                -FreezeTopRow `
                -BoldTopRow `
                -TableStyle Medium1 `
                -WorksheetName "ActiveAccounts"
            
            Write-Log "Successfully exported active users report"
            
            [System.Windows.MessageBox]::Show(
                "Active users report exported successfully!`n`nFile: $excelPath`n`nExported $($selectedProperties.Count) users.",
                "Export Successful",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
            
            # Open folder location in Explorer
            $folderPath = Split-Path $excelPath -Parent
            Invoke-Item $folderPath
            
        } else {
            Write-Log "Export cancelled by user"
        }
        
    } catch {
        Write-Log "Export error: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show(
            "Error exporting active users:`n`n$($_.Exception.Message)",
            "Export Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
})

# Employee Conversion Button Handler
$syncHash.EmployeeConversionButton.Add_Click({
    Write-Log "Opening Employee Conversion window..."
    Show-EmployeeConversionWindow
})

$syncHash.LockedOutUsersButton.Add_Click({
    Write-Log "Opening Locked Out Users window..."
    Show-LockedOutUsersWindow
})

# Function to show Employee Conversion window
function Show-EmployeeConversionWindow {
    # Check for Active Directory module
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.MessageBox]::Show(
            "Active Directory module is not available.`n`nThis feature requires the Active Directory PowerShell module (RSAT).",
            "Module Required",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }
    
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        [System.Windows.MessageBox]::Show(
            "Failed to load Active Directory module.`n`n$($_.Exception.Message)",
            "Module Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }
    
    [xml]$ConversionXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Employee Conversion" 
        Height="550" 
        Width="700" 
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#233A4A" Padding="20,15">
            <StackPanel>
                <TextBlock Text="Employee Conversion" FontSize="20" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="Convert employee status (e.g., Consultant to Full Time Employee)" FontSize="12" Foreground="#B0BEC5" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Main Content -->
        <Grid Grid.Row="1" Margin="30,20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            
            <!-- From Template -->
            <StackPanel Grid.Row="0" Margin="0,0,0,20">
                <TextBlock Text="From Template (Current Status):" FontWeight="Bold" Margin="0,0,0,8" FontSize="14"/>
                <ComboBox x:Name="FromTemplateCombo" Height="35" FontSize="13" Padding="10,8"/>
            </StackPanel>
            
            <!-- To Template -->
            <StackPanel Grid.Row="1" Margin="0,0,0,20">
                <TextBlock Text="To Template (New Status):" FontWeight="Bold" Margin="0,0,0,8" FontSize="14"/>
                <ComboBox x:Name="ToTemplateCombo" Height="35" FontSize="13" Padding="10,8"/>
            </StackPanel>
            
            <!-- Username -->
            <StackPanel Grid.Row="2" Margin="0,0,0,25">
                <TextBlock Text="Username (SamAccountName):" FontWeight="Bold" Margin="0,0,0,8" FontSize="14"/>
                <TextBox x:Name="UsernameBox" Height="35" FontSize="13" Padding="10,8" VerticalContentAlignment="Center"/>
            </StackPanel>
            
            <!-- Status Message -->
            <Border Grid.Row="3" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="1" Padding="15" Margin="0,0,0,20" CornerRadius="5">
                <TextBlock x:Name="StatusText" Text="Ready to convert employee" FontSize="12" TextWrapping="Wrap" Foreground="#6c757d"/>
            </Border>
            
            <!-- Preview Area -->
            <Border Grid.Row="4" Background="White" BorderBrush="#dee2e6" BorderThickness="1" CornerRadius="5" Padding="15">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <TextBlock x:Name="PreviewText" FontFamily="Consolas" FontSize="11" TextWrapping="Wrap" Foreground="#495057"/>
                </ScrollViewer>
            </Border>
        </Grid>
        
        <!-- Footer Buttons -->
        <Border Grid.Row="2" Background="#f8f9fa" BorderThickness="0,1,0,0" BorderBrush="#dee2e6" Padding="20,15">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="ConvertButton" 
                        Content="Convert User" 
                        Width="120" 
                        Height="35" 
                        Background="#28a745" 
                        Foreground="White" 
                        FontWeight="Bold"
                        Margin="0,0,10,0"/>
                <Button x:Name="CloseButton" 
                        Content="Close" 
                        Width="100" 
                        Height="35" 
                        Background="#6c757d" 
                        Foreground="White"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    
    try {
        $reader = New-Object System.Xml.XmlNodeReader $ConversionXAML
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        # Get controls
        $FromTemplateCombo = $window.FindName("FromTemplateCombo")
        $ToTemplateCombo = $window.FindName("ToTemplateCombo")
        $UsernameBox = $window.FindName("UsernameBox")
        $StatusText = $window.FindName("StatusText")
        $PreviewText = $window.FindName("PreviewText")
        $ConvertButton = $window.FindName("ConvertButton")
        $CloseButton = $window.FindName("CloseButton")
        
        # Get all template accounts (accounts starting with _Template)
        try {
            $templates = Get-ADUser -Filter 'SamAccountName -like "_Template*"' -Properties DistinguishedName, SamAccountName | 
                Sort-Object SamAccountName
            
            if ($templates.Count -eq 0) {
                [System.Windows.MessageBox]::Show(
                    "No template accounts found.`n`nTemplate accounts must start with '_Template'.",
                    "No Templates Found",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                # Template accounts check
                $window.Close()
                return
            }
            
            # Templates loaded
            
            # Populate combo boxes
            foreach ($template in $templates) {
                $FromTemplateCombo.Items.Add($template.SamAccountName) | Out-Null
                $ToTemplateCombo.Items.Add($template.SamAccountName) | Out-Null
            }
            
            $StatusText.Text = "Loaded $($templates.Count) template account(s). Select templates and enter username to begin."
            
        } catch {
            [System.Windows.MessageBox]::Show(
                "Error loading template accounts:`n`n$($_.Exception.Message)",
                "Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            # Error logged
            $window.Close()
            return
        }
        
        # Convert Button Click
        $ConvertButton.Add_Click({
            $fromTemplate = $FromTemplateCombo.SelectedItem
            $toTemplate = $ToTemplateCombo.SelectedItem
            $username = $UsernameBox.Text.Trim()
            
            # Validation
            if ([string]::IsNullOrWhiteSpace($fromTemplate)) {
                [System.Windows.MessageBox]::Show("Please select a 'From Template'.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            if ([string]::IsNullOrWhiteSpace($toTemplate)) {
                [System.Windows.MessageBox]::Show("Please select a 'To Template'.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            if ($fromTemplate -eq $toTemplate) {
                [System.Windows.MessageBox]::Show("'From Template' and 'To Template' cannot be the same.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            if ([string]::IsNullOrWhiteSpace($username)) {
                [System.Windows.MessageBox]::Show("Please enter a username.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            try {
                $ConvertButton.IsEnabled = $false
                $StatusText.Text = "Processing conversion..."
                $StatusText.Foreground = "#ffc107"
                $PreviewText.Text = ""
                
                
                # Step 1: Verify user exists
                try {
                    $user = Get-ADUser -Identity $username -Properties MemberOf, DistinguishedName -ErrorAction Stop
                } catch {
                    throw "User '$username' not found in Active Directory."
                }
                
                # Step 2: Get groups from FROM template (groups to remove)
                try {
                    $fromTemplateObj = Get-ADUser -Identity $fromTemplate -Properties MemberOf -ErrorAction Stop
                    $fromTemplateGroups = @()
                    if ($fromTemplateObj.MemberOf) {
                        $fromTemplateGroups = $fromTemplateObj.MemberOf
                    }
                } catch {
                    throw "Could not retrieve groups for template '$fromTemplate'."
                }
                
                # Step 3: Get groups from TO template (groups to add)
                try {
                    $toTemplateObj = Get-ADUser -Identity $toTemplate -Properties MemberOf, DistinguishedName -ErrorAction Stop
                    $toTemplateGroups = @()
                    if ($toTemplateObj.MemberOf) {
                        $toTemplateGroups = $toTemplateObj.MemberOf
                    }
                    $toTemplateOU = ($toTemplateObj.DistinguishedName -split ',', 2)[1]
                } catch {
                    throw "Could not retrieve groups for template '$toTemplate'."
                }
                
                # Step 4: Build preview
                $preview = ""
                $preview += "USER: $($user.Name) ($($user.SamAccountName))`n"
                $preview += "=" * 60 + "`n`n"
                
                # Step 5: Remove groups (ONLY those in fromTemplate, excluding Domain Users)
                $groupsToRemove = @()
                $groupsRemaining = @()
                
                foreach ($userGroup in $user.MemberOf) {
                    if ($fromTemplateGroups -contains $userGroup) {
                        # Get group name for display
                        try {
                            $groupObj = Get-ADGroup -Identity $userGroup -ErrorAction Stop
                            $groupName = $groupObj.Name
                            
                            # Don't remove Domain Users
                            if ($groupName -ne "Domain Users") {
                                $groupsToRemove += @{DN = $userGroup; Name = $groupName}
                            }
                        } catch {
                        }
                    } else {
                        # This group is NOT in the from template, so it stays
                        try {
                            $groupObj = Get-ADGroup -Identity $userGroup -ErrorAction Stop
                            $groupsRemaining += $groupObj.Name
                        } catch {
                        }
                    }
                }
                
                $preview += "GROUPS TO REMOVE ($($groupsToRemove.Count)):`n"
                if ($groupsToRemove.Count -eq 0) {
                    $preview += "  (none)`n"
                } else {
                    foreach ($grp in $groupsToRemove) {
                        $preview += "  - $($grp.Name)`n"
                    }
                }
                $preview += "`n"
                
                $preview += "GROUPS TO ADD ($($toTemplateGroups.Count)):`n"
                if ($toTemplateGroups.Count -eq 0) {
                    $preview += "  (none)`n"
                } else {
                    foreach ($grpDN in $toTemplateGroups) {
                        try {
                            $grpObj = Get-ADGroup -Identity $grpDN -ErrorAction Stop
                            $preview += "  + $($grpObj.Name)`n"
                        } catch {
                            $preview += "  + $grpDN`n"
                        }
                    }
                }
                $preview += "`n"
                
                $preview += "GROUPS PRESERVED ($($groupsRemaining.Count)):`n"
                if ($groupsRemaining.Count -eq 0) {
                    $preview += "  (none)`n"
                } else {
                    foreach ($grpName in $groupsRemaining) {
                        $preview += "  = $grpName`n"
                    }
                }
                $preview += "`n"
                
                $preview += "OU CHANGE:`n"
                $currentOU = ($user.DistinguishedName -split ',', 2)[1]
                $preview += "  From: $currentOU`n"
                $preview += "  To:   $toTemplateOU`n"
                
                $PreviewText.Text = $preview
                
                # Confirmation
                $result = [System.Windows.MessageBox]::Show(
                    "Review the changes in the preview area.`n`nProceed with conversion?",
                    "Confirm Conversion",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Question
                )
                
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    
                    # Execute: Remove groups
                    $removedCount = 0
                    foreach ($grp in $groupsToRemove) {
                        try {
                            Remove-ADGroupMember -Identity $grp.DN -Members $user.DistinguishedName -Confirm:$false -ErrorAction Stop
                            $removedCount++
                        } catch {
                        }
                    }
                    
                    # Execute: Add groups
                    $addedCount = 0
                    foreach ($grpDN in $toTemplateGroups) {
                        try {
                            Add-ADGroupMember -Identity $grpDN -Members $user.DistinguishedName -ErrorAction Stop
                            $grpObj = Get-ADGroup -Identity $grpDN -ErrorAction SilentlyContinue
                            $addedCount++
                        } catch {
                        }
                    }
                    
                    # Execute: Move OU
                    try {
                        Move-ADObject -Identity $user.DistinguishedName -TargetPath $toTemplateOU -ErrorAction Stop
                    } catch {
                    }
                    
                    $StatusText.Text = "Conversion complete! Removed $removedCount group(s), added $addedCount group(s), and moved OU."
                    $StatusText.Foreground = "#28a745"
                    
                    [System.Windows.MessageBox]::Show(
                        "Conversion completed successfully!`n`nRemoved: $removedCount group(s)`nAdded: $addedCount group(s)`nMoved to: $toTemplateOU",
                        "Success",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information
                    )
                    
                    
                } else {
                    $StatusText.Text = "Conversion cancelled by user"
                    $StatusText.Foreground = "#6c757d"
                }
                
            } catch {
                $errorMsg = $_.Exception.Message
                $StatusText.Text = "Error: $errorMsg"
                $StatusText.Foreground = "#dc3545"
                [System.Windows.MessageBox]::Show(
                    "Error during conversion:`n`n$errorMsg",
                    "Conversion Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
            } finally {
                $ConvertButton.IsEnabled = $true
            }
        })
        
        # Close button
        $CloseButton.Add_Click({ $window.Close() })
        
        # Show window
        $window.ShowDialog() | Out-Null
        
    } catch {
        $errorMsg = $_.Exception.Message
        [System.Windows.MessageBox]::Show(
            "Failed to open Employee Conversion window:`n`n$errorMsg",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Function to show Locked Out Users window
function Show-LockedOutUsersWindow {
    # Check for Active Directory module
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.MessageBox]::Show(
            "Active Directory module is not available.`n`nThis feature requires the Active Directory PowerShell module (RSAT).",
            "Module Required",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }
    
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        [System.Windows.MessageBox]::Show(
            "Failed to load Active Directory module.`n`n$($_.Exception.Message)",
            "Module Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }
    
    [xml]$LockedUsersXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Locked Out Users Management" 
        Height="600" 
        Width="1000" 
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#dc3545" Padding="20,15">
            <StackPanel>
                <TextBlock Text="Locked Out Users Management" FontSize="20" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="View and unlock user accounts that are currently locked out" FontSize="12" Foreground="#FFE0E0" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Search and Actions Bar -->
        <Border Grid.Row="1" Background="#f8f9fa" Padding="15,10" BorderBrush="#dee2e6" BorderThickness="0,1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <DockPanel Grid.Column="0">
                    <Button x:Name="ClearSearchButton" 
                            Content="Clear" 
                            Width="50" 
                            Height="30" 
                            Margin="5,0,0,0"
                            DockPanel.Dock="Right"
                            Background="#6c757d" 
                            Foreground="White"
                            FontWeight="Bold"
                            ToolTip="Clear search"/>
                    <TextBox x:Name="SearchBox" 
                            Height="30" 
                            Padding="8,5"
                            VerticalContentAlignment="Center"
                            FontSize="13"/>
                </DockPanel>
                
                <StackPanel Grid.Column="1" Orientation="Horizontal" Margin="10,0,0,0">
                    <Button x:Name="RefreshButton" 
                            Content="Refresh" 
                            Width="100" 
                            Height="30" 
                            Margin="0,0,5,0"
                            Background="#007bff" 
                            Foreground="White" 
                            FontWeight="Bold"/>
                    <Button x:Name="UnlockSelectedButton" 
                            Content="Unlock Selected" 
                            Width="130" 
                            Height="30" 
                            Margin="0,0,5,0"
                            Background="#28a745" 
                            Foreground="White" 
                            FontWeight="Bold"
                            IsEnabled="False"/>
                    <Button x:Name="ExportButton" 
                            Content="Export to Excel" 
                            Width="130" 
                            Height="30"
                            Background="#17a2b8" 
                            Foreground="White" 
                            FontWeight="Bold"
                            IsEnabled="False"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Data Grid -->
        <Border Grid.Row="2" BorderBrush="#dee2e6" BorderThickness="1" Margin="15,10">
            <DataGrid x:Name="LockedOutUsersDataGrid" 
                      AutoGenerateColumns="False" 
                      IsReadOnly="True"
                      SelectionMode="Extended"
                      CanUserAddRows="False"
                      CanUserDeleteRows="False"
                      GridLinesVisibility="Horizontal"
                      HeadersVisibility="Column"
                      AlternatingRowBackground="#f8f9fa"
                      RowHeight="28"
                      FontSize="12">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="180"/>
                    <DataGridTextColumn Header="Username" Binding="{Binding SamAccountName}" Width="120"/>
                    <DataGridTextColumn Header="Email" Binding="{Binding EmailAddress}" Width="200"/>
                    <DataGridTextColumn Header="Lockout Time" Binding="{Binding LockoutTimeFormatted}" Width="150"/>
                    <DataGridTextColumn Header="Bad Logon Count" Binding="{Binding BadLogonCount}" Width="110"/>
                    <DataGridTextColumn Header="Department" Binding="{Binding Department}" Width="120"/>
                    <DataGridTextColumn Header="Title" Binding="{Binding Title}" Width="*"/>
                </DataGrid.Columns>
            </DataGrid>
        </Border>
        
        <!-- Status Bar -->
        <Border Grid.Row="3" Background="#f8f9fa" Padding="15,10" BorderBrush="#dee2e6" BorderThickness="0,1,0,0">
            <StackPanel Orientation="Horizontal">
                <TextBlock x:Name="StatusText" Text="Click Refresh to load locked out users" FontSize="12" Foreground="#6c757d"/>
                <TextBlock x:Name="CountText" Text="" FontSize="12" Foreground="#007bff" FontWeight="Bold" Margin="10,0,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Close Button -->
        <Border Grid.Row="4" Background="White" Padding="15,10">
            <Button x:Name="CloseButton" 
                    Content="Close" 
                    Width="100" 
                    Height="35" 
                    HorizontalAlignment="Right"
                    Background="#6c757d" 
                    Foreground="White" 
                    FontWeight="Bold"/>
        </Border>
    </Grid>
</Window>
"@

    try {
        $reader = New-Object System.Xml.XmlNodeReader $LockedUsersXAML
        $LockedUsersWindow = [Windows.Markup.XamlReader]::Load($reader)
        
        # Get controls
        $SearchBox = $LockedUsersWindow.FindName("SearchBox")
        $ClearSearchButton = $LockedUsersWindow.FindName("ClearSearchButton")
        $RefreshButton = $LockedUsersWindow.FindName("RefreshButton")
        $UnlockSelectedButton = $LockedUsersWindow.FindName("UnlockSelectedButton")
        $ExportButton = $LockedUsersWindow.FindName("ExportButton")
        $LockedOutUsersDataGrid = $LockedUsersWindow.FindName("LockedOutUsersDataGrid")
        $StatusText = $LockedUsersWindow.FindName("StatusText")
        $CountText = $LockedUsersWindow.FindName("CountText")
        $CloseButton = $LockedUsersWindow.FindName("CloseButton")
        
        # Store locked users data at script scope for filtering
        $script:lockedOutUsers = @()
        $script:filteredLockedOutUsers = @()
        
        # Refresh button click handler
        $RefreshButton.Add_Click({
            try {
                $StatusText.Text = "Scanning for locked out users..."
                $StatusText.Foreground = "#ffc107"
                $RefreshButton.IsEnabled = $false
                $UnlockSelectedButton.IsEnabled = $false
                $ExportButton.IsEnabled = $false
                $SearchBox.Text = ""
                
                # Search for locked out users
                $lockedUsers = Search-ADAccount -LockedOut -UsersOnly | 
                    Get-ADUser -Properties DisplayName, EmailAddress, LockedOut, LockoutTime, 
                                         BadLogonCount, BadPasswordTime, Department, Title, 
                                         LastLogonDate, Enabled
                
                if ($lockedUsers) {
                    # Convert to array if single result
                    if ($lockedUsers -isnot [System.Array]) {
                        $lockedUsers = @($lockedUsers)
                    }
                    
                    # Always return as array, even for single results
                    $script:lockedOutUsers = @($lockedUsers | ForEach-Object {
                        # Format lockout time
                        $lockoutTimeStr = "N/A"
                        if ($_.LockoutTime -and $_.LockoutTime -ne [DateTime]::MinValue) {
                            try {
                                $lockoutTimeStr = Get-Date $_.LockoutTime -Format "MM/dd/yyyy hh:mm:ss tt"
                            } catch {
                                $lockoutTimeStr = "N/A"
                            }
                        }
                        
                        [PSCustomObject]@{
                            DisplayName = $_.DisplayName
                            SamAccountName = $_.SamAccountName
                            EmailAddress = $_.EmailAddress
                            LockoutTime = $_.LockoutTime
                            LockoutTimeFormatted = $lockoutTimeStr
                            BadLogonCount = $_.BadLogonCount
                            Department = $_.Department
                            Title = $_.Title
                            LastLogonDate = $_.LastLogonDate
                            Enabled = $_.Enabled
                        }
                    })
                    
                    $script:filteredLockedOutUsers = $script:lockedOutUsers
                    $LockedOutUsersDataGrid.ItemsSource = [System.Collections.ArrayList]$script:lockedOutUsers
                    
                    $StatusText.Text = "Loaded locked out users successfully"
                    $StatusText.Foreground = "#28a745"
                    $CountText.Text = "Total: $($script:lockedOutUsers.Count) user(s) locked out"
                    $ExportButton.IsEnabled = $true
                } else {
                    $script:lockedOutUsers = @()
                    $script:filteredLockedOutUsers = @()
                    $LockedOutUsersDataGrid.ItemsSource = $null
                    $StatusText.Text = "No locked out users found"
                    $StatusText.Foreground = "#28a745"
                    $CountText.Text = ""
                }
            } catch {
                $StatusText.Text = "Error: $($_.Exception.Message)"
                $StatusText.Foreground = "#dc3545"
                $CountText.Text = ""
                [System.Windows.MessageBox]::Show(
                    "Failed to query locked out users:`n`n$($_.Exception.Message)",
                    "Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
            } finally {
                $RefreshButton.IsEnabled = $true
            }
        }.GetNewClosure())
        
        # Search box text changed handler
        $SearchBox.Add_TextChanged({
            if ($script:lockedOutUsers.Count -eq 0) { return }
            
            $searchText = $SearchBox.Text.Trim()
            
            if ([string]::IsNullOrEmpty($searchText)) {
                $script:filteredLockedOutUsers = $script:lockedOutUsers
                $LockedOutUsersDataGrid.ItemsSource = [System.Collections.ArrayList]$script:lockedOutUsers
                $CountText.Text = "Total: $($script:lockedOutUsers.Count) user(s) locked out"
            } else {
                # Always return as array, even for single search result
                $script:filteredLockedOutUsers = @($script:lockedOutUsers | Where-Object {
                    $_.DisplayName -like "*$searchText*" -or
                    $_.SamAccountName -like "*$searchText*" -or
                    $_.EmailAddress -like "*$searchText*" -or
                    $_.Department -like "*$searchText*" -or
                    $_.Title -like "*$searchText*"
                })
                
                $LockedOutUsersDataGrid.ItemsSource = [System.Collections.ArrayList]$script:filteredLockedOutUsers
                $CountText.Text = "Showing: $($script:filteredLockedOutUsers.Count) of $($script:lockedOutUsers.Count) user(s)"
            }
        }.GetNewClosure())
        
        # Clear search button
        $ClearSearchButton.Add_Click({
            $SearchBox.Text = ""
        }.GetNewClosure())
        
        # DataGrid selection changed handler
        $LockedOutUsersDataGrid.Add_SelectionChanged({
            $UnlockSelectedButton.IsEnabled = $LockedOutUsersDataGrid.SelectedItems.Count -gt 0
        }.GetNewClosure())
        
        # Unlock selected button click handler
        $UnlockSelectedButton.Add_Click({
            $selectedUsers = @($LockedOutUsersDataGrid.SelectedItems)
            
            if ($selectedUsers.Count -eq 0) {
                [System.Windows.MessageBox]::Show(
                    "Please select at least one user to unlock.",
                    "No Selection",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                return
            }
            
            # Confirm action
            $confirmResult = [System.Windows.MessageBox]::Show(
                "Are you sure you want to unlock $($selectedUsers.Count) user account$(if($selectedUsers.Count -ne 1){'s'})?`n`nThis will allow the user$(if($selectedUsers.Count -ne 1){'s'}) to log in again.",
                "Confirm Unlock",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question
            )
            
            if ($confirmResult -eq [System.Windows.MessageBoxResult]::Yes) {
                $StatusText.Text = "Unlocking selected users..."
                $StatusText.Foreground = "#ffc107"
                $UnlockSelectedButton.IsEnabled = $false
                
                $successCount = 0
                $failCount = 0
                $errorMessages = @()
                
                foreach ($user in $selectedUsers) {
                    try {
                        Unlock-ADAccount -Identity $user.SamAccountName -ErrorAction Stop
                        $successCount++
                    } catch {
                        $failCount++
                        $errorMessages += "$($user.SamAccountName): $($_.Exception.Message)"
                    }
                }
                
                # Show results
                if ($failCount -eq 0) {
                    [System.Windows.MessageBox]::Show(
                        "Successfully unlocked $successCount user account$(if($successCount -ne 1){'s'}).",
                        "Success",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information
                    )
                    $StatusText.Text = "Successfully unlocked $successCount user(s)"
                    $StatusText.Foreground = "#28a745"
                    
                    # Refresh the list
                    $RefreshButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                } else {
                    $message = "Unlocked: $successCount`nFailed: $failCount`n`nErrors:`n" + ($errorMessages -join "`n")
                    [System.Windows.MessageBox]::Show(
                        $message,
                        "Partial Success",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                    $StatusText.Text = "Unlocked $successCount user(s), $failCount failed"
                    $StatusText.Foreground = "#ffc107"
                    
                    # Refresh the list
                    $RefreshButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                }
            }
        }.GetNewClosure())
        
        # Export to Excel button
        $ExportButton.Add_Click({
            if ($script:filteredLockedOutUsers.Count -eq 0) {
                [System.Windows.MessageBox]::Show(
                    "No data to export. Please load locked out users first.",
                    "No Data",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                return
            }
            
            try {
                # Create SaveFileDialog
                $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
                $saveDialog.Title = "Save Locked Out Users Report"
                $saveDialog.FileName = "LockedOutUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
                
                if ($saveDialog.ShowDialog()) {
                    $excelPath = $saveDialog.FileName
                    
                    $StatusText.Text = "Exporting to Excel..."
                    $StatusText.Foreground = "#ffc107"
                    $ExportButton.IsEnabled = $false
                    
                    # Prepare data for export
                    $exportData = $script:filteredLockedOutUsers | Select-Object `
                        DisplayName,
                        SamAccountName,
                        EmailAddress,
                        LockoutTimeFormatted,
                        BadLogonCount,
                        Department,
                        Title,
                        @{Name='Enabled';Expression={if($_.Enabled){'Yes'}else{'No'}}}
                    
                    # Export to Excel with formatting
                    $exportData | Export-Excel -Path $excelPath `
                        -AutoSize `
                        -AutoFilter `
                        -FreezeTopRow `
                        -BoldTopRow `
                        -TableName "LockedOutUsers" `
                        -TableStyle Medium2
                    
                    $StatusText.Text = "Exported successfully"
                    $StatusText.Foreground = "#28a745"
                    
                    [System.Windows.MessageBox]::Show(
                        "Locked out users exported successfully to:`n`n$excelPath",
                        "Export Complete",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Information
                    )
                }
            } catch {
                $StatusText.Text = "Export failed"
                $StatusText.Foreground = "#dc3545"
                [System.Windows.MessageBox]::Show(
                    "Failed to export to Excel:`n`n$($_.Exception.Message)",
                    "Export Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
            } finally {
                $ExportButton.IsEnabled = $true
            }
        }.GetNewClosure())
        
        # Close button
        $CloseButton.Add_Click({
            $LockedUsersWindow.Close()
        })
        
        # Show the window
        $LockedUsersWindow.ShowDialog() | Out-Null
        
    } catch {
        $errorMsg = $_.Exception.Message
        [System.Windows.MessageBox]::Show(
            "Failed to open Locked Out Users window:`n`n$errorMsg",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}


$syncHash.AutoRepliesButton.Add_Click({
    # Check if connected
    $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
        $result = [System.Windows.MessageBox]::Show(
            "You are not connected to Exchange Online.`n`nWould you like to connect now?",
            "Connection Required",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $syncHash.ConnectButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            
            # Wait a moment for connection to establish, then recheck
            Start-Sleep -Seconds 5
            $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
            
            if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
                Write-Log "Connection was not established. Please try again."
                return
            }
        } else {
            return
        }
    }
    
    Write-Log "Opening Automatic Replies window..."
    
    [xml]$AutoRepliesXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Automatic Replies (Out of Office) Management" 
        Height="650" 
        Width="750" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <GroupBox Grid.Row="0" Header="Mailbox Information" Margin="15,15,15,10" Padding="15">
            <StackPanel>
                <TextBlock Text="Mailbox Email Address:" FontWeight="Bold" Margin="0,0,0,5"/>
                <DockPanel Margin="0,0,0,10">
                    <Button x:Name="LoadAutoRepliesButton" 
                           Content="Load Settings" 
                           Width="120" 
                           Height="30" 
                           Margin="10,0,0,0" 
                           DockPanel.Dock="Right" 
                           Background="#007bff" 
                           Foreground="White" 
                           FontWeight="Bold"/>
                    <TextBox x:Name="MailboxIdentityBox" 
                            Height="30" 
                            Padding="5"
                            VerticalContentAlignment="Center"/>
                </DockPanel>
                
                <StackPanel x:Name="StatusPanel" Visibility="Collapsed" Margin="0,10,0,0">
                    <Border x:Name="StatusBorder" BorderThickness="1" Padding="10" CornerRadius="3">
                        <StackPanel>
                            <TextBlock x:Name="MailboxNameText" FontWeight="Bold" Margin="0,0,0,5"/>
                            <TextBlock x:Name="AutoReplyStateText" FontSize="11" Margin="0,0,0,3"/>
                            <TextBlock x:Name="ScheduledText" FontSize="11" Margin="0,0,0,3"/>
                        </StackPanel>
                    </Border>
                </StackPanel>
            </StackPanel>
        </GroupBox>
        
        <TabControl Grid.Row="1" Margin="15,0,15,10" x:Name="SettingsTabControl" IsEnabled="False">
            <TabItem Header="Auto Reply Settings">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="20">
                        <GroupBox Header="Status" Padding="10" Margin="0,0,0,15">
                            <StackPanel>
                                <RadioButton x:Name="DisabledRadio" Content="Disabled" GroupName="AutoReplyState" Margin="0,5" FontSize="13"/>
                                <RadioButton x:Name="EnabledRadio" Content="Enabled" GroupName="AutoReplyState" Margin="0,5" FontSize="13"/>
                                <RadioButton x:Name="ScheduledRadio" Content="Scheduled (Time Range)" GroupName="AutoReplyState" Margin="0,5" FontSize="13"/>
                            </StackPanel>
                        </GroupBox>
                        
                        <GroupBox Header="Schedule (Only for Scheduled)" Padding="10" Margin="0,0,0,15" x:Name="ScheduleGroup">
                            <StackPanel>
                                <StackPanel Orientation="Horizontal" Margin="0,5">
                                    <TextBlock Text="Start Date/Time:" Width="120" VerticalAlignment="Center" FontWeight="Bold"/>
                                    <DatePicker x:Name="StartDatePicker" Width="150" Margin="0,0,10,0"/>
                                    <TextBox x:Name="StartTimeBox" Width="80" Height="25" Padding="5" ToolTip="HH:mm format (e.g., 09:00)"/>
                                </StackPanel>
                                <StackPanel Orientation="Horizontal" Margin="0,5">
                                    <TextBlock Text="End Date/Time:" Width="120" VerticalAlignment="Center" FontWeight="Bold"/>
                                    <DatePicker x:Name="EndDatePicker" Width="150" Margin="0,0,10,0"/>
                                    <TextBox x:Name="EndTimeBox" Width="80" Height="25" Padding="5" ToolTip="HH:mm format (e.g., 17:00)"/>
                                </StackPanel>
                                <TextBlock Text="Time format: HH:mm (24-hour, e.g., 09:00 or 17:30)" 
                                          FontSize="10" 
                                          Foreground="Gray" 
                                          FontStyle="Italic" 
                                          Margin="120,5,0,0"/>
                            </StackPanel>
                        </GroupBox>
                        
                        <GroupBox Header="Internal Message (to people in your organization)" Padding="10" Margin="0,0,0,15">
                            <DockPanel>
                                <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,5" Background="#f0f0f0" Height="35">
                                    <Button x:Name="InternalBoldButton" Content="B" Width="30" Height="25" Margin="5,5,2,5" FontWeight="Bold" ToolTip="Bold"/>
                                    <Button x:Name="InternalItalicButton" Content="I" Width="30" Height="25" Margin="2,5" FontStyle="Italic" ToolTip="Italic"/>
                                    <Button x:Name="InternalUnderlineButton" Content="U" Width="30" Height="25" Margin="2,5" ToolTip="Underline">
                                        <Button.Template>
                                            <ControlTemplate TargetType="Button">
                                                <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1">
                                                    <TextBlock Text="{TemplateBinding Content}" TextDecorations="Underline" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                </Border>
                                            </ControlTemplate>
                                        </Button.Template>
                                    </Button>
                                    <Separator Width="1" Margin="5,5"/>
                                    <Button x:Name="InternalClearButton" Content="Clear Format" Width="90" Height="25" Margin="5" ToolTip="Remove all formatting"/>
                                </StackPanel>
                                <Border BorderBrush="#dee2e6" BorderThickness="1" Padding="5">
                                    <RichTextBox x:Name="InternalRichTextBox" 
                                                Height="120" 
                                                VerticalScrollBarVisibility="Auto"
                                                Background="White"
                                                AcceptsReturn="True"/>
                                </Border>
                            </DockPanel>
                        </GroupBox>
                        
                        <GroupBox Header="External Message (to people outside your organization)" Padding="10" Margin="0,0,0,15">
                            <StackPanel>
                                <CheckBox x:Name="ExternalEnabledCheck" Content="Send automatic replies to external senders" Margin="0,0,0,10" FontWeight="Bold"/>
                                <RadioButton x:Name="ExternalAllRadio" Content="Send to all external senders" GroupName="ExternalAudience" Margin="0,5" IsEnabled="False"/>
                                <RadioButton x:Name="ExternalKnownRadio" Content="Send to external senders in my contacts only" GroupName="ExternalAudience" Margin="0,5" IsEnabled="False"/>
                                <DockPanel Margin="0,10,0,0">
                                    <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,5" Background="#f0f0f0" Height="35">
                                        <Button x:Name="ExternalBoldButton" Content="B" Width="30" Height="25" Margin="5,5,2,5" FontWeight="Bold" ToolTip="Bold" IsEnabled="False"/>
                                        <Button x:Name="ExternalItalicButton" Content="I" Width="30" Height="25" Margin="2,5" FontStyle="Italic" ToolTip="Italic" IsEnabled="False"/>
                                        <Button x:Name="ExternalUnderlineButton" Content="U" Width="30" Height="25" Margin="2,5" ToolTip="Underline" IsEnabled="False">
                                            <Button.Template>
                                                <ControlTemplate TargetType="Button">
                                                    <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1">
                                                        <TextBlock Text="{TemplateBinding Content}" TextDecorations="Underline" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                    </Border>
                                                </ControlTemplate>
                                            </Button.Template>
                                        </Button>
                                        <Separator Width="1" Margin="5,5"/>
                                        <Button x:Name="ExternalClearButton" Content="Clear Format" Width="90" Height="25" Margin="5" ToolTip="Remove all formatting" IsEnabled="False"/>
                                    </StackPanel>
                                    <Border BorderBrush="#dee2e6" BorderThickness="1" Padding="5">
                                        <RichTextBox x:Name="ExternalRichTextBox" 
                                                    Height="100" 
                                                    VerticalScrollBarVisibility="Auto"
                                                    Background="White"
                                                    AcceptsReturn="True"
                                                    IsEnabled="False"/>
                                    </Border>
                                </DockPanel>
                            </StackPanel>
                        </GroupBox>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>
        
        <Border Grid.Row="2" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="0,1,0,0" Padding="15">
            <DockPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" DockPanel.Dock="Left">
                    <TextBlock x:Name="StatusTextBlock" 
                              Text="Enter a mailbox email address to begin" 
                              VerticalAlignment="Center"
                              FontSize="11"
                              Foreground="#666"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="SaveButton" 
                           Content="Save Changes" 
                           Width="130" 
                           Height="32" 
                           Margin="0,0,10,0" 
                           Background="#28a745" 
                           Foreground="White" 
                           FontWeight="Bold"
                           IsEnabled="False"/>
                    <Button x:Name="CloseButton" 
                           Content="Close" 
                           Width="80" 
                           Height="32" 
                           Background="#6c757d" 
                           Foreground="White"/>
                </StackPanel>
            </DockPanel>
        </Border>
    </Grid>
</Window>
"@
    
    $autoReader = New-Object System.Xml.XmlNodeReader $AutoRepliesXAML
    $AutoWindow = [Windows.Markup.XamlReader]::Load($autoReader)
    $AutoWindow.Owner = $syncHash.Window
    
    # Get controls
    $MailboxIdentityBox = $AutoWindow.FindName("MailboxIdentityBox")
    $LoadAutoRepliesButton = $AutoWindow.FindName("LoadAutoRepliesButton")
    $StatusPanel = $AutoWindow.FindName("StatusPanel")
    $StatusBorder = $AutoWindow.FindName("StatusBorder")
    $MailboxNameText = $AutoWindow.FindName("MailboxNameText")
    $AutoReplyStateText = $AutoWindow.FindName("AutoReplyStateText")
    $ScheduledText = $AutoWindow.FindName("ScheduledText")
    
    $SettingsTabControl = $AutoWindow.FindName("SettingsTabControl")
    $DisabledRadio = $AutoWindow.FindName("DisabledRadio")
    $EnabledRadio = $AutoWindow.FindName("EnabledRadio")
    $ScheduledRadio = $AutoWindow.FindName("ScheduledRadio")
    $ScheduleGroup = $AutoWindow.FindName("ScheduleGroup")
    $StartDatePicker = $AutoWindow.FindName("StartDatePicker")
    $StartTimeBox = $AutoWindow.FindName("StartTimeBox")
    $EndDatePicker = $AutoWindow.FindName("EndDatePicker")
    $EndTimeBox = $AutoWindow.FindName("EndTimeBox")
    
    $InternalRichTextBox = $AutoWindow.FindName("InternalRichTextBox")
    $InternalBoldButton = $AutoWindow.FindName("InternalBoldButton")
    $InternalItalicButton = $AutoWindow.FindName("InternalItalicButton")
    $InternalUnderlineButton = $AutoWindow.FindName("InternalUnderlineButton")
    $InternalClearButton = $AutoWindow.FindName("InternalClearButton")
    
    $ExternalEnabledCheck = $AutoWindow.FindName("ExternalEnabledCheck")
    $ExternalAllRadio = $AutoWindow.FindName("ExternalAllRadio")
    $ExternalKnownRadio = $AutoWindow.FindName("ExternalKnownRadio")
    $ExternalRichTextBox = $AutoWindow.FindName("ExternalRichTextBox")
    $ExternalBoldButton = $AutoWindow.FindName("ExternalBoldButton")
    $ExternalItalicButton = $AutoWindow.FindName("ExternalItalicButton")
    $ExternalUnderlineButton = $AutoWindow.FindName("ExternalUnderlineButton")
    $ExternalClearButton = $AutoWindow.FindName("ExternalClearButton")
    
    $StatusTextBlock = $AutoWindow.FindName("StatusTextBlock")
    $SaveButton = $AutoWindow.FindName("SaveButton")
    $CloseButton = $AutoWindow.FindName("CloseButton")
    
    $script:currentMailboxSettings = $null
    
    # Function to convert RichTextBox to HTML
    function Get-RichTextBoxHtml {
        param(
            [System.Windows.Controls.RichTextBox]$RichTextBox
        )
        
        $textRange = New-Object System.Windows.Documents.TextRange($RichTextBox.Document.ContentStart, $RichTextBox.Document.ContentEnd)
        $text = $textRange.Text
        
        if ([string]::IsNullOrWhiteSpace($text)) {
            return ""
        }
        
        $html = "<html><body style='font-family: Calibri, Arial, sans-serif; font-size: 11pt;'>"
        
        foreach ($block in $RichTextBox.Document.Blocks) {
            if ($block -is [System.Windows.Documents.Paragraph]) {
                $html += "<p>"
                
                foreach ($inline in $block.Inlines) {
                    if ($inline -is [System.Windows.Documents.Run]) {
                        $runText = [System.Net.WebUtility]::HtmlEncode($inline.Text)
                        
                        $isBold = $inline.FontWeight -eq [System.Windows.FontWeights]::Bold
                        $isItalic = $inline.FontStyle -eq [System.Windows.FontStyles]::Italic
                        $isUnderline = $inline.TextDecorations.Count -gt 0
                        
                        if ($isBold) { $runText = "<b>$runText</b>" }
                        if ($isItalic) { $runText = "<i>$runText</i>" }
                        if ($isUnderline) { $runText = "<u>$runText</u>" }
                        
                        $html += $runText
                    } elseif ($inline -is [System.Windows.Documents.LineBreak]) {
                        $html += "<br/>"
                    }
                }
                
                $html += "</p>"
            }
        }
        
        $html += "</body></html>"
        return $html
    }
    
    # Function to set RichTextBox from HTML
    function Set-RichTextBoxFromHtml {
        param(
            [System.Windows.Controls.RichTextBox]$RichTextBox,
            [string]$HtmlContent
        )
        
        $RichTextBox.Document.Blocks.Clear()
        
        if ([string]::IsNullOrWhiteSpace($HtmlContent)) {
            return
        }
        
        try {
            # Clean HTML
            $cleanHtml = $HtmlContent -replace '<html[^>]*>', '' -replace '</html>', ''
            $cleanHtml = $cleanHtml -replace '<body[^>]*>', '' -replace '</body>', ''
            $cleanHtml = $cleanHtml -replace '<head>.*?</head>', ''
            
            # Split by paragraphs
            $paragraphs = $cleanHtml -split '<p>|</p>' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            
            foreach ($paraText in $paragraphs) {
                $para = New-Object System.Windows.Documents.Paragraph
                $para.Margin = New-Object System.Windows.Thickness(0)
                
                # Process inline elements
                $currentText = $paraText
                $position = 0
                
                while ($position -lt $currentText.Length) {
                    # Check for tags
                    if ($currentText[$position] -eq '<') {
                        $tagEnd = $currentText.IndexOf('>', $position)
                        if ($tagEnd -gt $position) {
                            $tag = $currentText.Substring($position, $tagEnd - $position + 1)
                            
                            if ($tag -match '<(b|strong|i|em|u)>') {
                                $tagName = $matches[1]
                                $closeTag = "</$tagName>"
                                $closePos = $currentText.IndexOf($closeTag, $tagEnd)
                                
                                if ($closePos -gt $tagEnd) {
                                    $innerText = $currentText.Substring($tagEnd + 1, $closePos - $tagEnd - 1)
                                    $innerText = [System.Net.WebUtility]::HtmlDecode($innerText)
                                    
                                    $run = New-Object System.Windows.Documents.Run($innerText)
                                    
                                    if ($tagName -eq 'b' -or $tagName -eq 'strong') {
                                        $run.FontWeight = [System.Windows.FontWeights]::Bold
                                    }
                                    if ($tagName -eq 'i' -or $tagName -eq 'em') {
                                        $run.FontStyle = [System.Windows.FontStyles]::Italic
                                    }
                                    if ($tagName -eq 'u') {
                                        $run.TextDecorations = [System.Windows.TextDecorations]::Underline
                                    }
                                    
                                    $para.Inlines.Add($run)
                                    $position = $closePos + $closeTag.Length
                                    continue
                                }
                            } elseif ($tag -eq '<br>' -or $tag -eq '<br/>') {
                                $para.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
                                $position = $tagEnd + 1
                                continue
                            }
                            
                            $position = $tagEnd + 1
                        } else {
                            $position++
                        }
                    } else {
                        # Find next tag or end
                        $nextTag = $currentText.IndexOf('<', $position)
                        if ($nextTag -eq -1) { $nextTag = $currentText.Length }
                        
                        $plainText = $currentText.Substring($position, $nextTag - $position)
                        $plainText = [System.Net.WebUtility]::HtmlDecode($plainText)
                        
                        if (-not [string]::IsNullOrWhiteSpace($plainText)) {
                            $run = New-Object System.Windows.Documents.Run($plainText)
                            $para.Inlines.Add($run)
                        }
                        
                        $position = $nextTag
                    }
                }
                
                $RichTextBox.Document.Blocks.Add($para)
            }
            
        } catch {
            # Fallback: just add as plain text
            $plainText = $HtmlContent -replace '<[^>]+>', ''
            $plainText = [System.Net.WebUtility]::HtmlDecode($plainText)
            $para = New-Object System.Windows.Documents.Paragraph
            $para.Inlines.Add((New-Object System.Windows.Documents.Run($plainText)))
            $RichTextBox.Document.Blocks.Clear()
            $RichTextBox.Document.Blocks.Add($para)
        }
    }
    
    # Formatting button handlers for Internal message
    $InternalBoldButton.Add_Click({
        $selection = $InternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $currentWeight = $selection.GetPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty)
            if ($currentWeight -eq [System.Windows.FontWeights]::Bold) {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Normal)
            } else {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Bold)
            }
        }
        $InternalRichTextBox.Focus()
    })
    
    $InternalItalicButton.Add_Click({
        $selection = $InternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $currentStyle = $selection.GetPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty)
            if ($currentStyle -eq [System.Windows.FontStyles]::Italic) {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Normal)
            } else {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Italic)
            }
        }
        $InternalRichTextBox.Focus()
    })
    
    $InternalUnderlineButton.Add_Click({
        $selection = $InternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $currentDeco = $selection.GetPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty)
            if ($currentDeco -eq [System.Windows.TextDecorations]::Underline) {
                $selection.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, $null)
            } else {
                $selection.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, [System.Windows.TextDecorations]::Underline)
            }
        }
        $InternalRichTextBox.Focus()
    })
    
    $InternalClearButton.Add_Click({
        $selection = $InternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Normal)
            $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Normal)
            $selection.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, $null)
        }
        $InternalRichTextBox.Focus()
    })
    
    # Formatting button handlers for External message
    $ExternalBoldButton.Add_Click({
        $selection = $ExternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $currentWeight = $selection.GetPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty)
            if ($currentWeight -eq [System.Windows.FontWeights]::Bold) {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Normal)
            } else {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Bold)
            }
        }
        $ExternalRichTextBox.Focus()
    })
    
    $ExternalItalicButton.Add_Click({
        $selection = $ExternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $currentStyle = $selection.GetPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty)
            if ($currentStyle -eq [System.Windows.FontStyles]::Italic) {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Normal)
            } else {
                $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Italic)
            }
        }
        $ExternalRichTextBox.Focus()
    })
    
    $ExternalUnderlineButton.Add_Click({
        $selection = $ExternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $currentDeco = $selection.GetPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty)
            if ($currentDeco -eq [System.Windows.TextDecorations]::Underline) {
                $selection.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, $null)
            } else {
                $selection.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, [System.Windows.TextDecorations]::Underline)
            }
        }
        $ExternalRichTextBox.Focus()
    })
    
    $ExternalClearButton.Add_Click({
        $selection = $ExternalRichTextBox.Selection
        if (-not $selection.IsEmpty) {
            $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Normal)
            $selection.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Normal)
            $selection.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, $null)
        }
        $ExternalRichTextBox.Focus()
    })
    
    # Enable/disable schedule fields based on radio selection
    $DisabledRadio.Add_Checked({ 
        $ScheduleGroup.IsEnabled = $false 
    })
    $EnabledRadio.Add_Checked({ 
        $ScheduleGroup.IsEnabled = $false 
    })
    $ScheduledRadio.Add_Checked({ 
        $ScheduleGroup.IsEnabled = $true 
    })
    
    # Enable/disable external message fields
    $ExternalEnabledCheck.Add_Checked({
        $ExternalAllRadio.IsEnabled = $true
        $ExternalKnownRadio.IsEnabled = $true
        $ExternalRichTextBox.IsEnabled = $true
        $ExternalBoldButton.IsEnabled = $true
        $ExternalItalicButton.IsEnabled = $true
        $ExternalUnderlineButton.IsEnabled = $true
        $ExternalClearButton.IsEnabled = $true
        if (-not $ExternalAllRadio.IsChecked -and -not $ExternalKnownRadio.IsChecked) {
            $ExternalAllRadio.IsChecked = $true
        }
    })
    
    $ExternalEnabledCheck.Add_Unchecked({
        $ExternalAllRadio.IsEnabled = $false
        $ExternalKnownRadio.IsEnabled = $false
        $ExternalRichTextBox.IsEnabled = $false
        $ExternalBoldButton.IsEnabled = $false
        $ExternalItalicButton.IsEnabled = $false
        $ExternalUnderlineButton.IsEnabled = $false
        $ExternalClearButton.IsEnabled = $false
    })
    
    $LoadAutoRepliesButton.Add_Click({
        $mailboxIdentity = $MailboxIdentityBox.Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            [System.Windows.MessageBox]::Show("Please enter a mailbox email address", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $LoadAutoRepliesButton.IsEnabled = $false
            $StatusTextBlock.Text = "Loading automatic reply settings..."
            Write-Log "Loading automatic reply settings for: $mailboxIdentity"
            
            # Get mailbox info
            $mailbox = Get-Mailbox -Identity $mailboxIdentity -ErrorAction Stop
            
            # Get automatic reply configuration
            $autoReplyConfig = Get-MailboxAutoReplyConfiguration -Identity $mailboxIdentity -ErrorAction Stop
            
            $script:currentMailboxSettings = $autoReplyConfig
            
            # Update status panel
            $MailboxNameText.Text = "Mailbox: $($mailbox.DisplayName)"
            
            switch ($autoReplyConfig.AutoReplyState) {
                "Disabled" {
                    $AutoReplyStateText.Text = "Status: Disabled"
                    $AutoReplyStateText.Foreground = [System.Windows.Media.Brushes]::Gray
                    $StatusBorder.Background = [System.Windows.Media.Brushes]::LightGray
                    $StatusBorder.BorderBrush = [System.Windows.Media.Brushes]::Gray
                    $DisabledRadio.IsChecked = $true
                }
                "Enabled" {
                    $AutoReplyStateText.Text = "Status: Enabled"
                    $AutoReplyStateText.Foreground = [System.Windows.Media.Brushes]::Green
                    $StatusBorder.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(230, 255, 230))
                    $StatusBorder.BorderBrush = [System.Windows.Media.Brushes]::Green
                    $EnabledRadio.IsChecked = $true
                }
                "Scheduled" {
                    $AutoReplyStateText.Text = "Status: Scheduled"
                    $AutoReplyStateText.Foreground = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255, 140, 0))
                    $StatusBorder.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255, 248, 220))
                    $StatusBorder.BorderBrush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255, 140, 0))
                    $ScheduledRadio.IsChecked = $true
                }
            }
            
            if ($autoReplyConfig.AutoReplyState -eq "Scheduled") {
                $startLocal = $autoReplyConfig.StartTime.ToLocalTime()
                $endLocal = $autoReplyConfig.EndTime.ToLocalTime()
                $ScheduledText.Text = "Active: $($startLocal.ToString('g')) to $($endLocal.ToString('g'))"
                $ScheduledText.Visibility = [System.Windows.Visibility]::Visible
                
                $StartDatePicker.SelectedDate = $startLocal
                $StartTimeBox.Text = $startLocal.ToString("HH:mm")
                $EndDatePicker.SelectedDate = $endLocal
                $EndTimeBox.Text = $endLocal.ToString("HH:mm")
            } else {
                $ScheduledText.Visibility = [System.Windows.Visibility]::Collapsed
                $StartDatePicker.SelectedDate = (Get-Date).Date
                $StartTimeBox.Text = "09:00"
                $EndDatePicker.SelectedDate = (Get-Date).Date.AddDays(7)
                $EndTimeBox.Text = "17:00"
            }
            
            $StatusPanel.Visibility = [System.Windows.Visibility]::Visible
            
            # Load and render messages
            Set-RichTextBoxFromHtml -RichTextBox $InternalRichTextBox -HtmlContent $autoReplyConfig.InternalMessage
            Set-RichTextBoxFromHtml -RichTextBox $ExternalRichTextBox -HtmlContent $autoReplyConfig.ExternalMessage
            
            # External audience
            if ($autoReplyConfig.ExternalAudience -eq "None") {
                $ExternalEnabledCheck.IsChecked = $false
            } else {
                $ExternalEnabledCheck.IsChecked = $true
                if ($autoReplyConfig.ExternalAudience -eq "All") {
                    $ExternalAllRadio.IsChecked = $true
                } else {
                    $ExternalKnownRadio.IsChecked = $true
                }
            }
            
            $SettingsTabControl.IsEnabled = $true
            $SaveButton.IsEnabled = $true
            
            $StatusTextBlock.Text = "Settings loaded successfully"
            Write-Log "Successfully loaded automatic reply settings for $($mailbox.DisplayName)"
            
        } catch {
            Write-Log "Error loading automatic reply settings: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error loading settings:`n`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            $StatusTextBlock.Text = "Error loading settings"
            $StatusPanel.Visibility = [System.Windows.Visibility]::Collapsed
            $SettingsTabControl.IsEnabled = $false
            $SaveButton.IsEnabled = $false
        } finally {
            $LoadAutoRepliesButton.IsEnabled = $true
        }
    })
    
    $SaveButton.Add_Click({
        $mailboxIdentity = $MailboxIdentityBox.Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
            [System.Windows.MessageBox]::Show("No mailbox loaded", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # Determine state
        $newState = "Disabled"
        if ($EnabledRadio.IsChecked) {
            $newState = "Enabled"
        } elseif ($ScheduledRadio.IsChecked) {
            $newState = "Scheduled"
        }
        
        # Validate scheduled dates if needed
        if ($newState -eq "Scheduled") {
            if (-not $StartDatePicker.SelectedDate -or -not $EndDatePicker.SelectedDate) {
                [System.Windows.MessageBox]::Show("Please select both start and end dates for scheduled automatic replies", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            if ([string]::IsNullOrWhiteSpace($StartTimeBox.Text) -or [string]::IsNullOrWhiteSpace($EndTimeBox.Text)) {
                [System.Windows.MessageBox]::Show("Please enter both start and end times in HH:mm format (e.g., 09:00)", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            # Validate time format
            if ($StartTimeBox.Text -notmatch '^\d{2}:\d{2}$' -or $EndTimeBox.Text -notmatch '^\d{2}:\d{2}$') {
                [System.Windows.MessageBox]::Show("Time must be in HH:mm format (e.g., 09:00 or 17:30)", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            try {
                $startTime = [datetime]::Parse($StartTimeBox.Text)
                $endTime = [datetime]::Parse($EndTimeBox.Text)
            } catch {
                [System.Windows.MessageBox]::Show("Invalid time format. Please use HH:mm (e.g., 09:00)", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            $startDateTime = $StartDatePicker.SelectedDate.Date.Add($startTime.TimeOfDay)
            $endDateTime = $EndDatePicker.SelectedDate.Date.Add($endTime.TimeOfDay)
            
            if ($endDateTime -le $startDateTime) {
                [System.Windows.MessageBox]::Show("End date/time must be after start date/time", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
        }
        
        # Determine external audience
        $externalAudience = "None"
        if ($ExternalEnabledCheck.IsChecked) {
            if ($ExternalAllRadio.IsChecked) {
                $externalAudience = "All"
            } else {
                $externalAudience = "Known"
            }
        }
        
        $result = [System.Windows.MessageBox]::Show("Save automatic reply settings for this mailbox?", "Confirm", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                $SaveButton.IsEnabled = $false
                $StatusTextBlock.Text = "Saving settings..."
                Write-Log "Saving automatic reply settings for $mailboxIdentity"
                
                # Convert RichTextBox content to HTML
                $internalHtml = Get-RichTextBoxHtml -RichTextBox $InternalRichTextBox
                $externalHtml = Get-RichTextBoxHtml -RichTextBox $ExternalRichTextBox
                
                $setParams = @{
                    Identity = $mailboxIdentity
                    AutoReplyState = $newState
                    InternalMessage = $internalHtml
                    ExternalMessage = $externalHtml
                    ExternalAudience = $externalAudience
                }
                
                if ($newState -eq "Scheduled") {
                    $setParams.StartTime = $startDateTime
                    $setParams.EndTime = $endDateTime
                }
                
                Set-MailboxAutoReplyConfiguration @setParams -ErrorAction Stop
                
                Write-Log "Successfully saved automatic reply settings"
                $StatusTextBlock.Text = "Settings saved successfully"
                [System.Windows.MessageBox]::Show("Automatic reply settings saved successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                
                # Reload to show updated status
                $LoadAutoRepliesButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                
            } catch {
                Write-Log "Error saving automatic reply settings: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Error saving settings:`n`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                $StatusTextBlock.Text = "Error saving settings"
            } finally {
                $SaveButton.IsEnabled = $true
            }
        }
    })
    
    $CloseButton.Add_Click({ $AutoWindow.Close() })
    
    # Add Enter key support
    $MailboxIdentityBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $LoadAutoRepliesButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $AutoWindow.ShowDialog() | Out-Null
})

$syncHash.MessageTraceButton.Add_Click({
    # Check if connected
    $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
        $result = [System.Windows.MessageBox]::Show(
            "You are not connected to Exchange Online.`n`nWould you like to connect now?",
            "Connection Required",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $syncHash.ConnectButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            
            # Wait a moment for connection to establish, then recheck
            Start-Sleep -Seconds 5
            $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
            
            if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
                Write-Log "Connection was not established. Please try again."
                return
            }
        } else {
            return
        }
    }
    
    Write-Log "Opening Message Trace window..."
    
    [xml]$MessageTraceXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Message Trace / Tracking" 
        Height="750" 
        Width="900" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <TabControl Margin="10">
            <TabItem Header="Search">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="Search Criteria" Padding="15" Margin="0,0,0,15">
                        <StackPanel>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="15"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                
                                <StackPanel Grid.Column="0">
                                    <TextBlock Text="Sender Email Address:" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <TextBox x:Name="SenderEmailBox" MinHeight="30" Padding="5,8,5,5" VerticalContentAlignment="Center" Margin="0,0,0,15"/>

                                    
                                    <TextBlock Text="Recipient Email Address:" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <TextBox x:Name="RecipientEmailBox" MinHeight="30" Padding="5,8,5,5" VerticalContentAlignment="Center" Margin="0,0,0,15"/>
                                    
                                    <TextBlock Text="Subject Contains:" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <TextBox x:Name="SubjectBox" MinHeight="30" Padding="5,8,5,5" VerticalContentAlignment="Center"/>
                                </StackPanel>
                                
                                <StackPanel Grid.Column="2">
                                    <TextBlock Text="Message ID:" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <TextBox x:Name="MessageIdBox" MinHeight="30" Padding="5,8,5,5" VerticalContentAlignment="Center" Margin="0,0,0,15"/>

                                    <TextBlock Text="Status:" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <ComboBox x:Name="StatusCombo" MinHeight="30" Margin="0,0,0,15" SelectedIndex="0">
                                        <ComboBoxItem Content="All" Tag="All"/>
                                        <ComboBoxItem Content="Delivered" Tag="Delivered"/>
                                        <ComboBoxItem Content="Failed" Tag="Failed"/>
                                        <ComboBoxItem Content="Pending" Tag="Pending"/>
                                        <ComboBoxItem Content="Quarantined" Tag="Quarantined"/>
                                        <ComboBoxItem Content="FilteredAsSpam" Tag="FilteredAsSpam"/>
                                    </ComboBox>
                                    
                                    <TextBlock Text="Page Size (max results):" FontWeight="Bold" Margin="0,0,0,5"/>
                                    <ComboBox x:Name="PageSizeCombo" MinHeight="30" SelectedIndex="1">
                                        <ComboBoxItem Content="100" Tag="100"/>
                                        <ComboBoxItem Content="1000" Tag="1000"/>
                                        <ComboBoxItem Content="5000" Tag="5000"/>
                                    </ComboBox>
                                </StackPanel>
                            </Grid>
                        </StackPanel>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="1" Header="Date Range (Max 10 Days)" Padding="15" Margin="0,0,0,15">
                        <StackPanel>
                            <RadioButton x:Name="Last24HoursRadio" Content="Last 24 Hours" GroupName="DateRange" Margin="0,5" IsChecked="True"/>
                            <RadioButton x:Name="Last7DaysRadio" Content="Last 7 Days" GroupName="DateRange" Margin="0,5"/>
                            <RadioButton x:Name="CustomRangeRadio" Content="Custom Date Range" GroupName="DateRange" Margin="0,5"/>
                            
                            <StackPanel x:Name="CustomDatePanel" Margin="20,10,0,0" IsEnabled="False">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="15"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    
                                    <TextBlock Grid.Column="0" Text="Start:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="Bold"/>
                                    <DatePicker Grid.Column="1" x:Name="StartDatePicker" Height="25"/>
                                    
                                    <TextBlock Grid.Column="3" Text="End:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="Bold"/>
                                    <DatePicker Grid.Column="4" x:Name="EndDatePicker" Height="25"/>
                                </Grid>
                            </StackPanel>
                        </StackPanel>
                    </GroupBox>
                    
                    <Border Grid.Row="2" Background="#fff3cd" BorderBrush="#ffc107" BorderThickness="1" Padding="10" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock FontWeight="Bold" Foreground="#856404" Margin="0,0,0,5">
                                <Run Text="&#x24D8;"/> Search Tips:
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#856404">
                                <Run Text="&#x2022;"/> At least one search criteria is required (sender, recipient, or message ID)
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#856404">
                                <Run Text="&#x2022;"/> Date range is limited to 10 days maximum
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#856404">
                                <Run Text="&#x2022;"/> Large searches may take several minutes to complete
                            </TextBlock>
                        </StackPanel>
                    </Border>
                    
                    <StackPanel Grid.Row="7" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
                        <Button x:Name="SearchButton" Content="Search Messages" Width="140" Height="35" Margin="0,0,10,0" Background="#007bff" Foreground="White" FontWeight="Bold"/>
                        <Button x:Name="ClearButton" Content="Clear" Width="80" Height="35" Margin="0,0,10,0" Background="#6c757d" Foreground="White"/>
                        <Button x:Name="SearchCloseButton" Content="Close" Width="80" Height="35" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Results" x:Name="ResultsTab">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Grid.Row="0" Margin="0,0,0,10">
                        <TextBlock x:Name="ResultsCountText" Text="No results yet. Use the Search tab to find messages." FontWeight="Bold" Margin="0,0,0,10"/>
                        <TextBlock x:Name="ResultsInfoText" Text="" FontSize="11" Foreground="#666" TextWrapping="Wrap"/>
                    </StackPanel>
                    
                    <Border Grid.Row="1" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15">
                        <DataGrid x:Name="ResultsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single" AlternatingRowBackground="#f8f9fa">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Received" Binding="{Binding Received}" Width="140"/>
                                <DataGridTextColumn Header="Sender" Binding="{Binding SenderAddress}" Width="200"/>
                                <DataGridTextColumn Header="Recipient" Binding="{Binding RecipientAddress}" Width="200"/>
                                <DataGridTextColumn Header="Subject" Binding="{Binding Subject}" Width="250"/>
                                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100"/>
                                <DataGridTextColumn Header="Size (KB)" Binding="{Binding Size}" Width="80"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Border>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="ViewDetailsButton" Content="View Details" Width="120" Height="32" Margin="0,0,10,0" Background="#17a2b8" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="ExportResultsButton" Content="Export to Excel" Width="130" Height="32" Margin="0,0,10,0" Background="#28a745" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="ResultsCloseButton" Content="Close" Width="80" Height="32" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Message Details" x:Name="DetailsTab">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="Message Information" Padding="10" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Message ID:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="0" Grid.Column="1" x:Name="DetailMessageIdBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="1" Grid.Column="0" Text="Subject:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="1" Grid.Column="1" x:Name="DetailSubjectBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="2" Grid.Column="0" Text="From:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="2" Grid.Column="1" x:Name="DetailFromBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="3" Grid.Column="0" Text="To:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="3" Grid.Column="1" x:Name="DetailToBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="4" Grid.Column="0" Text="Received:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="4" Grid.Column="1" x:Name="DetailReceivedBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="5" Grid.Column="0" Text="Status:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="5" Grid.Column="1" x:Name="DetailStatusBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                        </Grid>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="1" Header="Delivery Events" Padding="10" Margin="0,0,0,15">
                        <Border BorderBrush="#dee2e6" BorderThickness="1">
                            <DataGrid x:Name="DetailsGrid" AutoGenerateColumns="False" IsReadOnly="True" AlternatingRowBackground="#f8f9fa">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Timestamp" Binding="{Binding Timestamp}" Width="140"/>
                                    <DataGridTextColumn Header="Event" Binding="{Binding Event}" Width="150"/>
                                    <DataGridTextColumn Header="Detail" Binding="{Binding Detail}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Border>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="CopyMessageIdButton" Content="Copy Message ID" Width="140" Height="32" Margin="0,0,10,0" Background="#6c757d" Foreground="White" IsEnabled="False"/>
                        <Button x:Name="DetailsCloseButton" Content="Close" Width="80" Height="32" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@
    
    $traceReader = New-Object System.Xml.XmlNodeReader $MessageTraceXAML
    $TraceWindow = [Windows.Markup.XamlReader]::Load($traceReader)
    $TraceWindow.Owner = $syncHash.Window
    
    # Search Tab Controls
    $SenderEmailBox = $TraceWindow.FindName("SenderEmailBox")
    $RecipientEmailBox = $TraceWindow.FindName("RecipientEmailBox")
    $SubjectBox = $TraceWindow.FindName("SubjectBox")
    $MessageIdBox = $TraceWindow.FindName("MessageIdBox")
    $StatusCombo = $TraceWindow.FindName("StatusCombo")
    $PageSizeCombo = $TraceWindow.FindName("PageSizeCombo")
    
    $Last24HoursRadio = $TraceWindow.FindName("Last24HoursRadio")
    $Last7DaysRadio = $TraceWindow.FindName("Last7DaysRadio")
    $CustomRangeRadio = $TraceWindow.FindName("CustomRangeRadio")
    $CustomDatePanel = $TraceWindow.FindName("CustomDatePanel")
    $StartDatePicker = $TraceWindow.FindName("StartDatePicker")
    $EndDatePicker = $TraceWindow.FindName("EndDatePicker")
    
    $SearchButton = $TraceWindow.FindName("SearchButton")
    $ClearButton = $TraceWindow.FindName("ClearButton")
    $SearchCloseButton = $TraceWindow.FindName("SearchCloseButton")
    
    # Results Tab Controls
    $ResultsTab = $TraceWindow.FindName("ResultsTab")
    $ResultsCountText = $TraceWindow.FindName("ResultsCountText")
    $ResultsInfoText = $TraceWindow.FindName("ResultsInfoText")
    $ResultsGrid = $TraceWindow.FindName("ResultsGrid")
    $ViewDetailsButton = $TraceWindow.FindName("ViewDetailsButton")
    $ExportResultsButton = $TraceWindow.FindName("ExportResultsButton")
    $ResultsCloseButton = $TraceWindow.FindName("ResultsCloseButton")
    
    # Details Tab Controls
    $DetailsTab = $TraceWindow.FindName("DetailsTab")
    $DetailMessageIdBox = $TraceWindow.FindName("DetailMessageIdBox")
    $DetailSubjectBox = $TraceWindow.FindName("DetailSubjectBox")
    $DetailFromBox = $TraceWindow.FindName("DetailFromBox")
    $DetailToBox = $TraceWindow.FindName("DetailToBox")
    $DetailReceivedBox = $TraceWindow.FindName("DetailReceivedBox")
    $DetailStatusBox = $TraceWindow.FindName("DetailStatusBox")
    $DetailsGrid = $TraceWindow.FindName("DetailsGrid")
    $CopyMessageIdButton = $TraceWindow.FindName("CopyMessageIdButton")
    $DetailsCloseButton = $TraceWindow.FindName("DetailsCloseButton")
    
    # Initialize date pickers
    $StartDatePicker.SelectedDate = (Get-Date).AddDays(-7)
    $EndDatePicker.SelectedDate = Get-Date
    
    # Store current results
    $script:currentTraceResults = $null
    $script:selectedMessage = $null
    
    # Enable/disable custom date panel
    $CustomRangeRadio.Add_Checked({ $CustomDatePanel.IsEnabled = $true })
    $Last24HoursRadio.Add_Checked({ $CustomDatePanel.IsEnabled = $false })
    $Last7DaysRadio.Add_Checked({ $CustomDatePanel.IsEnabled = $false })
    
    # Clear button
    $ClearButton.Add_Click({
        $SenderEmailBox.Clear()
        $RecipientEmailBox.Clear()
        $SubjectBox.Clear()
        $MessageIdBox.Clear()
        $StatusCombo.SelectedIndex = 0
        $PageSizeCombo.SelectedIndex = 1
        $Last24HoursRadio.IsChecked = $true
        Write-Log "Message trace search criteria cleared"
    })
    
    # Search button
    $SearchButton.Add_Click({
        $sender = $SenderEmailBox.Text.Trim()
        $recipient = $RecipientEmailBox.Text.Trim()
        $subject = $SubjectBox.Text.Trim()
        $messageId = $MessageIdBox.Text.Trim()
        
        # Validation
        if ([string]::IsNullOrWhiteSpace($sender) -and 
            [string]::IsNullOrWhiteSpace($recipient) -and 
            [string]::IsNullOrWhiteSpace($messageId)) {
            [System.Windows.MessageBox]::Show(
                "Please enter at least one search criteria:`n`nÃ¢â‚¬Â¢ Sender Email`nÃ¢â‚¬Â¢ Recipient Email`nÃ¢â‚¬Â¢ Message ID",
                "Validation",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return
        }
        
        # Get date range
        $startDate = $null
        $endDate = Get-Date
        
        if ($Last24HoursRadio.IsChecked) {
            $startDate = (Get-Date).AddHours(-24)
        } elseif ($Last7DaysRadio.IsChecked) {
            $startDate = (Get-Date).AddDays(-7)
        } else {
            if (-not $StartDatePicker.SelectedDate -or -not $EndDatePicker.SelectedDate) {
                [System.Windows.MessageBox]::Show("Please select both start and end dates", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            
            $startDate = $StartDatePicker.SelectedDate
            $endDate = $EndDatePicker.SelectedDate.AddDays(1).AddSeconds(-1)
            
            # Check 10-day limit
            $daysDiff = ($endDate - $startDate).TotalDays
            if ($daysDiff -gt 10) {
                [System.Windows.MessageBox]::Show("Date range cannot exceed 10 days", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
        }
        
        try {
            $SearchButton.IsEnabled = $false
            $ResultsCountText.Text = "Searching... This may take a few minutes for large result sets."
            $ResultsInfoText.Text = ""
            $ResultsGrid.ItemsSource = $null
            
            Write-Log "Starting message trace search..."
            Write-Log "  Sender: $sender"
            Write-Log "  Recipient: $recipient"
            Write-Log "  Start: $($startDate.ToString('yyyy-MM-dd HH:mm'))"
            Write-Log "  End: $($endDate.ToString('yyyy-MM-dd HH:mm'))"
            
            # Build parameters
            $traceParams = @{
                StartDate = $startDate
                EndDate = $endDate
                PageSize = [int]$PageSizeCombo.SelectedItem.Tag
            }
            
            if (-not [string]::IsNullOrWhiteSpace($sender)) {
                $traceParams.SenderAddress = $sender
            }
            if (-not [string]::IsNullOrWhiteSpace($recipient)) {
                $traceParams.RecipientAddress = $recipient
            }
            if (-not [string]::IsNullOrWhiteSpace($messageId)) {
                $traceParams.MessageId = $messageId
            }
            
            $status = $StatusCombo.SelectedItem.Tag
            if ($status -ne "All") {
                $traceParams.Status = $status
            }
            
            # Execute search
            $results = @(Get-MessageTrace @traceParams -ErrorAction Stop)
            
            Write-Log "Found $($results.Count) messages"
            
            if ($results.Count -eq 0) {
                $ResultsCountText.Text = "No messages found matching the search criteria"
                $ResultsInfoText.Text = "Try adjusting your search parameters or expanding the date range"
                $script:currentTraceResults = $null
                $ExportResultsButton.IsEnabled = $false
            } else {
                # Format results for display
                $displayResults = @()
                foreach ($msg in $results) {
                    $sizeKB = if ($msg.Size) { [math]::Round($msg.Size / 1KB, 2) } else { 0 }
                    
                    $displayResults += [PSCustomObject]@{
                        Received = $msg.Received.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
                        SenderAddress = $msg.SenderAddress
                        RecipientAddress = $msg.RecipientAddress
                        Subject = if ($msg.Subject) { $msg.Subject } else { "(No Subject)" }
                        Status = $msg.Status
                        Size = $sizeKB
                        MessageId = $msg.MessageId
                        MessageTraceId = $msg.MessageTraceId
                        FromIP = $msg.FromIP
                        ToIP = $msg.ToIP
                    }
                }
                
                $ResultsGrid.ItemsSource = $displayResults
                $script:currentTraceResults = $displayResults
                
                $ResultsCountText.Text = "Found $($results.Count) message(s)"
                $ResultsInfoText.Text = "Double-click a message or use 'View Details' to see full delivery information"
                $ExportResultsButton.IsEnabled = $true
                
                # Switch to Results tab
                $ResultsTab.IsSelected = $true
            }
            
        } catch {
            Write-Log "Error during message trace: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error searching messages:`n`n$($_.Exception.Message)",
                "Search Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            $ResultsCountText.Text = "Search failed"
            $ResultsInfoText.Text = $_.Exception.Message
        } finally {
            $SearchButton.IsEnabled = $true
        }
    })
    
    # Results grid selection changed
    $ResultsGrid.Add_SelectionChanged({
        $ViewDetailsButton.IsEnabled = ($null -ne $ResultsGrid.SelectedItem)
    })
    
    # Double-click to view details
    $ResultsGrid.Add_MouseDoubleClick({
        if ($ResultsGrid.SelectedItem) {
            $ViewDetailsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    # View Details button
    $ViewDetailsButton.Add_Click({
        $selected = $ResultsGrid.SelectedItem
        if ($null -eq $selected) { return }
        
        try {
            $ViewDetailsButton.IsEnabled = $false
            Write-Log "Loading message details for: $($selected.MessageId)"
            
            # Get message trace details
            $messageId = $selected.MessageId
            $recipient = $selected.RecipientAddress
            
            $details = @(Get-MessageTraceDetail -MessageTraceId $selected.MessageTraceId -RecipientAddress $recipient -ErrorAction Stop)
            
            Write-Log "Retrieved $($details.Count) detail events"
            
            # Populate message info
            $DetailMessageIdBox.Text = $selected.MessageId
            $DetailSubjectBox.Text = $selected.Subject
            $DetailFromBox.Text = $selected.SenderAddress
            $DetailToBox.Text = $selected.RecipientAddress
            $DetailReceivedBox.Text = $selected.Received
            $DetailStatusBox.Text = $selected.Status
            
            # Populate events
            $eventList = @()
            foreach ($detail in $details) {
                $eventList += [PSCustomObject]@{
                    Timestamp = $detail.Date.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
                    Event = $detail.Event
                    Detail = $detail.Detail
                }
            }
            
            $DetailsGrid.ItemsSource = $eventList | Sort-Object Timestamp
            $CopyMessageIdButton.IsEnabled = $true
            
            # Switch to Details tab
            $DetailsTab.IsSelected = $true
            
        } catch {
            Write-Log "Error loading message details: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error loading message details:`n`n$($_.Exception.Message)",
                "Details Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $ViewDetailsButton.IsEnabled = $true
        }
    })
    
    # Export to Excel
    $ExportResultsButton.Add_Click({
        if ($null -eq $script:currentTraceResults -or $script:currentTraceResults.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No results to export", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed. Cannot export to Excel.`n`nPlease run: Install-Module -Name ImportExcel", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Message Trace Report"
            $saveDialog.FileName = "MessageTrace_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $ExportResultsButton.IsEnabled = $false
                Write-Log "Exporting message trace results to Excel: $excelPath"
                
                $exportData = @()
                foreach ($msg in $script:currentTraceResults) {
                    $exportData += [PSCustomObject]@{
                        'Received' = $msg.Received
                        'Sender' = $msg.SenderAddress
                        'Recipient' = $msg.RecipientAddress
                        'Subject' = $msg.Subject
                        'Status' = $msg.Status
                        'Size (KB)' = $msg.Size
                        'Message ID' = $msg.MessageId
                        'From IP' = $msg.FromIP
                        'To IP' = $msg.ToIP
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "Message Trace" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "MessageTrace"
                
                Write-Log "Successfully exported $($exportData.Count) messages to Excel"
                
                [System.Windows.MessageBox]::Show(
                    "Message trace results exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
                
            } else {
                Write-Log "Export cancelled by user"
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error exporting to Excel:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $ExportResultsButton.IsEnabled = $true
        }
    })
    
    # Copy Message ID button
    $CopyMessageIdButton.Add_Click({
        if (-not [string]::IsNullOrWhiteSpace($DetailMessageIdBox.Text)) {
            [System.Windows.Forms.Clipboard]::SetText($DetailMessageIdBox.Text)
            [System.Windows.MessageBox]::Show("Message ID copied to clipboard!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
    })
    
    # Close buttons
    $SearchCloseButton.Add_Click({ $TraceWindow.Close() })
    $ResultsCloseButton.Add_Click({ $TraceWindow.Close() })
    $DetailsCloseButton.Add_Click({ $TraceWindow.Close() })
    
    # Enter key support
    $SenderEmailBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $SearchButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $RecipientEmailBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $SearchButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $MessageIdBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $SearchButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $TraceWindow.ShowDialog() | Out-Null
})

# Future Feature Placeholders
$syncHash.SendOnBehalfButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Send on Behalf Permissions feature is planned for version 2.10.0",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

$syncHash.ForwardingButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Email Forwarding Management feature is planned for version 2.11.0",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

$syncHash.ResourceMailboxButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Resource Mailbox Management feature is planned for version 2.13.0",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

$syncHash.DistributionGroupButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Distribution List Management feature is planned for version 2.15.0",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})



$syncHash.LitigationHoldButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Litigation Hold Management feature is planned for version 3.0.0",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

$syncHash.MailboxStatsButton.Add_Click({
    # Check if connected
    $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
        $result = [System.Windows.MessageBox]::Show(
            "You are not connected to Exchange Online.`n`nWould you like to connect now?",
            "Connection Required",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $syncHash.ConnectButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            
            # Wait for connection to complete
            Start-Sleep -Seconds 2
            $connInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
            
            if ($null -eq $connInfo -or $connInfo.State -ne 'Connected') {
                Write-Log "Connection was not established. Please try again."
                return
            }
        } else {
            return
        }
    }
    
    Write-Log "Opening Mailbox Size & Quota Report window..."
    
    [xml]$MailboxStatsXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Mailbox Size &amp; Quota Report" 
        Height="700" 
        Width="1000" 
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
    <Grid>
        <TabControl Margin="10">
            <TabItem Header="Scan">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="Scan Options" Padding="15" Margin="0,0,0,15">
                        <StackPanel>
                            <RadioButton x:Name="ScanAllRadio" Content="Scan All Mailboxes" GroupName="ScanType" Margin="0,5" IsChecked="True" FontSize="13"/>
                            <RadioButton x:Name="ScanSpecificRadio" Content="Scan Specific Mailbox" GroupName="ScanType" Margin="0,5" FontSize="13"/>
                            
                            <StackPanel x:Name="SpecificMailboxPanel" Margin="25,10,0,0" IsEnabled="False">
                                <TextBlock Text="Enter email address:" FontWeight="Bold" Margin="0,0,0,5"/>
                                <TextBox x:Name="SpecificMailboxBox" MinHeight="30" Padding="5,8,5,5" VerticalContentAlignment="Center" Width="300" HorizontalAlignment="Left"/>
                            </StackPanel>
                        </StackPanel>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="1" Header="Mailbox Type Filter" Padding="15" Margin="0,0,0,15">
                        <StackPanel Orientation="Horizontal">
                            <CheckBox x:Name="IncludeUserMailboxes" Content="User Mailboxes" Margin="0,0,20,0" IsChecked="True" FontSize="13"/>
                            <CheckBox x:Name="IncludeSharedMailboxes" Content="Shared Mailboxes" Margin="0,0,20,0" IsChecked="True" FontSize="13"/>
                            <CheckBox x:Name="IncludeArchives" Content="Archive Mailboxes" IsChecked="False" FontSize="13"/>
                        </StackPanel>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="2" Header="Progress" Padding="15" Margin="0,0,0,15" x:Name="ProgressGroup" Visibility="Collapsed">
                        <StackPanel>
                            <TextBlock x:Name="ProgressText" Text="Ready to scan..." Margin="0,0,0,10" FontWeight="Bold"/>
                            <ProgressBar x:Name="ScanProgressBar" Height="25" Minimum="0" Maximum="100" Value="0"/>
                            <TextBlock x:Name="ProgressDetailText" Text="" Margin="0,10,0,0" FontSize="11" Foreground="#666" TextWrapping="Wrap"/>
                        </StackPanel>
                    </GroupBox>
                    
                    <Border Grid.Row="3" Background="#d1ecf1" BorderBrush="#17a2b8" BorderThickness="1" Padding="15" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock FontWeight="Bold" Foreground="#0c5460" Margin="0,0,0,5" FontSize="13">
                                <Run Text="&#x24D8;"/> Scan Information:
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#0c5460" Margin="0,0,0,3">
                                <Run Text="&#x2022;"/> Scanning all mailboxes may take several minutes depending on organization size
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#0c5460" Margin="0,0,0,3">
                                <Run Text="&#x2022;"/> Results will show mailbox size, item count, quota, and percentage used
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#0c5460" Margin="0,0,0,3">
                                <Run Text="&#x2022;"/> Mailboxes over 80% quota will be highlighted in red
                            </TextBlock>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#0c5460">
                                <Run Text="&#x2022;"/> Export to Excel for detailed analysis and reporting
                            </TextBlock>
                        </StackPanel>
                    </Border>
                    
                    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="StartScanButton" Content="Start Scan" Width="120" Height="35" Margin="0,0,10,0" Background="#28a745" Foreground="White" FontWeight="Bold"/>
                        <Button x:Name="StopScanButton" Content="Stop Scan" Width="100" Height="35" Margin="0,0,10,0" Background="#dc3545" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="ScanCloseButton" Content="Close" Width="80" Height="35" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Results" x:Name="ResultsTab">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="Summary Statistics" Padding="15" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            
                            <StackPanel Grid.Column="0" Margin="0,0,15,0">
                                <TextBlock Text="Total Mailboxes" FontSize="11" Foreground="#666" Margin="0,0,0,5"/>
                                <TextBlock x:Name="TotalMailboxesText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#007bff"/>
                            </StackPanel>
                            
                            <StackPanel Grid.Column="1" Margin="0,0,15,0">
                                <TextBlock Text="Total Size" FontSize="11" Foreground="#666" Margin="0,0,0,5"/>
                                <TextBlock x:Name="TotalSizeText" Text="0 GB" FontSize="24" FontWeight="Bold" Foreground="#28a745"/>
                            </StackPanel>
                            
                            <StackPanel Grid.Column="2" Margin="0,0,15,0">
                                <TextBlock Text="Average Size" FontSize="11" Foreground="#666" Margin="0,0,0,5"/>
                                <TextBlock x:Name="AverageSizeText" Text="0 GB" FontSize="24" FontWeight="Bold" Foreground="#17a2b8"/>
                            </StackPanel>
                            
                            <StackPanel Grid.Column="3">
                                <TextBlock Text="Over 80% Quota" FontSize="11" Foreground="#666" Margin="0,0,0,5"/>
                                <TextBlock x:Name="OverQuotaText" Text="0" FontSize="24" FontWeight="Bold" Foreground="#dc3545"/>
                            </StackPanel>
                        </Grid>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="1" Margin="0,0,0,10">
                        <DockPanel>
                            <TextBlock DockPanel.Dock="Left" x:Name="ResultsCountText" Text="No results yet. Use the Scan tab to begin." FontWeight="Bold"/>
                            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" HorizontalAlignment="Right">
                                <TextBlock Text="Sort by:" VerticalAlignment="Center" Margin="0,0,10,0" FontSize="11"/>
                                <ComboBox x:Name="SortCombo" Width="150" Height="25" SelectedIndex="0">
                                    <ComboBoxItem Content="Display Name" Tag="DisplayName"/>
                                    <ComboBoxItem Content="Size (Largest First)" Tag="SizeDesc"/>
                                    <ComboBoxItem Content="Size (Smallest First)" Tag="SizeAsc"/>
                                    <ComboBoxItem Content="% Used (Highest First)" Tag="PercentDesc"/>
                                    <ComboBoxItem Content="% Used (Lowest First)" Tag="PercentAsc"/>
                                </ComboBox>
                            </StackPanel>
                        </DockPanel>
                    </StackPanel>
                    
                    <Border Grid.Row="2" BorderBrush="#dee2e6" BorderThickness="1" Margin="0,0,0,15">
                        <DataGrid x:Name="ResultsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single" AlternatingRowBackground="#f8f9fa">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="200"/>
                                <DataGridTextColumn Header="Email" Binding="{Binding PrimarySmtpAddress}" Width="220"/>
                                <DataGridTextColumn Header="Type" Binding="{Binding MailboxType}" Width="100"/>
                                <DataGridTextColumn Header="Size (GB)" Binding="{Binding SizeGB}" Width="90"/>
                                <DataGridTextColumn Header="Items" Binding="{Binding ItemCount}" Width="80"/>
                                <DataGridTextColumn Header="Quota (GB)" Binding="{Binding QuotaGB}" Width="90"/>
                                <DataGridTextColumn Header="% Used" Binding="{Binding PercentUsed}" Width="80"/>
                                <DataGridTextColumn Header="Status" Binding="{Binding QuotaStatus}" Width="100"/>
                            </DataGrid.Columns>
                            <DataGrid.RowStyle>
                                <Style TargetType="DataGridRow">
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding QuotaWarning}" Value="Critical">
                                            <Setter Property="Background" Value="#f8d7da"/>
                                            <Setter Property="Foreground" Value="#721c24"/>
                                        </DataTrigger>
                                        <DataTrigger Binding="{Binding QuotaWarning}" Value="Warning">
                                            <Setter Property="Background" Value="#fff3cd"/>
                                            <Setter Property="Foreground" Value="#856404"/>
                                        </DataTrigger>
                                    </Style.Triggers>
                                </Style>
                            </DataGrid.RowStyle>
                        </DataGrid>
                    </Border>
                    
                    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="ViewFoldersButton" Content="View Folders" Width="120" Height="32" Margin="0,0,10,0" Background="#17a2b8" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="ExportResultsButton" Content="Export to Excel" Width="130" Height="32" Margin="0,0,10,0" Background="#28a745" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="ResultsCloseButton" Content="Close" Width="80" Height="32" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Folder Details" x:Name="DetailsTab">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <GroupBox Grid.Row="0" Header="Mailbox Information" Padding="10" Margin="0,0,0,15">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="20"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Display Name:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="0" Grid.Column="1" x:Name="DetailDisplayNameBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="0" Grid.Column="3" Text="Email:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="0" Grid.Column="4" x:Name="DetailEmailBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="1" Grid.Column="0" Text="Total Size:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="1" Grid.Column="1" x:Name="DetailSizeBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="1" Grid.Column="3" Text="Total Items:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="1" Grid.Column="4" x:Name="DetailItemsBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="2" Grid.Column="0" Text="Quota:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="2" Grid.Column="1" x:Name="DetailQuotaBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                            
                            <TextBlock Grid.Row="2" Grid.Column="3" Text="% Used:" FontWeight="Bold" Margin="0,5,10,5"/>
                            <TextBox Grid.Row="2" Grid.Column="4" x:Name="DetailPercentBox" IsReadOnly="True" Margin="0,5" Padding="5"/>
                        </Grid>
                    </GroupBox>
                    
                    <GroupBox Grid.Row="1" Header="Folder Breakdown" Padding="10" Margin="0,0,0,15">
                        <Border BorderBrush="#dee2e6" BorderThickness="1">
                            <DataGrid x:Name="FoldersGrid" AutoGenerateColumns="False" IsReadOnly="True" AlternatingRowBackground="#f8f9fa">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Folder Name" Binding="{Binding FolderName}" Width="300"/>
                                    <DataGridTextColumn Header="Items" Binding="{Binding ItemCount}" Width="100"/>
                                    <DataGridTextColumn Header="Size (MB)" Binding="{Binding SizeMB}" Width="120"/>
                                    <DataGridTextColumn Header="% of Total" Binding="{Binding PercentOfTotal}" Width="100"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Border>
                    </GroupBox>
                    
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="ExportFoldersButton" Content="Export Folders" Width="130" Height="32" Margin="0,0,10,0" Background="#28a745" Foreground="White" FontWeight="Bold" IsEnabled="False"/>
                        <Button x:Name="DetailsCloseButton" Content="Close" Width="80" Height="32" Background="#6c757d" Foreground="White"/>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@
    
    $statsReader = New-Object System.Xml.XmlNodeReader $MailboxStatsXAML
    $StatsWindow = [Windows.Markup.XamlReader]::Load($statsReader)
    $StatsWindow.Owner = $syncHash.Window
    
    # Scan Tab Controls
    $ScanAllRadio = $StatsWindow.FindName("ScanAllRadio")
    $ScanSpecificRadio = $StatsWindow.FindName("ScanSpecificRadio")
    $SpecificMailboxPanel = $StatsWindow.FindName("SpecificMailboxPanel")
    $SpecificMailboxBox = $StatsWindow.FindName("SpecificMailboxBox")
    
    $IncludeUserMailboxes = $StatsWindow.FindName("IncludeUserMailboxes")
    $IncludeSharedMailboxes = $StatsWindow.FindName("IncludeSharedMailboxes")
    $IncludeArchives = $StatsWindow.FindName("IncludeArchives")
    
    $ProgressGroup = $StatsWindow.FindName("ProgressGroup")
    $ProgressText = $StatsWindow.FindName("ProgressText")
    $ScanProgressBar = $StatsWindow.FindName("ScanProgressBar")
    $ProgressDetailText = $StatsWindow.FindName("ProgressDetailText")
    
    $StartScanButton = $StatsWindow.FindName("StartScanButton")
    $StopScanButton = $StatsWindow.FindName("StopScanButton")
    $ScanCloseButton = $StatsWindow.FindName("ScanCloseButton")
    
    # Results Tab Controls
    $ResultsTab = $StatsWindow.FindName("ResultsTab")
    $TotalMailboxesText = $StatsWindow.FindName("TotalMailboxesText")
    $TotalSizeText = $StatsWindow.FindName("TotalSizeText")
    $AverageSizeText = $StatsWindow.FindName("AverageSizeText")
    $OverQuotaText = $StatsWindow.FindName("OverQuotaText")
    $ResultsCountText = $StatsWindow.FindName("ResultsCountText")
    $SortCombo = $StatsWindow.FindName("SortCombo")
    $ResultsGrid = $StatsWindow.FindName("ResultsGrid")
    $ViewFoldersButton = $StatsWindow.FindName("ViewFoldersButton")
    $ExportResultsButton = $StatsWindow.FindName("ExportResultsButton")
    $ResultsCloseButton = $StatsWindow.FindName("ResultsCloseButton")
    
    # Details Tab Controls
    $DetailsTab = $StatsWindow.FindName("DetailsTab")
    $DetailDisplayNameBox = $StatsWindow.FindName("DetailDisplayNameBox")
    $DetailEmailBox = $StatsWindow.FindName("DetailEmailBox")
    $DetailSizeBox = $StatsWindow.FindName("DetailSizeBox")
    $DetailItemsBox = $StatsWindow.FindName("DetailItemsBox")
    $DetailQuotaBox = $StatsWindow.FindName("DetailQuotaBox")
    $DetailPercentBox = $StatsWindow.FindName("DetailPercentBox")
    $FoldersGrid = $StatsWindow.FindName("FoldersGrid")
    $ExportFoldersButton = $StatsWindow.FindName("ExportFoldersButton")
    $DetailsCloseButton = $StatsWindow.FindName("DetailsCloseButton")
    
    # Store current results
    $script:currentMailboxStats = $null
    $script:shouldStopScan = $false
    
    # Enable/disable specific mailbox panel
    $ScanAllRadio.Add_Checked({ $SpecificMailboxPanel.IsEnabled = $false })
    $ScanSpecificRadio.Add_Checked({ $SpecificMailboxPanel.IsEnabled = $true; $SpecificMailboxBox.Focus() })
    
    # Start Scan button
    $StartScanButton.Add_Click({
        $scanAll = $ScanAllRadio.IsChecked
        $specificMailbox = $SpecificMailboxBox.Text.Trim()
        
        # Validation
        if (-not $scanAll -and [string]::IsNullOrWhiteSpace($specificMailbox)) {
            [System.Windows.MessageBox]::Show("Please enter a mailbox email address", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }

        if ($scanAll) {
            # Only check mailbox type filters when scanning all
            if (-not $IncludeUserMailboxes.IsChecked -and -not $IncludeSharedMailboxes.IsChecked -and -not $IncludeArchives.IsChecked) {
                [System.Windows.MessageBox]::Show("Please select at least one mailbox type to scan", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
        }
        
        if (-not $IncludeUserMailboxes.IsChecked -and -not $IncludeSharedMailboxes.IsChecked -and -not $IncludeArchives.IsChecked) {
            [System.Windows.MessageBox]::Show("Please select at least one mailbox type to scan", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $script:shouldStopScan = $false
            $StartScanButton.IsEnabled = $false
            $StopScanButton.IsEnabled = $true
            $ProgressGroup.Visibility = [System.Windows.Visibility]::Visible
            $ScanProgressBar.Value = 0
            
            Write-Log "Starting mailbox size scan..."
            
            # Get mailboxes to scan
            $mailboxesToScan = @()
            
            if ($scanAll) {
                $ProgressText.Text = "Retrieving mailbox list..."
                $ProgressDetailText.Text = "This may take a moment for large organizations..."
                
                Write-Log "Retrieving all mailboxes..."
                
                $recipientTypeDetails = @()
                if ($IncludeUserMailboxes.IsChecked) { $recipientTypeDetails += "UserMailbox" }
                if ($IncludeSharedMailboxes.IsChecked) { $recipientTypeDetails += "SharedMailbox" }
                
                if ($recipientTypeDetails.Count -gt 0) {
                    $mailboxesToScan += @(Get-Mailbox -RecipientTypeDetails $recipientTypeDetails -ResultSize Unlimited -ErrorAction Stop)
                }
                
                if ($IncludeArchives.IsChecked) {
                    Write-Log "Including archive mailboxes..."
                    $archiveMailboxes = @(Get-Mailbox -Archive -ResultSize Unlimited -ErrorAction Stop)
                    $mailboxesToScan += $archiveMailboxes
                }
                
                Write-Log "Found $($mailboxesToScan.Count) mailboxes to scan"
                
            } else {
                $ProgressText.Text = "Retrieving mailbox information..."
                Write-Log "Retrieving specific mailbox: $specificMailbox"
                
                $mailboxesToScan = @(Get-Mailbox -Identity $specificMailbox -ErrorAction Stop)
            }
            
            if ($mailboxesToScan.Count -eq 0) {
                throw "No mailboxes found matching the criteria"
            }
            
            $ProgressText.Text = "Scanning $($mailboxesToScan.Count) mailbox(es)..."
            $ProgressDetailText.Text = "Processing mailbox statistics..."
            
            # Collect stats
            $results = @()
            $processedCount = 0
            $totalCount = $mailboxesToScan.Count
            
            foreach ($mbx in $mailboxesToScan) {
                if ($script:shouldStopScan) {
                    Write-Log "Scan stopped by user"
                    break
                }
                
                $processedCount++
                $percentComplete = [Math]::Round(($processedCount / $totalCount) * 100)
                $ScanProgressBar.Value = $percentComplete
                $ProgressDetailText.Text = "Processing $($mbx.DisplayName) ($processedCount of $totalCount)"
                
                # Force UI update
                [System.Windows.Forms.Application]::DoEvents()
                
                try {
                    Write-Log "Processing: $($mbx.DisplayName)"
                    
                    # Get mailbox statistics
                    $stats = Get-MailboxStatistics -Identity $mbx.Identity -ErrorAction Stop
                    
                    # Calculate size in GB
                    $sizeBytes = 0
                    if ($stats.TotalItemSize) {
                        $sizeString = $stats.TotalItemSize.ToString()
                        if ($sizeString -match '([0-9,]+)\s*bytes') {
                            $sizeBytes = [int64]($matches[1] -replace ',', '')
                        }
                    }
                    $sizeGB = [Math]::Round($sizeBytes / 1GB, 2)
                    
                    # Get quota information
                    $quotaGB = 0
                    $quotaStatus = "N/A"
                    $percentUsed = 0
                    $quotaWarning = "Normal"
                    
                    # Safely get ProhibitSendQuota
                    $prohibitQuota = $null
                    try {
                        $prohibitQuota = $mbx.ProhibitSendQuota
                    } catch {
                        Write-Log "Could not retrieve ProhibitSendQuota for $($mbx.DisplayName)"
                    }
                    
                    if ($null -ne $prohibitQuota -and $prohibitQuota -ne "Unlimited" -and $prohibitQuota.ToString() -ne "Unlimited") {
                        $quotaString = $prohibitQuota.ToString()
                        if ($quotaString -match '([0-9.]+)\s*GB') {
                            $quotaGB = [Math]::Round([decimal]$matches[1], 2)
                        } elseif ($quotaString -match '([0-9.]+)\s*MB') {
                            $quotaGB = [Math]::Round([decimal]$matches[1] / 1024, 2)
                        }
                        
                        if ($quotaGB -gt 0) {
                            $percentUsed = [Math]::Round(($sizeGB / $quotaGB) * 100, 1)
                            
                            if ($percentUsed -ge 95) {
                                $quotaStatus = "Critical"
                                $quotaWarning = "Critical"
                            } elseif ($percentUsed -ge 80) {
                                $quotaStatus = "Warning"
                                $quotaWarning = "Warning"
                            } else {
                                $quotaStatus = "Normal"
                            }
                        }
                    } else {
                        $quotaGB = "Unlimited"
                        $quotaStatus = "N/A"
                    }
                    
                    # Determine mailbox type
                    $mailboxType = switch ($mbx.RecipientTypeDetails) {
                        "UserMailbox" { "User" }
                        "SharedMailbox" { "Shared" }
                        "RoomMailbox" { "Room" }
                        "EquipmentMailbox" { "Equipment" }
                        default { if ($mbx.RecipientTypeDetails) { $mbx.RecipientTypeDetails.ToString() } else { "Unknown" } }
                    }
                    
                    # Check for archive
                    try {
                        if ($mbx.ArchiveStatus -eq "Active") {
                            $mailboxType += " (Archive)"
                        }
                    } catch {
                        # ArchiveStatus not available, skip
                    }
                    
                    # Safely get email address
                    $emailAddress = "N/A"
                    try {
                        if ($mbx.PrimarySmtpAddress) {
                            $emailAddress = $mbx.PrimarySmtpAddress.ToString()
                        } elseif ($mbx.EmailAddresses) {
                            $smtpAddr = $mbx.EmailAddresses | Where-Object { $_ -like "smtp:*" } | Select-Object -First 1
                            if ($smtpAddr) {
                                $emailAddress = $smtpAddr.ToString() -replace '^smtp:', ''
                            }
                        }
                    } catch {
                        Write-Log "Could not retrieve email address for $($mbx.DisplayName)"
                    }
                    
                    # Safely get last logon time
                    $lastLogon = "Never"
                    try {
                        if ($stats.LastLogonTime) {
                            $lastLogon = $stats.LastLogonTime.ToString("yyyy-MM-dd HH:mm")
                        }
                    } catch {
                        # LastLogonTime not available
                    }
                    
                    $results += [PSCustomObject]@{
                        DisplayName = if ($mbx.DisplayName) { $mbx.DisplayName } else { "Unknown" }
                        PrimarySmtpAddress = $emailAddress
                        MailboxType = $mailboxType
                        SizeGB = $sizeGB
                        SizeBytes = $sizeBytes
                        ItemCount = if ($stats.ItemCount) { $stats.ItemCount } else { 0 }
                        QuotaGB = $quotaGB
                        PercentUsed = if ($quotaGB -eq "Unlimited") { "N/A" } else { "$percentUsed%" }
                        PercentValue = $percentUsed
                        QuotaStatus = $quotaStatus
                        QuotaWarning = $quotaWarning
                        Identity = $mbx.Identity
                        LastLogonTime = $lastLogon
                    }
                    
                } catch {
                    Write-Log "Error processing $($mbx.DisplayName): $($_.Exception.Message)"
                    # Continue to next mailbox instead of failing entire scan
                }
            }
            
            # Check if we got any results
            if ($null -eq $results -or $results.Count -eq 0) {
                throw "No mailbox statistics could be retrieved"
            }
            
            Write-Log "Successfully scanned $($results.Count) mailboxes"
            
            # Sort results by size (largest first) - filter out any null SizeBytes first
            try {
                $results = @($results | Where-Object { $null -ne $_.SizeBytes } | Sort-Object -Property SizeBytes -Descending)
            } catch {
                Write-Log "Warning: Could not sort results, displaying unsorted"
            }
            
            # Store results
            $script:currentMailboxStats = $results
            
            # Calculate summary statistics with null protection
            $totalMailboxes = $results.Count
            
            $totalSizeGB = 0
            try {
                $sizeSum = $results | Where-Object { $null -ne $_.SizeGB } | Measure-Object -Property SizeGB -Sum
                if ($null -ne $sizeSum -and $null -ne $sizeSum.Sum) {
                    $totalSizeGB = [Math]::Round($sizeSum.Sum, 2)
                }
            } catch {
                Write-Log "Warning: Could not calculate total size"
            }
            
            $avgSizeGB = 0
            try {
                if ($totalMailboxes -gt 0) {
                    $avgSizeGB = [Math]::Round($totalSizeGB / $totalMailboxes, 2)
                }
            } catch {
                Write-Log "Warning: Could not calculate average size"
            }
            
            $overQuotaCount = 0
            try {
                $overQuotaCount = @($results | Where-Object { 
                    $null -ne $_.PercentValue -and 
                    $_.PercentValue -ge 80 -and 
                    $_.PercentValue -ne 0 
                }).Count
            } catch {
                Write-Log "Warning: Could not calculate over-quota count"
            }
            
            # Update summary statistics
            try {
                $TotalMailboxesText.Text = $totalMailboxes.ToString()
                $TotalSizeText.Text = "$totalSizeGB GB"
                $AverageSizeText.Text = "$avgSizeGB GB"
                $OverQuotaText.Text = $overQuotaCount.ToString()
            } catch {
                Write-Log "Warning: Could not update summary text: $($_.Exception.Message)"
            }
            
            # Display results in grid
            try {
                $ResultsGrid.ItemsSource = $null
                $ResultsGrid.ItemsSource = $results
                $ResultsCountText.Text = "Showing $($results.Count) mailbox(es)"
                $ExportResultsButton.IsEnabled = $true
            } catch {
                Write-Log "Error binding results to grid: $($_.Exception.Message)"
                throw "Could not display results: $($_.Exception.Message)"
            }
            
            # Update progress
            $ProgressText.Text = "Scan complete!"
            $ProgressDetailText.Text = "Found $($results.Count) mailboxes. Total size: $totalSizeGB GB"
            $ScanProgressBar.Value = 100
            
            # Switch to Results tab
            $ResultsTab.IsSelected = $true

            
        } catch {
            Write-Log "Error during scan: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error scanning mailboxes:`n`n$($_.Exception.Message)",
                "Scan Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            $ProgressText.Text = "Scan failed"
            $ProgressDetailText.Text = $_.Exception.Message
        } finally {
            $StartScanButton.IsEnabled = $true
            $StopScanButton.IsEnabled = $false
        }
    })
    
    # Stop Scan button
    $StopScanButton.Add_Click({
        $script:shouldStopScan = $true
        $StopScanButton.IsEnabled = $false
        $ProgressText.Text = "Stopping scan..."
        Write-Log "Stop scan requested"
    })
    
    # Sort combo changed
    $SortCombo.Add_SelectionChanged({
        if ($null -eq $script:currentMailboxStats) { return }
        
        $sortTag = $SortCombo.SelectedItem.Tag
        $sorted = $null
        
        switch ($sortTag) {
            "DisplayName" { $sorted = $script:currentMailboxStats | Sort-Object DisplayName }
            "SizeDesc" { $sorted = $script:currentMailboxStats | Sort-Object -Property SizeBytes -Descending }
            "SizeAsc" { $sorted = $script:currentMailboxStats | Sort-Object -Property SizeBytes }
            "PercentDesc" { $sorted = $script:currentMailboxStats | Sort-Object -Property PercentValue -Descending }
            "PercentAsc" { $sorted = $script:currentMailboxStats | Sort-Object -Property PercentValue }
        }
        
        $ResultsGrid.ItemsSource = $sorted
    })
    
    # Results grid selection changed
    $ResultsGrid.Add_SelectionChanged({
        $ViewFoldersButton.IsEnabled = ($null -ne $ResultsGrid.SelectedItem)
    })
    
    # Double-click to view folders
    $ResultsGrid.Add_MouseDoubleClick({
        if ($ResultsGrid.SelectedItem) {
            $ViewFoldersButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    # View Folders button
    $ViewFoldersButton.Add_Click({
        $selected = $ResultsGrid.SelectedItem
        if ($null -eq $selected) { return }
        
        try {
            $ViewFoldersButton.IsEnabled = $false
            Write-Log "Loading folder statistics for: $($selected.DisplayName)"
            
            # Get folder statistics
            $folderStats = @(Get-MailboxFolderStatistics -Identity $selected.Identity -ErrorAction Stop)
            
            Write-Log "Retrieved $($folderStats.Count) folders"
            
            # Populate mailbox info
            $DetailDisplayNameBox.Text = $selected.DisplayName
            $DetailEmailBox.Text = $selected.PrimarySmtpAddress
            $DetailSizeBox.Text = "$($selected.SizeGB) GB"
            $DetailItemsBox.Text = $selected.ItemCount.ToString("N0")
            $DetailQuotaBox.Text = if ($selected.QuotaGB -eq "Unlimited") { "Unlimited" } else { "$($selected.QuotaGB) GB" }
            $DetailPercentBox.Text = $selected.PercentUsed
            
            # Format folder data
            $folderList = @()
            $totalSizeBytes = $selected.SizeBytes
            
            foreach ($folder in $folderStats) {
                $folderSizeBytes = 0
                if ($folder.FolderSize) {
                    $sizeString = $folder.FolderSize.ToString()
                    if ($sizeString -match '([0-9,]+)\s*bytes') {
                        $folderSizeBytes = [int64]($matches[1] -replace ',', '')
                    }
                }
                
                $folderSizeMB = [Math]::Round($folderSizeBytes / 1MB, 2)
                $percentOfTotal = if ($totalSizeBytes -gt 0) { 
                    [Math]::Round(($folderSizeBytes / $totalSizeBytes) * 100, 1) 
                } else { 
                    0 
                }
                
                $folderList += [PSCustomObject]@{
                    FolderName = $folder.Name
                    ItemCount = $folder.ItemsInFolder
                    SizeMB = $folderSizeMB
                    SizeBytes = $folderSizeBytes
                    PercentOfTotal = "$percentOfTotal%"
                }
            }
            
            # Sort by size (largest first)
            $folderList = $folderList | Sort-Object -Property SizeBytes -Descending
            
            $FoldersGrid.ItemsSource = $folderList
            $ExportFoldersButton.IsEnabled = $true
            
            # Switch to Details tab
            $DetailsTab.IsSelected = $true
            
        } catch {
            Write-Log "Error loading folder statistics: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error loading folder statistics:`n`n$($_.Exception.Message)",
                "Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $ViewFoldersButton.IsEnabled = $true
        }
    })
    
    # Export Results to Excel
    $ExportResultsButton.Add_Click({
        if ($null -eq $script:currentMailboxStats -or $script:currentMailboxStats.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No results to export", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed. Cannot export to Excel.`n`nPlease run: Install-Module -Name ImportExcel", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Mailbox Size Report"
            $saveDialog.FileName = "MailboxSizeReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $ExportResultsButton.IsEnabled = $false
                Write-Log "Exporting mailbox statistics to Excel: $excelPath"
                
                $exportData = @()
                foreach ($mbx in $script:currentMailboxStats) {
                    $exportData += [PSCustomObject]@{
                        'Display Name' = $mbx.DisplayName
                        'Email Address' = $mbx.PrimarySmtpAddress
                        'Mailbox Type' = $mbx.MailboxType
                        'Size (GB)' = $mbx.SizeGB
                        'Item Count' = $mbx.ItemCount
                        'Quota (GB)' = $mbx.QuotaGB
                        'Percent Used' = $mbx.PercentUsed
                        'Quota Status' = $mbx.QuotaStatus
                        'Last Logon' = $mbx.LastLogonTime
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                # Export with conditional formatting
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "Mailbox Sizes" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "MailboxSizes"
                
                Write-Log "Successfully exported $($exportData.Count) mailboxes to Excel"
                
                [System.Windows.MessageBox]::Show(
                    "Mailbox size report exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
                
            } else {
                Write-Log "Export cancelled by user"
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error exporting to Excel:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $ExportResultsButton.IsEnabled = $true
        }
    })
    
    # Export Folders to Excel
    $ExportFoldersButton.Add_Click({
        $folders = $FoldersGrid.ItemsSource
        if ($null -eq $folders -or $folders.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No folder data to export", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed.", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save Folder Breakdown Report"
            $mailboxName = $DetailDisplayNameBox.Text -replace '[\\/:*?"<>|]', '_'
            $saveDialog.FileName = "FolderBreakdown_${mailboxName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $exportData = @()
                foreach ($folder in $folders) {
                    $exportData += [PSCustomObject]@{
                        'Mailbox' = $DetailDisplayNameBox.Text
                        'Email' = $DetailEmailBox.Text
                        'Folder Name' = $folder.FolderName
                        'Item Count' = $folder.ItemCount
                        'Size (MB)' = $folder.SizeMB
                        'Percent of Total' = $folder.PercentOfTotal
                        'Export Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "Folder Breakdown" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "FolderBreakdown"
                
                Write-Log "Exported folder breakdown to Excel"
                
                [System.Windows.MessageBox]::Show(
                    "Folder breakdown exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show("Error exporting: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    })
    
    # Close buttons
    $ScanCloseButton.Add_Click({ $StatsWindow.Close() })
    $ResultsCloseButton.Add_Click({ $StatsWindow.Close() })
    $DetailsCloseButton.Add_Click({ $StatsWindow.Close() })
    
    # Enter key support
    $SpecificMailboxBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $StartScanButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $StatsWindow.ShowDialog() | Out-Null
})

$syncHash.IPScannerButton.Add_Click({
    Write-Log "Opening IP Network Scanner"
    
    [xml]$ScannerXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="IP Network Scanner" 
        Height="700" 
        Width="1000" 
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#233A4A" Padding="20,15">
            <StackPanel>
                <TextBlock Text="IP Network Scanner" FontSize="20" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="Scan IP ranges to discover active devices on the network" FontSize="12" Foreground="#B0BEC5" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Main Content -->
        <Grid Grid.Row="1" Margin="20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            
            <!-- Scan Configuration -->
            <GroupBox Grid.Row="0" Header="Scan Configuration" Padding="15" Margin="0,0,0,15">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <TextBlock Grid.Column="0" Text="Start IP:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="Bold"/>
                    <TextBox x:Name="StartIPBox" Grid.Column="1" Height="30" VerticalContentAlignment="Center" Padding="5" Margin="0,0,20,0"/>
                    
                    <TextBlock Grid.Column="2" Text="End IP:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="Bold"/>
                    <TextBox x:Name="EndIPBox" Grid.Column="3" Height="30" VerticalContentAlignment="Center" Padding="5" Margin="0,0,20,0"/>
                    
                    <Button x:Name="StartScanButton" Grid.Column="4" Content="Start Scan" Width="120" Height="35" Background="#28a745" Foreground="White" FontWeight="Bold"/>
                </Grid>
            </GroupBox>
            
            <!-- Results -->
            <GroupBox Grid.Row="1" Header="Scan Results" Padding="10">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Progress and Stats -->
                    <StackPanel Grid.Row="0" Margin="0,0,0,10">
                        <TextBlock x:Name="ScanStatusText" Text="Ready to scan" FontWeight="Bold" Margin="0,0,0,5"/>
                        <ProgressBar x:Name="ScanProgressBar" Height="20" Minimum="0" Maximum="100" Value="0"/>
                        <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                            <TextBlock Text="Total Scanned: " FontWeight="Bold"/>
                            <TextBlock x:Name="TotalScannedText" Text="0" Margin="0,0,20,0"/>
                            <TextBlock Text="Online: " FontWeight="Bold" Foreground="Green"/>
                            <TextBlock x:Name="OnlineCountText" Text="0" Foreground="Green" Margin="0,0,20,0"/>
                            <TextBlock Text="Offline: " FontWeight="Bold" Foreground="Red"/>
                            <TextBlock x:Name="OfflineCountText" Text="0" Foreground="Red"/>
                        </StackPanel>
                    </StackPanel>
                    
                    <!-- Results Grid -->
                    <DataGrid Grid.Row="1" 
                            x:Name="ResultsGrid" 
                            AutoGenerateColumns="False" 
                            IsReadOnly="True"
                            SelectionMode="Extended"
                            GridLinesVisibility="All"
                            AlternatingRowBackground="#f8f9fa"
                            CanUserSortColumns="True"
                            CanUserResizeColumns="True">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="IP Address" Binding="{Binding IPAddress}" Width="150"/>
                            <DataGridTextColumn Header="Hostname" Binding="{Binding Hostname}" Width="*"/>
                            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Style.Triggers>
                                            <Trigger Property="Text" Value="Online">
                                                <Setter Property="Foreground" Value="Green"/>
                                                <Setter Property="FontWeight" Value="Bold"/>
                                            </Trigger>
                                            <Trigger Property="Text" Value="Offline">
                                                <Setter Property="Foreground" Value="Red"/>
                                            </Trigger>
                                        </Style.Triggers>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Header="MAC Address" Binding="{Binding MACAddress}" Width="150"/>
                            <DataGridTextColumn Header="Response Time (ms)" Binding="{Binding ResponseTime}" Width="150"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </GroupBox>
        </Grid>
        
        <!-- Footer -->
        <Border Grid.Row="2" Background="#f8f9fa" BorderBrush="#dee2e6" BorderThickness="0,1,0,0" Padding="15">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="ExportButton" Content="Export to Excel" Width="120" Height="30" Margin="0,0,10,0" Background="#17a2b8" Foreground="White" IsEnabled="False"/>
                <Button x:Name="CloseButton" Content="Close" Width="80" Height="30" Background="#6c757d" Foreground="White"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@
    
    $scanReader = New-Object System.Xml.XmlNodeReader $ScannerXAML
    $ScanWindow = [Windows.Markup.XamlReader]::Load($scanReader)
    $ScanWindow.Owner = $syncHash.Window
    
    # Get controls
    $StartIPBox = $ScanWindow.FindName("StartIPBox")
    $EndIPBox = $ScanWindow.FindName("EndIPBox")
    $StartScanButton = $ScanWindow.FindName("StartScanButton")
    $ScanStatusText = $ScanWindow.FindName("ScanStatusText")
    $ScanProgressBar = $ScanWindow.FindName("ScanProgressBar")
    $TotalScannedText = $ScanWindow.FindName("TotalScannedText")
    $OnlineCountText = $ScanWindow.FindName("OnlineCountText")
    $OfflineCountText = $ScanWindow.FindName("OfflineCountText")
    $ResultsGrid = $ScanWindow.FindName("ResultsGrid")
    $ExportButton = $ScanWindow.FindName("ExportButton")
    $CloseButton = $ScanWindow.FindName("CloseButton")
    
    # Function to validate IP address
    function Test-IPAddress {
        param([string]$IP)
        try {
            $null = [System.Net.IPAddress]::Parse($IP)
            return $true
        } catch {
            return $false
        }
    }
    
    # Function to convert IP to integer for range calculation
    function ConvertTo-IPInteger {
        param([string]$IP)
        $bytes = [System.Net.IPAddress]::Parse($IP).GetAddressBytes()
        [Array]::Reverse($bytes)
        return [BitConverter]::ToUInt32($bytes, 0)
    }
    
    # Function to convert integer back to IP
    function ConvertFrom-IPInteger {
        param([uint32]$Int)
        $bytes = [BitConverter]::GetBytes($Int)
        [Array]::Reverse($bytes)
        return [System.Net.IPAddress]::new($bytes).ToString()
    }
    
    # Function to get MAC address from ARP table
    function Get-MACFromARP {
        param([string]$IP)
        try {
            $arpResult = arp -a $IP 2>$null
            $macLine = $arpResult | Where-Object { $_ -match $IP }
            if ($macLine -match '([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})') {
                return $matches[0].ToUpper()
            }
        } catch {
            # Silently fail
        }
        return "N/A"
    }
    
    # Start Scan Button Click
    $StartScanButton.Add_Click({
        $startIP = $StartIPBox.Text.Trim()
        $endIP = $EndIPBox.Text.Trim()
        
        # Validate IPs
        if (-not (Test-IPAddress $startIP)) {
            [System.Windows.MessageBox]::Show("Invalid start IP address", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Test-IPAddress $endIP)) {
            [System.Windows.MessageBox]::Show("Invalid end IP address", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        Write-Log "Starting IP scan from $startIP to $endIP"
        
        # Disable scan button during scan
        $StartScanButton.IsEnabled = $false
        $ExportButton.IsEnabled = $false
        
        # Clear previous results
        $ResultsGrid.ItemsSource = $null
        $results = [System.Collections.Generic.List[object]]::new()
        
        # Convert IPs to integers for range
        $startInt = ConvertTo-IPInteger $startIP
        $endInt = ConvertTo-IPInteger $endIP
        
        if ($startInt -gt $endInt) {
            [System.Windows.MessageBox]::Show("Start IP must be less than or equal to End IP", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            $StartScanButton.IsEnabled = $true
            return
        }
        
        $totalIPs = $endInt - $startInt + 1
        
        if ($totalIPs -gt 1000) {
            $response = [System.Windows.MessageBox]::Show(
                "You are about to scan $totalIPs IP addresses. This may take a while. Continue?",
                "Large Scan Warning",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )
            if ($response -eq [System.Windows.MessageBoxResult]::No) {
                $StartScanButton.IsEnabled = $true
                return
            }
        }
        
        # Reset counters
        $onlineCount = 0
        $offlineCount = 0
        $scannedCount = 0
        
        $ScanStatusText.Text = "Scanning in progress..."
        $ScanProgressBar.Value = 0
        $TotalScannedText.Text = "0"
        $OnlineCountText.Text = "0"
        $OfflineCountText.Text = "0"
        
        # Create runspace pool for parallel scanning (50 concurrent threads)
        $runspacePool = [runspacefactory]::CreateRunspacePool(1, 50)
        $runspacePool.Open()
        $runspaces = New-Object System.Collections.ArrayList
        
        # Script block for each IP scan
        $scanScriptBlock = {
            param($IP)
            
            $result = [PSCustomObject]@{
                IPAddress = $IP
                Hostname = "N/A"
                Status = "Offline"
                MACAddress = "N/A"
                ResponseTime = "N/A"
            }
            
            try {
                # Test connection with BufferSize for faster ping
                $pingResult = Test-Connection -ComputerName $IP -Count 1 -BufferSize 32 -Quiet -ErrorAction Stop
                
                if ($pingResult) {
                    $result.Status = "Online"
                    
                    # Get response time
                    try {
                        $pingDetail = Test-Connection -ComputerName $IP -Count 1 -BufferSize 32 -ErrorAction Stop
                        if ($pingDetail.ResponseTime) {
                            $result.ResponseTime = $pingDetail.ResponseTime.ToString()
                        } elseif ($pingDetail.Latency) {
                            $result.ResponseTime = $pingDetail.Latency.ToString()
                        }
                    } catch {
                        $result.ResponseTime = "<1"
                    }
                    
                    # Try to resolve hostname
                    try {
                        $hostEntry = [System.Net.Dns]::GetHostEntry($IP)
                        $result.Hostname = $hostEntry.HostName
                    } catch {
                        $result.Hostname = "Unable to resolve"
                    }
                    
                    # Ping once more to populate ARP table
                    $null = Test-NetConnection -ComputerName $IP -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction SilentlyContinue 2>$null
                    
                    # Try to get MAC address from ARP
                    try {
                        Start-Sleep -Milliseconds 50
                        $arpResult = arp -a $IP 2>$null
                        if ($arpResult) {
                            $macLine = $arpResult | Where-Object { $_ -match $IP }
                            if ($macLine -match '([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})') {
                                $result.MACAddress = $matches[0].ToUpper()
                            }
                        }
                    } catch {
                        # Silently fail
                    }
                }
            } catch {
                # IP is offline or unreachable
            }
            
            return $result
        }
        
        # Launch all scan jobs
        for ($i = $startInt; $i -le $endInt; $i++) {
            $currentIP = ConvertFrom-IPInteger $i
            
            $powershell = [powershell]::Create().AddScript($scanScriptBlock).AddArgument($currentIP)
            $powershell.RunspacePool = $runspacePool
            
            [void]$runspaces.Add([PSCustomObject]@{
                Pipe = $powershell
                Status = $powershell.BeginInvoke()
            })
        }
        
        # Monitor progress and collect results
        $completed = 0
        while ($runspaces.Status.IsCompleted -contains $false) {
            $completedNow = ($runspaces.Status.IsCompleted -eq $true).Count
            
            if ($completedNow -gt $completed) {
                $completed = $completedNow
                $percentComplete = [math]::Round(($completed / $totalIPs) * 100)
                $ScanProgressBar.Value = $percentComplete
                $TotalScannedText.Text = $completed.ToString()
                $ScanStatusText.Text = "Scanning in progress... ($completed of $totalIPs)"
                $ScanWindow.Dispatcher.Invoke([Action]{}, [Windows.Threading.DispatcherPriority]::Background)
            }
            
            Start-Sleep -Milliseconds 100
        }
        
        # Collect all results
        foreach ($runspace in $runspaces) {
            $result = $runspace.Pipe.EndInvoke($runspace.Status)
            $results.Add($result)
            
            if ($result.Status -eq "Online") {
                $onlineCount++
            } else {
                $offlineCount++
            }
            
            $runspace.Pipe.Dispose()
        }
        
        # Cleanup
        $runspacePool.Close()
        $runspacePool.Dispose()
        
        # Update final counts
        $OnlineCountText.Text = $onlineCount.ToString()
        $OfflineCountText.Text = $offlineCount.ToString()
        
        # Update grid with results
        $ResultsGrid.ItemsSource = $results
        
        # Complete
        $ScanStatusText.Text = "Scan complete! Found $onlineCount online devices out of $totalIPs addresses scanned."
        $ScanProgressBar.Value = 100
        $StartScanButton.IsEnabled = $true
        $ExportButton.IsEnabled = $true
        
        Write-Log "IP scan complete: $onlineCount online, $offlineCount offline"
    })
    
    # Export to Excel
    $ExportButton.Add_Click({
        if ($null -eq $ResultsGrid.ItemsSource -or $ResultsGrid.Items.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No scan results to export", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            [System.Windows.MessageBox]::Show("ImportExcel module is not installed.", "Module Missing", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Export failed: ImportExcel module not installed"
            return
        }
        
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $saveDialog.Title = "Save IP Scan Results"
            $saveDialog.FileName = "IPScan_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
            
            if ($saveDialog.ShowDialog()) {
                $excelPath = $saveDialog.FileName
                
                $ExportButton.IsEnabled = $false
                Write-Log "Exporting IP scan results to Excel: $excelPath"
                
                $exportData = @()
                foreach ($item in $ResultsGrid.ItemsSource) {
                    $exportData += [PSCustomObject]@{
                        'IP Address' = $item.IPAddress
                        'Hostname' = $item.Hostname
                        'Status' = $item.Status
                        'MAC Address' = $item.MACAddress
                        'Response Time (ms)' = $item.ResponseTime
                        'Scan Date' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
                
                # Export with formatting
                $exportData | Export-Excel -Path $excelPath `
                    -WorksheetName "IP Scan Results" `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle Medium1 `
                    -TableName "IPScanResults"
                
                Write-Log "Successfully exported $($exportData.Count) IP scan results to Excel"
                
                [System.Windows.MessageBox]::Show(
                    "IP scan results exported successfully!`n`nFile: $excelPath",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
                
                # Open folder location in Explorer
                $folderPath = Split-Path $excelPath -Parent
                Invoke-Item $folderPath
                
            } else {
                Write-Log "Export cancelled by user"
            }
            
        } catch {
            Write-Log "Export error: $($_.Exception.Message)"
            [System.Windows.MessageBox]::Show(
                "Error exporting to Excel:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        } finally {
            $ExportButton.IsEnabled = $true
        }
    })
    
    # Close button
    $CloseButton.Add_Click({ $ScanWindow.Close() })
    
    # Enter key support for IP boxes
    $StartIPBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $StartScanButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    $EndIPBox.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return) {
            $StartScanButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
        }
    })
    
    # Set default IPs for convenience (can be customized)
    $StartIPBox.Text = "192.168.1.1"
    $EndIPBox.Text = "192.168.1.254"
    
    $ScanWindow.ShowDialog() | Out-Null
})



# =====================================================
# INTUNE MOBILE DEVICES MODULE
# =====================================================

function Show-IntuneMobileDevicesWindow {
    $xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Intune Mobile Devices" 
    Height="700" 
    Width="1200" 
    WindowStartupLocation="CenterScreen"
    Background="#F5F5F5">
    
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        
        <Style TargetType="DataGrid">
            <Setter Property="AutoGenerateColumns" Value="True"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="SelectionMode" Value="Extended"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="RowBackground" Value="White"/>
            <Setter Property="AlternatingRowBackground" Value="#F9F9F9"/>
            <Setter Property="BorderBrush" Value="#E0E0E0"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#233A4A" Padding="20,15" Margin="-20,-20,0,15">
            <StackPanel>
                <TextBlock Text="Intune Mobile Devices" FontSize="20" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="View and export all mobile devices managed by Microsoft Intune" 
                          FontSize="12" Foreground="#B0BEC5" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Control Panel -->
        <Border Grid.Row="1" Background="White" BorderBrush="#E0E0E0" BorderThickness="1" 
                Padding="15" Margin="0,0,0,15" CornerRadius="4">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <!-- Top Row: Buttons and Status -->
                <Grid Grid.Row="0" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <StackPanel Grid.Column="0" Orientation="Horizontal" Margin="0,0,15,0">
                        <Button x:Name="LoadDevicesButton" Content="Load Devices" Width="130"/>
                        <Button x:Name="RefreshButton" Content="Refresh" Width="100" Background="#28a745"/>
                    </StackPanel>
                    
                    <!-- Search Box -->
                    <StackPanel Grid.Column="1" Orientation="Horizontal" Margin="0,0,15,0">
                        <TextBlock Text="Search:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="SemiBold"/>
                        <TextBox x:Name="SearchBox" Width="250" Height="26" VerticalContentAlignment="Center"
                                Padding="5,0,0,0" ToolTip="Search by device name, user, model, or OS"/>
                        <Button x:Name="ClearSearchButton" Content="Clear" Width="60" Margin="5,0,0,0"/>
                    </StackPanel>
                    
                    <TextBlock Grid.Column="2" x:Name="StatusText" Text="Ready" VerticalAlignment="Center" 
                              FontSize="13" Foreground="#666"/>
                    
                    <StackPanel Grid.Column="3" Orientation="Horizontal">
                        <TextBlock Text="Total Devices:" VerticalAlignment="Center" Margin="0,0,10,0" FontWeight="SemiBold"/>
                        <TextBlock x:Name="DeviceCountText" Text="0" VerticalAlignment="Center" 
                                  FontSize="16" FontWeight="Bold" Foreground="#0078D4" Margin="0,0,20,0"/>
                        <Button x:Name="ExportButton" Content="Export to Excel" Width="140" Background="#28a745"/>
                    </StackPanel>
                </Grid>
                
                <!-- Bottom Row: Device Type Filters -->
                <Border Grid.Row="1" Background="#F8F9FA" BorderBrush="#E0E0E0" BorderThickness="1" 
                        Padding="10" CornerRadius="3">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Device Type Filter (applied on Load/Refresh):" VerticalAlignment="Center" 
                                  FontWeight="SemiBold" Margin="0,0,15,0"/>
                        <CheckBox x:Name="FilteriOS" Content="iOS/iPadOS" IsChecked="True" 
                                 VerticalAlignment="Center" Margin="0,0,15,0"/>
                        <CheckBox x:Name="FilterAndroid" Content="Android" IsChecked="True" 
                                 VerticalAlignment="Center" Margin="0,0,15,0"/>
                        <CheckBox x:Name="FilterWindows" Content="Windows" IsChecked="True" 
                                 VerticalAlignment="Center" Margin="0,0,15,0"/>
                        <CheckBox x:Name="FilterMacOS" Content="macOS" IsChecked="True" 
                                 VerticalAlignment="Center"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>
        
        <!-- Summary Stats -->
        <Border Grid.Row="2" Background="White" BorderBrush="#E0E0E0" BorderThickness="1" 
                CornerRadius="4" Margin="0,0,0,15">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <!-- Stats Bar -->
                <Border Grid.Row="0" Background="#F8F9FA" BorderBrush="#E0E0E0" 
                       BorderThickness="0,0,0,1" Padding="15,10">
                    <StackPanel Orientation="Horizontal">
                        <StackPanel Orientation="Horizontal" Margin="0,0,25,0">
                            <TextBlock Text="iOS/iPadOS: " FontWeight="SemiBold" Foreground="#666"/>
                            <TextBlock x:Name="IosCountText" Text="0" FontWeight="Bold" Foreground="#0078D4"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,25,0">
                            <TextBlock Text="Android: " FontWeight="SemiBold" Foreground="#666"/>
                            <TextBlock x:Name="AndroidCountText" Text="0" FontWeight="Bold" Foreground="#3DDC84"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,25,0">
                            <TextBlock Text="Windows: " FontWeight="SemiBold" Foreground="#666"/>
                            <TextBlock x:Name="WindowsCountText" Text="0" FontWeight="Bold" Foreground="#0078D4"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,25,0">
                            <TextBlock Text="macOS: " FontWeight="SemiBold" Foreground="#666"/>
                            <TextBlock x:Name="MacOSCountText" Text="0" FontWeight="Bold" Foreground="#A2AAAD"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,0,25,0">
                            <TextBlock Text="Compliant: " FontWeight="SemiBold" Foreground="#666"/>
                            <TextBlock x:Name="CompliantCountText" Text="0" FontWeight="Bold" Foreground="#28a745"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Non-Compliant: " FontWeight="SemiBold" Foreground="#666"/>
                            <TextBlock x:Name="NonCompliantCountText" Text="0" FontWeight="Bold" Foreground="#dc3545"/>
                        </StackPanel>
                    </StackPanel>
                </Border>
                
                <!-- DataGrid -->
                <DataGrid x:Name="DevicesDataGrid" Grid.Row="1" Margin="15"/>
            </Grid>
        </Border>
        
        <!-- Close Button -->
        <Button Grid.Row="3" x:Name="CloseButton" Content="Close" Width="100" 
                HorizontalAlignment="Right" Background="#6c757d"/>
    </Grid>
</Window>
"@
    
    try {
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)
        $reader.Close()
        
        # Get controls
        $LoadDevicesButton = $window.FindName("LoadDevicesButton")
        $RefreshButton = $window.FindName("RefreshButton")
        $ExportButton = $window.FindName("ExportButton")
        $CloseButton = $window.FindName("CloseButton")
        $DevicesDataGrid = $window.FindName("DevicesDataGrid")
        $StatusText = $window.FindName("StatusText")
        $DeviceCountText = $window.FindName("DeviceCountText")
        $IosCountText = $window.FindName("IosCountText")
        $AndroidCountText = $window.FindName("AndroidCountText")
        $WindowsCountText = $window.FindName("WindowsCountText")
        $MacOSCountText = $window.FindName("MacOSCountText")
        $CompliantCountText = $window.FindName("CompliantCountText")
        $NonCompliantCountText = $window.FindName("NonCompliantCountText")
        $SearchBox = $window.FindName("SearchBox")
        $ClearSearchButton = $window.FindName("ClearSearchButton")
        
        # Get filter checkboxes
        $FilteriOS = $window.FindName("FilteriOS")
        $FilterAndroid = $window.FindName("FilterAndroid")
        $FilterWindows = $window.FindName("FilterWindows")
        $FilterMacOS = $window.FindName("FilterMacOS")
        
        # Script-level variable to store device data
        $script:intuneDevices = @()
        $script:allDevices = @()  # Store all devices before filtering
        
        # Export to Excel function
        function Export-IntuneDevicesToExcel {
            if ($script:intuneDevices.Count -eq 0) {
                [System.Windows.MessageBox]::Show(
                    "No devices to export. Please load devices first.",
                    "No Data",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                return
            }
            
            try {
                $ExportButton.IsEnabled = $false
                # Create SaveFileDialog
                $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
                $saveDialog.Title = "Save Intune Mobile Devices Report"
                $saveDialog.FileName = "IntuneMobileDevices_$(Get-Date -Format 'yyyyMMdd-HHmmss').xlsx"
                
                # Set initial directory to Documents\Reports\Intune if it exists
                $defaultFolder = Join-Path $env:USERPROFILE "Documents\Reports\Intune"
                if (Test-Path $defaultFolder) {
                    $saveDialog.InitialDirectory = $defaultFolder
                }
                
                if (-not $saveDialog.ShowDialog()) {
                    $ExportButton.IsEnabled = $true
                    return
                }
                
                $outputFile = $saveDialog.FileName
                $StatusText.Text = "Exporting to Excel..."
                $StatusText.Foreground = "#FF9800"
                
                # Export to Excel with formatting
                $script:intuneDevices | Export-Excel -Path $outputFile `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableStyle "Medium1" `
                    -WorksheetName "Mobile Devices"
                
                # Add conditional formatting
                $excel = Open-ExcelPackage -Path $outputFile
                $worksheet = $excel.Workbook.Worksheets["Mobile Devices"]
                
                # Find columns
                $headers = 1..$worksheet.Dimension.Columns | ForEach-Object {
                    $worksheet.Cells[1, $_].Value
                }
                
                # Format Compliance State column
                $complianceColIndex = ($headers.IndexOf("Compliance State")) + 1
                if ($complianceColIndex -gt 0) {
                    $colName = Get-ExcelColumnName $complianceColIndex
                    
                    Add-ConditionalFormatting -Worksheet $worksheet `
                        -Range "${colName}2:${colName}$($worksheet.Dimension.Rows)" `
                        -RuleType Equal -ConditionValue "compliant" `
                        -BackgroundColor ([System.Drawing.Color]::LightGreen) `
                        -ForegroundColor ([System.Drawing.Color]::DarkGreen)
                    
                    Add-ConditionalFormatting -Worksheet $worksheet `
                        -Range "${colName}2:${colName}$($worksheet.Dimension.Rows)" `
                        -RuleType Equal -ConditionValue "noncompliant" `
                        -BackgroundColor ([System.Drawing.Color]::LightPink) `
                        -ForegroundColor ([System.Drawing.Color]::DarkRed)
                }
                
                # Format Jail Broken column
                $jailBrokenColIndex = ($headers.IndexOf("Jail Broken")) + 1
                if ($jailBrokenColIndex -gt 0) {
                    $colName = Get-ExcelColumnName $jailBrokenColIndex
                    Add-ConditionalFormatting -Worksheet $worksheet `
                        -Range "${colName}2:${colName}$($worksheet.Dimension.Rows)" `
                        -RuleType Equal -ConditionValue "Detected" `
                        -BackgroundColor ([System.Drawing.Color]::Orange) `
                        -ForegroundColor ([System.Drawing.Color]::DarkRed)
                }
                
                # Format IMEI column as number
                $imeiColIndex = ($headers.IndexOf("IMEI")) + 1
                if ($imeiColIndex -gt 0) {
                    $colName = Get-ExcelColumnName $imeiColIndex
                    $worksheet.Cells["${colName}2:${colName}$($worksheet.Dimension.Rows)"].Style.Numberformat.Format = "0"
                }
                
                Close-ExcelPackage $excel
                
                $StatusText.Text = "Export complete"
                $StatusText.Foreground = "#28a745"
                
                # Ask to open file
                $result = [System.Windows.MessageBox]::Show(
                    "Export completed successfully!`n`n$outputFile`n`nWould you like to open the file?",
                    "Export Successful",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Information
                )
                
                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    Start-Process $outputFile
                }
                
            } catch {
                $errorMsg = $_.Exception.Message
                $StatusText.Text = "Export failed"
                $StatusText.Foreground = "#dc3545"
                [System.Windows.MessageBox]::Show(
                    "Failed to export to Excel:`n`n$errorMsg",
                    "Export Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
            } finally {
                $ExportButton.IsEnabled = $true
            }
        }
        
        # Helper function for Excel column names
        function Get-ExcelColumnName {
            param([int]$ColumnNumber)
            $columnName = ""
            while ($ColumnNumber -gt 0) {
                $modulo = ($ColumnNumber - 1) % 26
                $columnName = [char](65 + $modulo) + $columnName
                $ColumnNumber = [math]::Floor(($ColumnNumber - $modulo) / 26)
            }
            return $columnName
        }
        
        # Button event handlers
        $LoadDevicesButton.Add_Click({
            try {
                # Get fresh references to controls
                $LoadDevicesButton = $window.FindName("LoadDevicesButton")
                $RefreshButton = $window.FindName("RefreshButton")
                $StatusText = $window.FindName("StatusText")
                $DevicesDataGrid = $window.FindName("DevicesDataGrid")
                $ExportButton = $window.FindName("ExportButton")
                $DeviceCountText = $window.FindName("DeviceCountText")
                $IosCountText = $window.FindName("IosCountText")
                $AndroidCountText = $window.FindName("AndroidCountText")
                $WindowsCountText = $window.FindName("WindowsCountText")
                $MacOSCountText = $window.FindName("MacOSCountText")
                $CompliantCountText = $window.FindName("CompliantCountText")
                $NonCompliantCountText = $window.FindName("NonCompliantCountText")
                $SearchBox = $window.FindName("SearchBox")
                $FilteriOS = $window.FindName("FilteriOS")
                $FilterAndroid = $window.FindName("FilterAndroid")
                $FilterWindows = $window.FindName("FilterWindows")
                $FilterMacOS = $window.FindName("FilterMacOS")
                
                $LoadDevicesButton.IsEnabled = $false
                $RefreshButton.IsEnabled = $false
                $StatusText.Text = "Connecting to Microsoft Graph..."
                $StatusText.Foreground = "#FF9800"
                
                # Check if Microsoft.Graph modules are installed
                $requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.DeviceManagement')
                foreach ($module in $requiredModules) {
                    if (-not (Get-Module -ListAvailable -Name $module)) {
                        $StatusText.Text = "Installing $module module..."
                        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
                    }
                }
                
                # Import modules
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                Import-Module Microsoft.Graph.DeviceManagement -ErrorAction Stop
                
                # Check if already connected
                $context = Get-MgContext -ErrorAction SilentlyContinue
                if ($null -eq $context) {
                    # Connect to Microsoft Graph
                    Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "User.Read.All" -NoWelcome
                }
                
                $StatusText.Text = "Retrieving devices from Intune..."
                
                # Get all managed devices using pagination
                $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
                $allDevices = @()
                
                do {
                    $response = Invoke-MgGraphRequest -Uri $uri -Method GET
                    $allDevices += $response.value
                    $uri = $response.'@odata.nextLink'
                    
                    $StatusText.Text = "Retrieved $($allDevices.Count) devices..."
                } while ($uri)
                
                # Store all devices
                $script:allDevices = $allDevices
                
                $StatusText.Text = "Filtering devices..."
                
                # Apply filter based on checkboxes
                # Build list of selected operating systems
                $selectedOS = @()
                if ($FilteriOS.IsChecked) { $selectedOS += @('iOS', 'iPadOS') }
                if ($FilterAndroid.IsChecked) { $selectedOS += 'Android' }
                if ($FilterWindows.IsChecked) { $selectedOS += 'Windows' }
                if ($FilterMacOS.IsChecked) { $selectedOS += 'macOS' }
                
                # Filter devices based on selected OS types
                $filteredDevices = $script:allDevices | Where-Object {
                    $os = $_.operatingSystem
                    # Check if Windows checkbox is selected and OS contains Windows
                    if ($FilterWindows.IsChecked -and $os -like 'Windows*') {
                        return $true
                    }
                    # Check if OS matches any other selected types
                    return $os -in $selectedOS
                }
                
                # Format device information
                $script:intuneDevices = @()
                foreach ($device in $filteredDevices) {
                    $script:intuneDevices += [PSCustomObject]@{
                        'Device Name' = $device.deviceName
                        'User Display Name' = $device.userDisplayName
                        'User Principal Name' = $device.userPrincipalName
                        'Operating System' = $device.operatingSystem
                        'OS Version' = $device.osVersion
                        'Model' = $device.model
                        'Manufacturer' = $device.manufacturer
                        'IMEI' = $device.imei
                        'Serial Number' = $device.serialNumber
                        'Phone Number' = $device.phoneNumber
                        'Enrollment Date' = if ($device.enrolledDateTime) { 
                            (Get-Date $device.enrolledDateTime).ToString('yyyy-MM-dd HH:mm:ss') 
                        } else { 'N/A' }
                        'Last Sync' = if ($device.lastSyncDateTime) { 
                            (Get-Date $device.lastSyncDateTime).ToString('yyyy-MM-dd HH:mm:ss') 
                        } else { 'N/A' }
                        'Compliance State' = $device.complianceState
                        'Management State' = $device.managementState
                        'Ownership' = $device.managedDeviceOwnerType
                        'Supervised' = $device.isSupervised
                        'Encrypted' = $device.isEncrypted
                        'Jail Broken' = $device.jailBroken
                        'Total Storage (GB)' = if ($device.totalStorageSpaceInBytes) { 
                            [math]::Round($device.totalStorageSpaceInBytes / 1GB, 2) 
                        } else { 'N/A' }
                        'Free Storage (GB)' = if ($device.freeStorageSpaceInBytes) { 
                            [math]::Round($device.freeStorageSpaceInBytes / 1GB, 2) 
                        } else { 'N/A' }
                    }
                }
                
                # Update DataGrid
                $DevicesDataGrid.ItemsSource = $script:intuneDevices
                
                # Clear search box
                $SearchBox.Text = ""
                
                # Update statistics
                $total = $script:intuneDevices.Count
                $DeviceCountText.Text = $total.ToString()
                
                if ($total -gt 0) {
                    $iosCount = ($script:intuneDevices | Where-Object { $_.'Operating System' -in @('iOS', 'iPadOS') }).Count
                    $androidCount = ($script:intuneDevices | Where-Object { $_.'Operating System' -eq 'Android' }).Count
                    $windowsCount = ($script:intuneDevices | Where-Object { $_.'Operating System' -like 'Windows*' }).Count
                    $macosCount = ($script:intuneDevices | Where-Object { $_.'Operating System' -eq 'macOS' }).Count
                    $compliantCount = ($script:intuneDevices | Where-Object { $_.'Compliance State' -eq 'compliant' }).Count
                    $nonCompliantCount = ($script:intuneDevices | Where-Object { $_.'Compliance State' -eq 'noncompliant' }).Count
                    
                    $IosCountText.Text = $iosCount.ToString()
                    $AndroidCountText.Text = $androidCount.ToString()
                    $WindowsCountText.Text = $windowsCount.ToString()
                    $MacOSCountText.Text = $macosCount.ToString()
                    $CompliantCountText.Text = $compliantCount.ToString()
                    $NonCompliantCountText.Text = $nonCompliantCount.ToString()
                } else {
                    $IosCountText.Text = "0"
                    $AndroidCountText.Text = "0"
                    $WindowsCountText.Text = "0"
                    $MacOSCountText.Text = "0"
                    $CompliantCountText.Text = "0"
                    $NonCompliantCountText.Text = "0"
                }
                
                $StatusText.Text = "Loaded $($script:intuneDevices.Count) devices"
                $StatusText.Foreground = "#28a745"
                
                $ExportButton.IsEnabled = $true
                
            } catch {
                $errorMsg = $_.Exception.Message
                $StatusText = $window.FindName("StatusText")
                if ($StatusText) {
                    $StatusText.Text = "Error: Click for details"
                    $StatusText.Foreground = "#dc3545"
                }
                [System.Windows.MessageBox]::Show(
                    "Failed to load Intune devices:`n`n$errorMsg`n`nPlease ensure:`n- Microsoft Graph modules are installed`n- You have proper permissions`n- Network connection is available",
                    "Error Loading Devices",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
            } finally {
                $LoadDevicesButton = $window.FindName("LoadDevicesButton")
                $RefreshButton = $window.FindName("RefreshButton")
                if ($LoadDevicesButton) { $LoadDevicesButton.IsEnabled = $true }
                if ($RefreshButton) { $RefreshButton.IsEnabled = $true }
            }
        }.GetNewClosure())
        
        $RefreshButton.Add_Click({
            # Simply trigger the Load Devices button
            $LoadDevicesButton = $window.FindName("LoadDevicesButton")
            if ($LoadDevicesButton) {
                $LoadDevicesButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
            }
        }.GetNewClosure())
        
        # Search functionality
        $SearchBox.Add_TextChanged({
            try {
                # Get fresh references to controls
                $SearchBox = $window.FindName("SearchBox")
                $DevicesDataGrid = $window.FindName("DevicesDataGrid")
                $DeviceCountText = $window.FindName("DeviceCountText")
                
                if (-not $SearchBox -or -not $DevicesDataGrid -or -not $DeviceCountText) {
                    return
                }
                
                $searchTerm = $SearchBox.Text.Trim()
                
                if ([string]::IsNullOrWhiteSpace($searchTerm)) {
                    # If search is empty, show all devices
                    if ($script:intuneDevices) {
                        $DevicesDataGrid.ItemsSource = $script:intuneDevices
                        $DeviceCountText.Text = $script:intuneDevices.Count.ToString()
                    }
                } else {
                    # Filter devices based on search term
                    if ($script:intuneDevices -and $script:intuneDevices.Count -gt 0) {
                        $filtered = $script:intuneDevices | Where-Object {
                            $_.'Device Name' -like "*$searchTerm*" -or
                            $_.'User Display Name' -like "*$searchTerm*" -or
                            $_.'User Principal Name' -like "*$searchTerm*" -or
                            $_.'Operating System' -like "*$searchTerm*" -or
                            $_.'Model' -like "*$searchTerm*" -or
                            $_.'Serial Number' -like "*$searchTerm*" -or
                            $_.'IMEI' -like "*$searchTerm*"
                        }
                        
                        if ($filtered) {
                            $DevicesDataGrid.ItemsSource = $filtered
                            $DeviceCountText.Text = $filtered.Count.ToString()
                        } else {
                            $DevicesDataGrid.ItemsSource = @()
                            $DeviceCountText.Text = "0"
                        }
                    }
                }
            } catch {
                # Silently fail - search is not critical
            }
        }.GetNewClosure())
        
        $ClearSearchButton.Add_Click({
            $SearchBox = $window.FindName("SearchBox")
            $DevicesDataGrid = $window.FindName("DevicesDataGrid")
            $DeviceCountText = $window.FindName("DeviceCountText")
            
            $SearchBox.Text = ""
            $DevicesDataGrid.ItemsSource = $script:intuneDevices
            $DeviceCountText.Text = $script:intuneDevices.Count.ToString()
        }.GetNewClosure())
        
        $ExportButton.Add_Click({ Export-IntuneDevicesToExcel }.GetNewClosure())
        $CloseButton.Add_Click({ $window.Close() })
        
        # Initial state
        $ExportButton.IsEnabled = $false
        
        # Show window
        $window.ShowDialog() | Out-Null
        
    } catch {
        $errorMsg = $_.Exception.Message
        [System.Windows.MessageBox]::Show(
            "Failed to open Intune Mobile Devices window:`n`n$errorMsg",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Button Click Handler
$syncHash.IntuneMobileButton.Add_Click({
    Write-Log "Opening Intune Mobile Devices window..."
    Show-IntuneMobileDevicesWindow
})



$syncHash.PermissionAuditButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Permission Audit Report feature is planned for version 3.6.0",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})



# Intune & SCCM - Future Features
$syncHash.SCCMDevicesButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "SCCM Device Management feature is planned for a future version.`n`nThis will include:`n- View all SCCM managed devices`n- Device collections`n- Deployment status`n- Hardware inventory",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

$syncHash.IntuneComplianceButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Compliance Policy Reports feature is planned for a future version.`n`nThis will include:`n- Compliance status reports`n- Policy assignment details`n- Non-compliant device lists`n- Trend analysis",
        "Coming Soon",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
})

# Microsoft 365 - Generate Temporary Access Pass
$syncHash.GenerateTAPButton.Add_Click({
    Write-Log "Opening Generate Temporary Access Pass window..."
    Show-GenerateTAPWindow
})

function Show-GenerateTAPWindow {
    # Check for Microsoft Graph module
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        [System.Windows.MessageBox]::Show(
            "Microsoft Graph Authentication module is not available.`n`nPlease install the Microsoft Graph PowerShell SDK.",
            "Module Required",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }
    
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    } catch {
        [System.Windows.MessageBox]::Show(
            "Failed to load Microsoft Graph Authentication module.`n`n$($_.Exception.Message)",
            "Module Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }
    
    try {
        # Check if Microsoft Graph is connected
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -eq $graphContext) {
            $result = [System.Windows.MessageBox]::Show(
                "You are not connected to Microsoft Graph.`n`nWould you like to connect now?`n`nRequired permissions: UserAuthenticationMethod.ReadWrite.All",
                "Connection Required",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question
            )
            
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                try {
                    Connect-MgGraph -Scopes "UserAuthenticationMethod.ReadWrite.All" -NoWelcome
                    $graphContext = Get-MgContext -ErrorAction SilentlyContinue
                    
                    if ($null -eq $graphContext) {
                        [System.Windows.MessageBox]::Show(
                            "Failed to connect to Microsoft Graph. Please try again.",
                            "Connection Failed",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Error
                        )
                        return
                    }
                } catch {
                    $errorMsg = $_.Exception.Message
                    [System.Windows.MessageBox]::Show(
                        "Failed to connect to Microsoft Graph:`n`n$errorMsg",
                        "Connection Error",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Error
                    )
                    return
                }
            } else {
                return
            }
        }

        [xml]$TAPWindowXAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Generate Temporary Access Pass" 
        Height="550" 
        Width="650" 
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Instructions -->
        <Border Grid.Row="0" Background="#e3f2fd" BorderBrush="#2196f3" BorderThickness="1" Padding="10" Margin="0,0,0,15" CornerRadius="3">
            <StackPanel>
                <TextBlock Text="Temporary Access Pass (TAP) Generator" FontWeight="Bold" FontSize="14" Foreground="#1976d2" Margin="0,0,0,5"/>
                <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#555">
                    A Temporary Access Pass is a time-limited passcode that can be used as a strong credential to onboard passwordless methods like Windows Hello for Business, Microsoft Authenticator, or FIDO2 security keys.
                </TextBlock>
            </StackPanel>
        </Border>

        <!-- User Input -->
        <GroupBox Grid.Row="1" Header="User Account" Padding="10" Margin="0,0,0,15">
            <StackPanel>
                <TextBlock Text="Enter the User Principal Name (UPN) or Username:" Margin="0,0,0,5"/>
                <TextBox x:Name="UsernameTextBox" 
                         Height="30" 
                         Padding="5"
                         ToolTip="Example: john.doe@company.com or john.doe"/>
            </StackPanel>
        </GroupBox>

        <!-- TAP Settings -->
        <GroupBox Grid.Row="2" Header="TAP Settings" Padding="10" Margin="0,0,0,15">
            <StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="Lifetime (minutes):" Width="150" VerticalAlignment="Center"/>
                    <TextBox x:Name="LifetimeTextBox" 
                             Width="100" 
                             Height="30" 
                             Padding="5"
                             Text="60"
                             ToolTip="TAP validity period in minutes (10-43200)"/>
                </StackPanel>
                
                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="One-time use:" Width="150" VerticalAlignment="Center"/>
                    <CheckBox x:Name="OneTimeUseCheckBox" 
                              IsChecked="True" 
                              VerticalAlignment="Center"
                              ToolTip="If checked, TAP can only be used once"/>
                </StackPanel>
            </StackPanel>
        </GroupBox>

        <!-- Result Display -->
        <GroupBox Grid.Row="4" Header="Generated Temporary Access Pass" Padding="10" x:Name="ResultGroupBox" Visibility="Collapsed">
            <StackPanel>
                <Border Background="#fff3cd" BorderBrush="#ffc107" BorderThickness="1" Padding="10" Margin="0,0,0,10" CornerRadius="3">
                    <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#856404">
                        IMPORTANT: This TAP will only be displayed once. Copy it now and provide it to the user. It cannot be retrieved again.
                    </TextBlock>
                </Border>
                
                <TextBlock Text="Temporary Access Pass:" FontWeight="Bold" Margin="0,0,0,5"/>
                <Border Background="White" BorderBrush="#dee2e6" BorderThickness="1" Padding="10" CornerRadius="3">
                    <TextBox x:Name="TAPResultTextBox" 
                             IsReadOnly="True" 
                             FontFamily="Consolas"
                             FontSize="18"
                             FontWeight="Bold"
                             Background="Transparent"
                             BorderThickness="0"
                             Foreground="#28a745"
                             TextAlignment="Center"
                             HorizontalAlignment="Stretch"/>
                </Border>
                
                <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                    <Button x:Name="CopyTAPButton" 
                            Content="Copy to Clipboard" 
                            Width="150" 
                            Height="35" 
                            Margin="0,0,10,0"
                            Background="#007bff"
                            Foreground="White"
                            FontWeight="Bold"
                            Cursor="Hand"/>
                    <StackPanel x:Name="TAPDetailsPanel" VerticalAlignment="Center">
                        <TextBlock x:Name="TAPDetailsText" 
                                   FontSize="10" 
                                   Foreground="#666"
                                   TextWrapping="Wrap"/>
                    </StackPanel>
                </StackPanel>
            </StackPanel>
        </GroupBox>

        <!-- Buttons -->
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="GenerateButton" 
                    Content="Generate TAP" 
                    Width="120" 
                    Height="35" 
                    Margin="0,0,10,0"
                    Background="#28a745"
                    Foreground="White"
                    FontWeight="Bold"
                    Cursor="Hand"/>
            <Button x:Name="CloseButton" 
                    Content="Close" 
                    Width="100" 
                    Height="35"
                    Background="#6c757d"
                    Foreground="White"
                    FontWeight="Bold"
                    Cursor="Hand"/>
        </StackPanel>
    </Grid>
</Window>
"@

        $tapReader = New-Object System.Xml.XmlNodeReader $TAPWindowXAML
        $TAPWindow = [Windows.Markup.XamlReader]::Load($tapReader)

        # Get controls
        $UsernameTextBox = $TAPWindow.FindName("UsernameTextBox")
        $LifetimeTextBox = $TAPWindow.FindName("LifetimeTextBox")
        $OneTimeUseCheckBox = $TAPWindow.FindName("OneTimeUseCheckBox")
        $GenerateButton = $TAPWindow.FindName("GenerateButton")
        $CloseButton = $TAPWindow.FindName("CloseButton")
        $ResultGroupBox = $TAPWindow.FindName("ResultGroupBox")
        $TAPResultTextBox = $TAPWindow.FindName("TAPResultTextBox")
        $CopyTAPButton = $TAPWindow.FindName("CopyTAPButton")
        $TAPDetailsText = $TAPWindow.FindName("TAPDetailsText")

        # Generate Button Click
        $GenerateButton.Add_Click({
            try {
                $username = $UsernameTextBox.Text.Trim()
                if ([string]::IsNullOrWhiteSpace($username)) {
                    [System.Windows.MessageBox]::Show(
                        "Please enter a username or UPN.",
                        "Input Required",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                    return
                }

                # Validate lifetime
                $lifetime = 60
                if (-not [int]::TryParse($LifetimeTextBox.Text, [ref]$lifetime) -or $lifetime -lt 10 -or $lifetime -gt 43200) {
                    [System.Windows.MessageBox]::Show(
                        "Lifetime must be a number between 10 and 43200 minutes.",
                        "Invalid Lifetime",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                    return
                }

                $GenerateButton.IsEnabled = $false
                $GenerateButton.Content = "Generating..."

                # Get user ID
                $user = Get-MgUser -Filter "userPrincipalName eq '$username' or mailNickname eq '$username'" -ErrorAction Stop
                if ($null -eq $user) {
                    throw "User not found: $username"
                }

                # Create TAP parameters
                $tapParams = @{
                    LifetimeInMinutes = $lifetime
                    IsUsableOnce = $OneTimeUseCheckBox.IsChecked
                }

                # Generate TAP
                $tap = New-MgUserAuthenticationTemporaryAccessPassMethod -UserId $user.Id -BodyParameter $tapParams -ErrorAction Stop

                # Display result
                $TAPResultTextBox.Text = $tap.TemporaryAccessPass
                $ResultGroupBox.Visibility = [System.Windows.Visibility]::Visible
                
                $expiryTime = (Get-Date).AddMinutes($lifetime).ToString("MM/dd/yyyy hh:mm tt")
                $usageType = if ($OneTimeUseCheckBox.IsChecked) { "One-time use" } else { "Multi-use" }
                $TAPDetailsText.Text = "User: $($user.DisplayName)`nExpires: $expiryTime`nType: $usageType"
                
                $GenerateButton.Content = "Generate Another"
                $GenerateButton.IsEnabled = $true

                # Focus username field for next generation
                $UsernameTextBox.Focus()

            } catch {
                $errorMsg = $_.Exception.Message
                [System.Windows.MessageBox]::Show(
                    "Failed to generate Temporary Access Pass:`n`n$errorMsg",
                    "Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
                $GenerateButton.Content = "Generate TAP"
                $GenerateButton.IsEnabled = $true
            }
        })

        # Copy Button Click
        $CopyTAPButton.Add_Click({
            try {
                $tap = $TAPResultTextBox.Text
                if (-not [string]::IsNullOrWhiteSpace($tap)) {
                    [System.Windows.Clipboard]::SetText($tap)
                    $CopyTAPButton.Content = "Copied!"
                    
                    # Reset button text after 2 seconds
                    $timer = New-Object System.Windows.Threading.DispatcherTimer
                    $timer.Interval = [TimeSpan]::FromSeconds(2)
                    $timer.Add_Tick({
                        $CopyTAPButton.Content = "Copy to Clipboard"
                        $timer.Stop()
                    })
                    $timer.Start()
                }
            } catch {
                # Silently handle clipboard errors
            }
        })

        # Close Button Click
        $CloseButton.Add_Click({
            $TAPWindow.Close()
        })

        # Enter key support for username field
        $UsernameTextBox.Add_KeyDown({
            param($sender, $e)
            if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
                $GenerateButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            }
        })

        $TAPWindow.ShowDialog() | Out-Null

    } catch {
        $errorMsg = $_.Exception.Message
        [System.Windows.MessageBox]::Show(
            "Failed to open Generate TAP window:`n`n$errorMsg",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

Write-Log "IT Operations Center initialized"
Write-Log "Ready for IT operations management"

$Window.ShowDialog() | Out-Null