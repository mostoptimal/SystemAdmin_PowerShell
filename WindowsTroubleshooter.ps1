<#
.SYNOPSIS
    Windows Troubleshooting Tool for System Administrators
.DESCRIPTION
    A portable, single-file PowerShell GUI application that provides comprehensive
    Windows system diagnostics, troubleshooting, and monitoring capabilities.
    Uses Windows Forms for the GUI - runs on any Windows 10/11 system without dependencies.
.EXAMPLE
    PS> .\WindowsTroubleshooter.ps1
    Launches the GUI application
.EXAMPLE
    Right-click -> Run with PowerShell (as Administrator)
    Launches with elevated privileges for full functionality
.NOTES
    Author: System Administrator Tools
    Version: 1.0.0
    Requires: PowerShell 5.1+ (built into Windows 10/11)
#>

#Requires -Version 5.1

# ============================================================================
# INITIALIZATION & ADMIN CHECK
# ============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# Check if running as administrator
$script:IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# Get script directory for locating scripts folder
$script:ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
$script:ScriptsFolder = Join-Path $script:ScriptRoot "scripts"

# Global state
$script:CurrentJob = $null
$script:OutputHistory = @()
$script:LastRunTime = $null

# ============================================================================
# COLOR SCHEME & STYLING
# ============================================================================

$script:Colors = @{
    Background      = [System.Drawing.Color]::FromArgb(30, 30, 30)
    BackgroundLight = [System.Drawing.Color]::FromArgb(45, 45, 45)
    BackgroundDark  = [System.Drawing.Color]::FromArgb(20, 20, 20)
    Accent          = [System.Drawing.Color]::FromArgb(0, 122, 204)
    AccentHover     = [System.Drawing.Color]::FromArgb(28, 151, 234)
    Text            = [System.Drawing.Color]::FromArgb(220, 220, 220)
    TextDim         = [System.Drawing.Color]::FromArgb(150, 150, 150)
    Success         = [System.Drawing.Color]::FromArgb(76, 175, 80)
    Warning         = [System.Drawing.Color]::FromArgb(255, 193, 7)
    Error           = [System.Drawing.Color]::FromArgb(244, 67, 54)
    Border          = [System.Drawing.Color]::FromArgb(60, 60, 60)
}

$script:Fonts = @{
    Title    = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    Header   = New-Object System.Drawing.Font("Segoe UI Semibold", 11)
    Normal   = New-Object System.Drawing.Font("Segoe UI", 10)
    Small    = New-Object System.Drawing.Font("Segoe UI", 9)
    Mono     = New-Object System.Drawing.Font("Cascadia Code,Consolas", 10)
    MonoSmall = New-Object System.Drawing.Font("Cascadia Code,Consolas", 9)
}

# ============================================================================
# SCRIPT CATEGORIES & DEFINITIONS
# ============================================================================

$script:Categories = @{
    "Dashboard" = @{
        Icon = [char]0x2302  # House
        Description = "Quick system health overview"
        Scripts = @()  # Special tab - no scripts list
    }
    "System Health" = @{
        Icon = [char]0x2665  # Heart
        Description = "CPU, RAM, drives, and hardware diagnostics"
        Scripts = @(
            @{ Name = "check-cpu"; Description = "Check CPU status, temperature, and usage"; Admin = $false }
            @{ Name = "check-ram"; Description = "Check RAM modules and memory status"; Admin = $false }
            @{ Name = "check-drives"; Description = "Check all drive health and S.M.A.R.T. data"; Admin = $false }
            @{ Name = "check-drive-space"; Description = "Check available disk space on all drives"; Admin = $false }
            @{ Name = "check-health"; Description = "Comprehensive system health check"; Admin = $false }
            @{ Name = "check-hardware"; Description = "Full hardware diagnostics"; Admin = $false }
            @{ Name = "check-bios"; Description = "Check BIOS/UEFI information"; Admin = $false }
            @{ Name = "check-gpu"; Description = "Check GPU status and details"; Admin = $false }
            @{ Name = "check-motherboard"; Description = "Check motherboard information"; Admin = $false }
            @{ Name = "check-power"; Description = "Check power supply and battery status"; Admin = $false }
            @{ Name = "check-uptime"; Description = "Check system uptime since last boot"; Admin = $false }
            @{ Name = "check-swap-space"; Description = "Check virtual memory/page file status"; Admin = $false }
            @{ Name = "check-temp-dir"; Description = "Check temporary directory status"; Admin = $false }
            @{ Name = "query-smart-data"; Description = "Query S.M.A.R.T. data from drives"; Admin = $true }
            @{ Name = "list-system-info"; Description = "List comprehensive system information"; Admin = $false }
            @{ Name = "list-os-releases"; Description = "List Windows version and build info"; Admin = $false }
            @{ Name = "list-timezone"; Description = "Display current timezone settings"; Admin = $false }
        )
    }
    "Network" = @{
        Icon = [char]0x260D  # Network symbol
        Description = "Network connectivity and diagnostics"
        Scripts = @(
            @{ Name = "check-dns"; Description = "Test DNS resolution speed and servers"; Admin = $false }
            @{ Name = "check-network"; Description = "Comprehensive network health check"; Admin = $false }
            @{ Name = "check-vpn"; Description = "Check VPN connection status"; Admin = $false }
            @{ Name = "check-firewall"; Description = "Check Windows Firewall status"; Admin = $false }
            @{ Name = "check-ip-address"; Description = "Check public and private IP addresses"; Admin = $false }
            @{ Name = "check-mac-address"; Description = "Check network adapter MAC addresses"; Admin = $false }
            @{ Name = "check-subnet-mask"; Description = "Check network subnet configuration"; Admin = $false }
            @{ Name = "check-default-gateway"; Description = "Check default gateway settings"; Admin = $false }
            @{ Name = "list-network-shares"; Description = "List available network shares"; Admin = $false }
            @{ Name = "list-network-adapters"; Description = "List all network adapters"; Admin = $false }
            @{ Name = "list-local-ip"; Description = "List local IP addresses"; Admin = $false }
            @{ Name = "list-public-ip"; Description = "Get public IP address"; Admin = $false }
            @{ Name = "ping-host"; Description = "Ping a specified host"; Admin = $false }
            @{ Name = "scan-ports"; Description = "Scan common ports on a host"; Admin = $false }
            @{ Name = "test-internet"; Description = "Test internet connectivity"; Admin = $false }
            @{ Name = "wake-on-lan"; Description = "Send Wake-on-LAN packet"; Admin = $false }
            @{ Name = "list-wifi-networks"; Description = "List available WiFi networks"; Admin = $false }
        )
    }
    "Services" = @{
        Icon = [char]0x2699  # Gear
        Description = "Windows services and processes management"
        Scripts = @(
            @{ Name = "list-services"; Description = "List all Windows services with status"; Admin = $false }
            @{ Name = "list-processes"; Description = "List running processes"; Admin = $false }
            @{ Name = "list-tasks"; Description = "List scheduled tasks"; Admin = $false }
            @{ Name = "list-automatic-services"; Description = "List services set to start automatically"; Admin = $false }
            @{ Name = "list-running-services"; Description = "List currently running services"; Admin = $false }
            @{ Name = "list-stopped-services"; Description = "List stopped services"; Admin = $false }
            @{ Name = "check-windows-updates"; Description = "Check for pending Windows updates"; Admin = $false }
            @{ Name = "list-installed-updates"; Description = "List installed Windows updates"; Admin = $false }
            @{ Name = "list-print-jobs"; Description = "List pending print jobs"; Admin = $false }
            @{ Name = "list-printers"; Description = "List installed printers"; Admin = $false }
            @{ Name = "list-clipboard"; Description = "Show clipboard contents"; Admin = $false }
        )
    }
    "Software" = @{
        Icon = [char]0x2630  # Apps
        Description = "Installed software and applications"
        Scripts = @(
            @{ Name = "list-apps"; Description = "List installed applications"; Admin = $false }
            @{ Name = "list-installed-software"; Description = "List all installed software"; Admin = $false }
            @{ Name = "list-cli-tools"; Description = "List available CLI tools"; Admin = $false }
            @{ Name = "list-modules"; Description = "List PowerShell modules"; Admin = $false }
            @{ Name = "list-cmdlets"; Description = "List available cmdlets"; Admin = $false }
            @{ Name = "list-aliases"; Description = "List PowerShell aliases"; Admin = $false }
            @{ Name = "list-scripts"; Description = "List scripts in repository"; Admin = $false }
            @{ Name = "check-powershell"; Description = "Check PowerShell version and config"; Admin = $false }
            @{ Name = "check-dotnet"; Description = "Check .NET Framework versions"; Admin = $false }
            @{ Name = "list-environment-variables"; Description = "List environment variables"; Admin = $false }
            @{ Name = "list-path"; Description = "List PATH environment variable"; Admin = $false }
        )
    }
    "Storage" = @{
        Icon = [char]0x2395  # Disk
        Description = "File system and storage management"
        Scripts = @(
            @{ Name = "list-drives"; Description = "List all drives and partitions"; Admin = $false }
            @{ Name = "list-volumes"; Description = "List storage volumes"; Admin = $false }
            @{ Name = "list-partitions"; Description = "List disk partitions"; Admin = $false }
            @{ Name = "check-file-system"; Description = "Check file system health"; Admin = $true }
            @{ Name = "list-recycle-bin"; Description = "List recycle bin contents"; Admin = $false }
            @{ Name = "clear-recycle-bin"; Description = "Empty the recycle bin"; Admin = $false }
            @{ Name = "list-hidden-files"; Description = "List hidden files in a directory"; Admin = $false }
            @{ Name = "list-empty-dirs"; Description = "Find empty directories"; Admin = $false }
            @{ Name = "list-large-files"; Description = "Find large files"; Admin = $false }
            @{ Name = "list-duplicate-files"; Description = "Find duplicate files"; Admin = $false }
            @{ Name = "list-recent-files"; Description = "List recently modified files"; Admin = $false }
            @{ Name = "check-repo"; Description = "Check Git repository status"; Admin = $false }
        )
    }
    "Security" = @{
        Icon = [char]0x26A0  # Shield
        Description = "Security status and configuration"
        Scripts = @(
            @{ Name = "check-windows-system-files"; Description = "Verify Windows system file integrity"; Admin = $true }
            @{ Name = "list-user-groups"; Description = "List user groups and memberships"; Admin = $false }
            @{ Name = "list-users"; Description = "List local user accounts"; Admin = $false }
            @{ Name = "list-current-user"; Description = "Show current user details"; Admin = $false }
            @{ Name = "list-credentials"; Description = "List stored credentials"; Admin = $false }
            @{ Name = "check-symlinks"; Description = "Check for symbolic links"; Admin = $false }
            @{ Name = "list-certificates"; Description = "List installed certificates"; Admin = $false }
            @{ Name = "add-firewall-rules"; Description = "Add firewall rules"; Admin = $true }
            @{ Name = "list-firewall-rules"; Description = "List firewall rules"; Admin = $false }
        )
    }
    "All Scripts" = @{
        Icon = [char]0x2261  # Menu
        Description = "Browse all available scripts"
        Scripts = @()  # Will be populated dynamically
    }
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Get-ScriptPath {
    param([string]$ScriptName)
    $path = Join-Path $script:ScriptsFolder "$ScriptName.ps1"
    if (Test-Path $path) { return $path }
    return $null
}

function Get-AllScripts {
    $allScripts = @()
    if (Test-Path $script:ScriptsFolder) {
        Get-ChildItem -Path $script:ScriptsFolder -Filter "*.ps1" | ForEach-Object {
            $name = $_.BaseName
            $description = ""
            # Try to extract description from script
            $content = Get-Content $_.FullName -TotalCount 20 -ErrorAction SilentlyContinue
            $synopsisFound = $false
            foreach ($line in $content) {
                if ($line -match '\.SYNOPSIS') { $synopsisFound = $true; continue }
                if ($synopsisFound -and $line -match '^\s+(.+)$') {
                    $description = $Matches[1].Trim()
                    break
                }
                if ($line -match '\.DESCRIPTION') { break }
            }
            $allScripts += @{
                Name = $name
                Description = if ($description) { $description } else { "Run $name script" }
                Admin = $false
            }
        }
    }
    return $allScripts | Sort-Object { $_.Name }
}

function Format-OutputText {
    param([string]$Text)
    # Convert emoji status indicators for display
    $Text = $Text -replace '✅', '[OK] '
    $Text = $Text -replace '⚠️', '[WARN] '
    $Text = $Text -replace '❌', '[ERROR] '
    return $Text
}

function Get-StatusColor {
    param([string]$Text)
    if ($Text -match '✅|\[OK\]|success|healthy|normal|running') {
        return $script:Colors.Success
    }
    elseif ($Text -match '⚠️|\[WARN\]|warning|pending|stopped') {
        return $script:Colors.Warning
    }
    elseif ($Text -match '❌|\[ERROR\]|error|failed|critical') {
        return $script:Colors.Error
    }
    return $script:Colors.Text
}

function Invoke-ScriptAsync {
    param(
        [string]$ScriptPath,
        [System.Windows.Forms.TextBox]$OutputBox,
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabel
    )
    
    if (-not (Test-Path $ScriptPath)) {
        $OutputBox.AppendText("`r`n[ERROR] Script not found: $ScriptPath`r`n")
        return
    }
    
    $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
    $timestamp = Get-Date -Format "HH:mm:ss"
    
    $OutputBox.AppendText("`r`n[$timestamp] Running: $scriptName`r`n")
    $OutputBox.AppendText("=" * 60 + "`r`n")
    $StatusLabel.Text = "Running: $scriptName..."
    
    try {
        # Run script and capture output
        $output = & $ScriptPath 2>&1 | Out-String
        $formattedOutput = Format-OutputText $output
        $OutputBox.AppendText($formattedOutput)
        $OutputBox.AppendText("`r`n")
        
        $script:LastRunTime = Get-Date
        $StatusLabel.Text = "Completed: $scriptName | Last run: $($script:LastRunTime.ToString('HH:mm:ss'))"
        
        # Store in history
        $script:OutputHistory += @{
            Timestamp = $timestamp
            Script = $scriptName
            Output = $formattedOutput
        }
    }
    catch {
        $OutputBox.AppendText("`r`n[ERROR] $($_.Exception.Message)`r`n")
        $StatusLabel.Text = "Error running: $scriptName"
    }
    
    $OutputBox.ScrollToCaret()
}

function Run-QuickHealthCheck {
    param(
        [System.Windows.Forms.TextBox]$OutputBox,
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabel,
        [hashtable]$StatusLabels
    )
    
    $OutputBox.Clear()
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $OutputBox.AppendText("WINDOWS SYSTEM HEALTH REPORT`r`n")
    $OutputBox.AppendText("Generated: $timestamp`r`n")
    $OutputBox.AppendText("Computer: $env:COMPUTERNAME`r`n")
    $OutputBox.AppendText("User: $env:USERNAME`r`n")
    $OutputBox.AppendText("=" * 60 + "`r`n`r`n")
    
    $healthScripts = @(
        @{ Name = "check-uptime"; Label = "UPTIME" }
        @{ Name = "check-cpu"; Label = "CPU" }
        @{ Name = "check-ram"; Label = "RAM" }
        @{ Name = "check-drive-space"; Label = "STORAGE" }
        @{ Name = "check-network"; Label = "NETWORK" }
    )
    
    $StatusLabel.Text = "Running health check..."
    
    foreach ($item in $healthScripts) {
        $scriptPath = Get-ScriptPath $item.Name
        if ($scriptPath) {
            $OutputBox.AppendText("[$($item.Label)]`r`n")
            try {
                $output = & $scriptPath 2>&1 | Out-String
                $formattedOutput = Format-OutputText $output
                $OutputBox.AppendText($formattedOutput)
                
                # Update dashboard status if available
                if ($StatusLabels -and $StatusLabels.ContainsKey($item.Label)) {
                    $status = if ($output -match '✅') { "OK" } 
                              elseif ($output -match '⚠️') { "Warning" } 
                              else { "Check" }
                    $StatusLabels[$item.Label].Text = $status
                }
            }
            catch {
                $OutputBox.AppendText("[ERROR] Could not run $($item.Name)`r`n")
            }
            $OutputBox.AppendText("`r`n")
        }
    }
    
    $OutputBox.AppendText("=" * 60 + "`r`n")
    $OutputBox.AppendText("Health check completed at $(Get-Date -Format 'HH:mm:ss')`r`n")
    $OutputBox.ScrollToCaret()
    
    $script:LastRunTime = Get-Date
    $StatusLabel.Text = "Health check completed | $($script:LastRunTime.ToString('HH:mm:ss'))"
}

function Export-Results {
    param(
        [string]$Content,
        [string]$Format = "txt"
    )
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Title = "Export Results"
    $saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    if ($Format -eq "html") {
        $saveDialog.Filter = "HTML Files (*.html)|*.html"
        $saveDialog.FileName = "SystemReport_$timestamp.html"
    }
    else {
        $saveDialog.Filter = "Text Files (*.txt)|*.txt"
        $saveDialog.FileName = "SystemReport_$timestamp.txt"
    }
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($Format -eq "html") {
            $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>System Report - $env:COMPUTERNAME</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #1e1e1e; color: #ddd; padding: 20px; }
        h1 { color: #0078d4; border-bottom: 2px solid #0078d4; padding-bottom: 10px; }
        pre { background: #2d2d2d; padding: 15px; border-radius: 5px; overflow-x: auto; }
        .ok { color: #4caf50; }
        .warn { color: #ffc107; }
        .error { color: #f44336; }
        .timestamp { color: #888; font-size: 0.9em; }
    </style>
</head>
<body>
    <h1>Windows System Report</h1>
    <p class="timestamp">Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Computer: $env:COMPUTERNAME</p>
    <pre>$([System.Web.HttpUtility]::HtmlEncode($Content))</pre>
</body>
</html>
"@
            $htmlContent | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
        }
        else {
            $Content | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
        }
        
        [System.Windows.Forms.MessageBox]::Show(
            "Report saved to:`n$($saveDialog.FileName)",
            "Export Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
}

function Request-AdminElevation {
    if (-not $script:IsAdmin) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "This action requires administrator privileges.`n`nDo you want to restart the application as Administrator?",
            "Administrator Required",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $scriptPath = $MyInvocation.MyCommand.Definition
            if (-not $scriptPath) { $scriptPath = $PSCommandPath }
            Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
            $script:MainForm.Close()
        }
        return $false
    }
    return $true
}

# ============================================================================
# BUILD MAIN FORM
# ============================================================================

$script:MainForm = New-Object System.Windows.Forms.Form
$MainForm = $script:MainForm
$MainForm.Text = "Windows Troubleshooting Tool"
$MainForm.Size = New-Object System.Drawing.Size(1200, 800)
$MainForm.MinimumSize = New-Object System.Drawing.Size(900, 600)
$MainForm.StartPosition = "CenterScreen"
$MainForm.BackColor = $script:Colors.Background
$MainForm.ForeColor = $script:Colors.Text
$MainForm.Font = $script:Fonts.Normal

# Set icon (using shield for admin tool)
try {
    $MainForm.Icon = [System.Drawing.SystemIcons]::Shield
} catch {}

# ============================================================================
# HEADER PANEL
# ============================================================================

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$headerPanel.Height = 60
$headerPanel.BackColor = $script:Colors.BackgroundDark
$headerPanel.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 10)

# Title
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Windows Troubleshooting Tool"
$titleLabel.Font = $script:Fonts.Title
$titleLabel.ForeColor = $script:Colors.Text
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(15, 18)
$headerPanel.Controls.Add($titleLabel)

# Admin status indicator
$adminLabel = New-Object System.Windows.Forms.Label
$adminLabel.AutoSize = $true
$adminLabel.Location = New-Object System.Drawing.Point(350, 22)
if ($script:IsAdmin) {
    $adminLabel.Text = "[Administrator]"
    $adminLabel.ForeColor = $script:Colors.Success
} else {
    $adminLabel.Text = "[Standard User]"
    $adminLabel.ForeColor = $script:Colors.Warning
}
$adminLabel.Font = $script:Fonts.Small
$headerPanel.Controls.Add($adminLabel)

# Elevate button (if not admin)
if (-not $script:IsAdmin) {
    $elevateBtn = New-Object System.Windows.Forms.Button
    $elevateBtn.Text = "Run as Admin"
    $elevateBtn.Size = New-Object System.Drawing.Size(110, 30)
    $elevateBtn.Location = New-Object System.Drawing.Point(($MainForm.ClientSize.Width - 270), 15)
    $elevateBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $elevateBtn.BackColor = $script:Colors.Warning
    $elevateBtn.ForeColor = $script:Colors.BackgroundDark
    $elevateBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $elevateBtn.Font = $script:Fonts.Small
    $elevateBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $elevateBtn.Add_Click({ Request-AdminElevation })
    $headerPanel.Controls.Add($elevateBtn)
}

# Export button
$exportBtn = New-Object System.Windows.Forms.Button
$exportBtn.Text = "Export"
$exportBtn.Size = New-Object System.Drawing.Size(80, 30)
$exportBtn.Location = New-Object System.Drawing.Point(($MainForm.ClientSize.Width - 150), 15)
$exportBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$exportBtn.BackColor = $script:Colors.Accent
$exportBtn.ForeColor = $script:Colors.Text
$exportBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$exportBtn.Font = $script:Fonts.Small
$exportBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$headerPanel.Controls.Add($exportBtn)

# Help button
$helpBtn = New-Object System.Windows.Forms.Button
$helpBtn.Text = "?"
$helpBtn.Size = New-Object System.Drawing.Size(30, 30)
$helpBtn.Location = New-Object System.Drawing.Point(($MainForm.ClientSize.Width - 60), 15)
$helpBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$helpBtn.BackColor = $script:Colors.BackgroundLight
$helpBtn.ForeColor = $script:Colors.Text
$helpBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$helpBtn.Font = $script:Fonts.Small
$helpBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$helpBtn.Add_Click({
    [System.Windows.Forms.MessageBox]::Show(
        "Windows Troubleshooting Tool v1.0`n`n" +
        "A portable system diagnostics tool for Windows administrators.`n`n" +
        "Features:`n" +
        "- Quick health dashboard`n" +
        "- Category-based script browser`n" +
        "- Search and filter scripts`n" +
        "- Export results to TXT/HTML`n" +
        "- Admin elevation support`n`n" +
        "Keyboard Shortcuts:`n" +
        "Ctrl+F - Focus search`n" +
        "Ctrl+R - Run selected script`n" +
        "Ctrl+H - Quick health check`n" +
        "Ctrl+E - Export results`n" +
        "Ctrl+C - Copy output`n`n" +
        "Scripts folder: $script:ScriptsFolder",
        "About",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})
$headerPanel.Controls.Add($helpBtn)

$MainForm.Controls.Add($headerPanel)

# ============================================================================
# TAB CONTROL
# ============================================================================

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabControl.Font = $script:Fonts.Normal
$tabControl.Padding = New-Object System.Drawing.Point(15, 5)

# Create tabs for each category
$script:OutputBoxes = @{}
$script:ScriptLists = @{}
$script:DashboardStatusLabels = @{}

foreach ($categoryName in $script:Categories.Keys) {
    $category = $script:Categories[$categoryName]
    
    $tabPage = New-Object System.Windows.Forms.TabPage
    $tabPage.Text = "$($category.Icon) $categoryName"
    $tabPage.BackColor = $script:Colors.Background
    $tabPage.Padding = New-Object System.Windows.Forms.Padding(10)
    
    if ($categoryName -eq "Dashboard") {
        # ============================================================
        # DASHBOARD TAB - Special layout
        # ============================================================
        
        $dashPanel = New-Object System.Windows.Forms.TableLayoutPanel
        $dashPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
        $dashPanel.ColumnCount = 2
        $dashPanel.RowCount = 2
        $dashPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
        $dashPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 60)))
        $dashPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 200)))
        $dashPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
        
        # Status cards panel
        $cardsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
        $cardsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
        $cardsPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
        $cardsPanel.WrapContents = $false
        $cardsPanel.AutoScroll = $true
        $cardsPanel.Padding = New-Object System.Windows.Forms.Padding(5)
        
        $statusItems = @("UPTIME", "CPU", "RAM", "STORAGE", "NETWORK")
        foreach ($item in $statusItems) {
            $cardPanel = New-Object System.Windows.Forms.Panel
            $cardPanel.Size = New-Object System.Drawing.Size(280, 30)
            $cardPanel.BackColor = $script:Colors.BackgroundLight
            $cardPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 5)
            
            $itemLabel = New-Object System.Windows.Forms.Label
            $itemLabel.Text = $item
            $itemLabel.Font = $script:Fonts.Small
            $itemLabel.ForeColor = $script:Colors.TextDim
            $itemLabel.Location = New-Object System.Drawing.Point(10, 7)
            $itemLabel.AutoSize = $true
            $cardPanel.Controls.Add($itemLabel)
            
            $statusLabel = New-Object System.Windows.Forms.Label
            $statusLabel.Text = "---"
            $statusLabel.Font = $script:Fonts.Small
            $statusLabel.ForeColor = $script:Colors.Text
            $statusLabel.Location = New-Object System.Drawing.Point(100, 7)
            $statusLabel.AutoSize = $true
            $cardPanel.Controls.Add($statusLabel)
            
            $script:DashboardStatusLabels[$item] = $statusLabel
            $cardsPanel.Controls.Add($cardPanel)
        }
        
        # Run Health Check button
        $healthBtn = New-Object System.Windows.Forms.Button
        $healthBtn.Text = "Run Quick Health Check"
        $healthBtn.Size = New-Object System.Drawing.Size(280, 40)
        $healthBtn.BackColor = $script:Colors.Accent
        $healthBtn.ForeColor = $script:Colors.Text
        $healthBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $healthBtn.Font = $script:Fonts.Header
        $healthBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
        $healthBtn.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)
        $cardsPanel.Controls.Add($healthBtn)
        
        $dashPanel.Controls.Add($cardsPanel, 0, 0)
        $dashPanel.SetRowSpan($cardsPanel, 2)
        
        # Output panel for dashboard
        $dashOutputBox = New-Object System.Windows.Forms.TextBox
        $dashOutputBox.Multiline = $true
        $dashOutputBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $dashOutputBox.Dock = [System.Windows.Forms.DockStyle]::Fill
        $dashOutputBox.BackColor = $script:Colors.BackgroundDark
        $dashOutputBox.ForeColor = $script:Colors.Text
        $dashOutputBox.Font = $script:Fonts.Mono
        $dashOutputBox.ReadOnly = $true
        $dashOutputBox.WordWrap = $true
        $dashOutputBox.Text = "Click 'Run Quick Health Check' to analyze your system.`r`n`r`nThis will check:`r`n- System uptime and pending reboots`r`n- CPU status and temperature`r`n- RAM usage and modules`r`n- Disk space on all drives`r`n- Network connectivity"
        
        $dashPanel.Controls.Add($dashOutputBox, 1, 0)
        $dashPanel.SetRowSpan($dashOutputBox, 2)
        
        $script:OutputBoxes["Dashboard"] = $dashOutputBox
        
        $tabPage.Controls.Add($dashPanel)
        
        # Wire up health check button
        $healthBtn.Add_Click({
            Run-QuickHealthCheck -OutputBox $script:OutputBoxes["Dashboard"] -StatusLabel $statusBarLabel -StatusLabels $script:DashboardStatusLabels
        }.GetNewClosure())
    }
    else {
        # ============================================================
        # CATEGORY TABS - Script list layout
        # ============================================================
        
        $splitContainer = New-Object System.Windows.Forms.SplitContainer
        $splitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
        $splitContainer.Orientation = [System.Windows.Forms.Orientation]::Vertical
        $splitContainer.SplitterDistance = 350
        $splitContainer.BackColor = $script:Colors.Background
        $splitContainer.Panel1.BackColor = $script:Colors.Background
        $splitContainer.Panel2.BackColor = $script:Colors.Background
        
        # LEFT PANEL - Search and script list
        $leftPanel = New-Object System.Windows.Forms.Panel
        $leftPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
        $leftPanel.Padding = New-Object System.Windows.Forms.Padding(5)
        
        # Search box
        $searchBox = New-Object System.Windows.Forms.TextBox
        $searchBox.Dock = [System.Windows.Forms.DockStyle]::Top
        $searchBox.Height = 28
        $searchBox.BackColor = $script:Colors.BackgroundLight
        $searchBox.ForeColor = $script:Colors.Text
        $searchBox.Font = $script:Fonts.Normal
        $searchBox.Text = "Search scripts..."
        $searchBox.ForeColor = $script:Colors.TextDim
        
        # Search placeholder behavior
        $searchBox.Add_GotFocus({
            if ($this.Text -eq "Search scripts...") {
                $this.Text = ""
                $this.ForeColor = $script:Colors.Text
            }
        })
        $searchBox.Add_LostFocus({
            if ($this.Text -eq "") {
                $this.Text = "Search scripts..."
                $this.ForeColor = $script:Colors.TextDim
            }
        })
        
        # Script list
        $scriptList = New-Object System.Windows.Forms.ListView
        $scriptList.Dock = [System.Windows.Forms.DockStyle]::Fill
        $scriptList.View = [System.Windows.Forms.View]::Details
        $scriptList.FullRowSelect = $true
        $scriptList.GridLines = $false
        $scriptList.BackColor = $script:Colors.BackgroundDark
        $scriptList.ForeColor = $script:Colors.Text
        $scriptList.Font = $script:Fonts.Small
        $scriptList.BorderStyle = [System.Windows.Forms.BorderStyle]::None
        $scriptList.Columns.Add("Script", 150)
        $scriptList.Columns.Add("Description", 300)
        $scriptList.Columns.Add("Admin", 50)
        
        # Populate scripts
        $scriptsToAdd = if ($categoryName -eq "All Scripts") { Get-AllScripts } else { $category.Scripts }
        foreach ($scriptInfo in $scriptsToAdd) {
            $item = New-Object System.Windows.Forms.ListViewItem($scriptInfo.Name)
            $item.SubItems.Add($scriptInfo.Description)
            $item.SubItems.Add($(if ($scriptInfo.Admin) { "Yes" } else { "" }))
            $item.Tag = $scriptInfo
            if ($scriptInfo.Admin) {
                $item.ForeColor = $script:Colors.Warning
            }
            $scriptList.Items.Add($item)
        }
        
        $script:ScriptLists[$categoryName] = $scriptList
        
        # Search filtering
        $searchBox.Add_TextChanged({
            param($sender, $e)
            $searchText = $sender.Text.ToLower()
            if ($searchText -eq "search scripts..." -or $searchText -eq "") { return }
            
            $listView = $script:ScriptLists[$tabControl.SelectedTab.Text.Substring(2)]
            if ($null -eq $listView) { return }
            
            $listView.BeginUpdate()
            foreach ($item in $listView.Items) {
                $visible = $item.Text.ToLower().Contains($searchText) -or 
                           $item.SubItems[1].Text.ToLower().Contains($searchText)
                # Note: ListView doesn't support hiding items directly, 
                # so we change the color to indicate matches
                if ($visible) {
                    $item.BackColor = $script:Colors.BackgroundLight
                } else {
                    $item.BackColor = $script:Colors.BackgroundDark
                }
            }
            $listView.EndUpdate()
        }.GetNewClosure())
        
        # Run button
        $runBtn = New-Object System.Windows.Forms.Button
        $runBtn.Text = "Run Selected Script"
        $runBtn.Dock = [System.Windows.Forms.DockStyle]::Bottom
        $runBtn.Height = 35
        $runBtn.BackColor = $script:Colors.Accent
        $runBtn.ForeColor = $script:Colors.Text
        $runBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $runBtn.Font = $script:Fonts.Normal
        $runBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
        
        $leftPanel.Controls.Add($scriptList)
        $leftPanel.Controls.Add($searchBox)
        $leftPanel.Controls.Add($runBtn)
        
        $splitContainer.Panel1.Controls.Add($leftPanel)
        
        # RIGHT PANEL - Output
        $outputBox = New-Object System.Windows.Forms.TextBox
        $outputBox.Multiline = $true
        $outputBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
        $outputBox.Dock = [System.Windows.Forms.DockStyle]::Fill
        $outputBox.BackColor = $script:Colors.BackgroundDark
        $outputBox.ForeColor = $script:Colors.Text
        $outputBox.Font = $script:Fonts.Mono
        $outputBox.ReadOnly = $true
        $outputBox.WordWrap = $true
        $outputBox.Text = "Select a script and click 'Run Selected Script' to see output here.`r`n`r`nDouble-click a script to run it immediately."
        
        $script:OutputBoxes[$categoryName] = $outputBox
        
        # Output toolbar
        $outputToolbar = New-Object System.Windows.Forms.Panel
        $outputToolbar.Dock = [System.Windows.Forms.DockStyle]::Bottom
        $outputToolbar.Height = 35
        $outputToolbar.BackColor = $script:Colors.BackgroundLight
        
        $copyBtn = New-Object System.Windows.Forms.Button
        $copyBtn.Text = "Copy"
        $copyBtn.Size = New-Object System.Drawing.Size(70, 25)
        $copyBtn.Location = New-Object System.Drawing.Point(5, 5)
        $copyBtn.BackColor = $script:Colors.BackgroundDark
        $copyBtn.ForeColor = $script:Colors.Text
        $copyBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $copyBtn.Font = $script:Fonts.Small
        $copyBtn.Add_Click({
            $catName = $tabControl.SelectedTab.Text.Substring(2)
            if ($script:OutputBoxes.ContainsKey($catName)) {
                [System.Windows.Forms.Clipboard]::SetText($script:OutputBoxes[$catName].Text)
            }
        }.GetNewClosure())
        $outputToolbar.Controls.Add($copyBtn)
        
        $clearBtn = New-Object System.Windows.Forms.Button
        $clearBtn.Text = "Clear"
        $clearBtn.Size = New-Object System.Drawing.Size(70, 25)
        $clearBtn.Location = New-Object System.Drawing.Point(80, 5)
        $clearBtn.BackColor = $script:Colors.BackgroundDark
        $clearBtn.ForeColor = $script:Colors.Text
        $clearBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $clearBtn.Font = $script:Fonts.Small
        $clearBtn.Add_Click({
            $catName = $tabControl.SelectedTab.Text.Substring(2)
            if ($script:OutputBoxes.ContainsKey($catName)) {
                $script:OutputBoxes[$catName].Clear()
            }
        }.GetNewClosure())
        $outputToolbar.Controls.Add($clearBtn)
        
        $splitContainer.Panel2.Controls.Add($outputBox)
        $splitContainer.Panel2.Controls.Add($outputToolbar)
        
        $tabPage.Controls.Add($splitContainer)
        
        # Wire up run button
        $runBtn.Add_Click({
            $catName = $tabControl.SelectedTab.Text.Substring(2)
            $list = $script:ScriptLists[$catName]
            $output = $script:OutputBoxes[$catName]
            
            if ($list.SelectedItems.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Please select a script to run.", "No Selection", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                return
            }
            
            $scriptInfo = $list.SelectedItems[0].Tag
            if ($scriptInfo.Admin -and -not $script:IsAdmin) {
                if (-not (Request-AdminElevation)) { return }
            }
            
            $scriptPath = Get-ScriptPath $scriptInfo.Name
            if ($scriptPath) {
                Invoke-ScriptAsync -ScriptPath $scriptPath -OutputBox $output -StatusLabel $statusBarLabel
            } else {
                $output.AppendText("`r`n[ERROR] Script not found: $($scriptInfo.Name).ps1`r`n")
            }
        }.GetNewClosure())
        
        # Double-click to run
        $scriptList.Add_DoubleClick({
            $catName = $tabControl.SelectedTab.Text.Substring(2)
            $list = $script:ScriptLists[$catName]
            $output = $script:OutputBoxes[$catName]
            
            if ($list.SelectedItems.Count -eq 0) { return }
            
            $scriptInfo = $list.SelectedItems[0].Tag
            if ($scriptInfo.Admin -and -not $script:IsAdmin) {
                if (-not (Request-AdminElevation)) { return }
            }
            
            $scriptPath = Get-ScriptPath $scriptInfo.Name
            if ($scriptPath) {
                Invoke-ScriptAsync -ScriptPath $scriptPath -OutputBox $output -StatusLabel $statusBarLabel
            }
        }.GetNewClosure())
    }
    
    $tabControl.TabPages.Add($tabPage)
}

# ============================================================================
# STATUS BAR
# ============================================================================

$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = $script:Colors.BackgroundDark
$statusBar.ForeColor = $script:Colors.Text

$statusBarLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusBarLabel.Text = "Ready"
$statusBarLabel.Spring = $true
$statusBarLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusBar.Items.Add($statusBarLabel)

$scriptsCountLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$scriptCount = if (Test-Path $script:ScriptsFolder) { 
    (Get-ChildItem -Path $script:ScriptsFolder -Filter "*.ps1" -ErrorAction SilentlyContinue | Measure-Object).Count 
} else { 0 }
$scriptsCountLabel.Text = "Scripts: $scriptCount"
$statusBar.Items.Add($scriptsCountLabel)

$MainForm.Controls.Add($statusBar)

# Add main content panel to hold tab control
$mainContent = New-Object System.Windows.Forms.Panel
$mainContent.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainContent.Controls.Add($tabControl)
$MainForm.Controls.Add($mainContent)

# ============================================================================
# KEYBOARD SHORTCUTS
# ============================================================================

$MainForm.KeyPreview = $true
$MainForm.Add_KeyDown({
    param($sender, $e)
    
    if ($e.Control) {
        switch ($e.KeyCode) {
            "F" {
                # Focus search
                $catName = $tabControl.SelectedTab.Text.Substring(2)
                # Try to find search box in current tab
                $e.Handled = $true
            }
            "R" {
                # Run selected
                $catName = $tabControl.SelectedTab.Text.Substring(2)
                if ($script:ScriptLists.ContainsKey($catName) -and $script:ScriptLists[$catName].SelectedItems.Count -gt 0) {
                    $scriptInfo = $script:ScriptLists[$catName].SelectedItems[0].Tag
                    $scriptPath = Get-ScriptPath $scriptInfo.Name
                    if ($scriptPath) {
                        Invoke-ScriptAsync -ScriptPath $scriptPath -OutputBox $script:OutputBoxes[$catName] -StatusLabel $statusBarLabel
                    }
                }
                $e.Handled = $true
            }
            "H" {
                # Health check
                $tabControl.SelectedIndex = 0
                Run-QuickHealthCheck -OutputBox $script:OutputBoxes["Dashboard"] -StatusLabel $statusBarLabel -StatusLabels $script:DashboardStatusLabels
                $e.Handled = $true
            }
            "E" {
                # Export
                $catName = $tabControl.SelectedTab.Text.Substring(2)
                if ($script:OutputBoxes.ContainsKey($catName)) {
                    Export-Results -Content $script:OutputBoxes[$catName].Text -Format "txt"
                }
                $e.Handled = $true
            }
        }
    }
})

# ============================================================================
# EXPORT BUTTON HANDLER
# ============================================================================

$exportBtn.Add_Click({
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    
    $txtItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $txtItem.Text = "Export as TXT"
    $txtItem.Add_Click({
        $catName = $tabControl.SelectedTab.Text.Substring(2)
        if ($script:OutputBoxes.ContainsKey($catName)) {
            Export-Results -Content $script:OutputBoxes[$catName].Text -Format "txt"
        }
    })
    $contextMenu.Items.Add($txtItem)
    
    $htmlItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $htmlItem.Text = "Export as HTML"
    $htmlItem.Add_Click({
        $catName = $tabControl.SelectedTab.Text.Substring(2)
        if ($script:OutputBoxes.ContainsKey($catName)) {
            Export-Results -Content $script:OutputBoxes[$catName].Text -Format "html"
        }
    })
    $contextMenu.Items.Add($htmlItem)
    
    $contextMenu.Show($exportBtn, [System.Drawing.Point]::new(0, $exportBtn.Height))
})

# ============================================================================
# STARTUP CHECK
# ============================================================================

if (-not (Test-Path $script:ScriptsFolder)) {
    $statusBarLabel.Text = "Warning: Scripts folder not found at $script:ScriptsFolder"
    $statusBarLabel.ForeColor = $script:Colors.Warning
}

# ============================================================================
# RUN APPLICATION
# ============================================================================

[System.Windows.Forms.Application]::Run($MainForm)
