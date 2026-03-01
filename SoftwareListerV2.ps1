#Requires -Version 5.1
<#
.SYNOPSIS
    SoftwareLister - Comprehensive system inventory tool for Windows
.DESCRIPTION
    Scans and lists all installed software from:
    - Windows Registry (traditional programs)
    - AppX packages (Microsoft Store / UWP apps)
    - Winget (Windows Package Manager)

    Additional tabs:
    - Driver Manager: Query, backup, and import Windows drivers
    - Service Manager: (coming next)

    Features:
    - Export to CSV, HTML, JSON formats
    - Compare with previous scans to detect changes
    - Custom name overrides for better reporting
    - Configurable properties via config.json
.NOTES
    Author: SoftwareLister
    Version: 2.0
#>

param(
    [switch]$DebugMode
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Web

# Load shared report engine (export / import / compare)
. "$PSScriptRoot\ReportEngine.ps1"

#region Global Variables
$script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ConfigPath = Join-Path $script:ScriptPath "config.json"
$script:Config = $null
$script:SoftwareList = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
$script:ComparisonData = $null
$script:ComparisonSnapshots = @()  # Stores version history: @{Date, Data} for each compared scan
$script:Window = $null

# Driver Manager globals
$script:AllDevices = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:DrvC       = @{}

# Service Manager globals
$script:AllServices = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:AclCache    = @{}
$script:SvcC        = @{}

# Debug mode (set by -DebugMode switch at launch)
$script:DebugMode = $DebugMode.IsPresent
#endregion

# Writes a timestamped debug line to the terminal when -DebugMode is active.
# Safe to call from any thread (uses Console.WriteLine which is thread-safe).
function Write-DebugLine {
    param([string]$Tag, [string]$Message)
    if ($script:DebugMode) {
        $ts = (Get-Date).ToString('HH:mm:ss.fff')
        [Console]::WriteLine("[$ts][DBG][$Tag] $Message")
    }
}

#region Configuration Functions
function Load-Configuration {
    if (Test-Path $script:ConfigPath) {
        try {
            $script:Config = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
            return $true
        } catch {
            [System.Windows.MessageBox]::Show("Error loading config.json: $($_.Exception.Message)", "Configuration Error", "OK", "Error")
            return $false
        }
    } else {
        [System.Windows.MessageBox]::Show("config.json not found. Please ensure it exists in the same folder as the script.", "Configuration Error", "OK", "Error")
        return $false
    }
}

function Save-Configuration {
    try {
        $script:Config | ConvertTo-Json -Depth 10 | Set-Content $script:ConfigPath -Encoding UTF8
        return $true
    } catch {
        [System.Windows.MessageBox]::Show("Error saving config.json: $($_.Exception.Message)", "Save Error", "OK", "Error")
        return $false
    }
}

function Get-EnabledProperties {
    $enabled = @()
    foreach ($prop in $script:Config.properties.PSObject.Properties) {
        if ($prop.Value.enabled -eq $true) {
            $enabled += $prop.Name
        }
    }
    return $enabled
}

#endregion

#region Theme Functions
$script:ThemeColors = @{}

# Add Windows API for dark title bar
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class DwmApi {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

    public const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

    public static void SetDarkTitleBar(IntPtr hwnd, bool enabled) {
        int value = enabled ? 1 : 0;
        DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));
    }
}
"@ -ErrorAction SilentlyContinue

function Get-WindowsTheme {
    try {
        $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $value = Get-ItemPropertyValue -Path $regPath -Name "AppsUseLightTheme" -ErrorAction Stop
        if ($value -eq 0) { return "dark" }
        else { return "light" }
    } catch {
        return "light"
    }
}

function Get-CurrentTheme {
    $themeSetting = $script:Config.display.theme
    if (-not $themeSetting) { $themeSetting = "system" }

    switch ($themeSetting.ToLower()) {
        "dark" { return "dark" }
        "light" { return "light" }
        default { return Get-WindowsTheme }
    }
}

function Get-ThemeColors {
    param([string]$Theme)

    $bc = [System.Windows.Media.BrushConverter]::new()

    if ($Theme -eq "dark") {
        return @{
            WindowBackground = $bc.ConvertFrom("#1e1e1e")
            PanelBackground = $bc.ConvertFrom("#252526")
            CardBackground = $bc.ConvertFrom("#2d2d30")
            BorderColor = $bc.ConvertFrom("#3f3f46")
            TextPrimary = $bc.ConvertFrom("#e0e0e0")
            TextSecondary = $bc.ConvertFrom("#a0a0a0")
            AccentColor = $bc.ConvertFrom("#0078d4")
            AccentHover = $bc.ConvertFrom("#1a8cde")
            RowBackground = $bc.ConvertFrom("#2d2d30")
            RowAlternate = $bc.ConvertFrom("#333337")
            HeaderBackground = $bc.ConvertFrom("#0078d4")
            HeaderForeground = $bc.ConvertFrom("#ffffff")
            StatusBarBackground = $bc.ConvertFrom("#007acc")
            StatusBarForeground = $bc.ConvertFrom("#ffffff")
            ButtonBackground = $bc.ConvertFrom("#3c3c3c")
            ButtonForeground = $bc.ConvertFrom("#e0e0e0")
            ButtonBorder = $bc.ConvertFrom("#555555")
            InputBackground = $bc.ConvertFrom("#3c3c3c")
            InputForeground = $bc.ConvertFrom("#e0e0e0")
            InputBorder = $bc.ConvertFrom("#555555")
            NewRowBackground = $bc.ConvertFrom("#1e3a1e")
            RemovedRowBackground = $bc.ConvertFrom("#3a1e1e")
            UpdatedRowBackground = $bc.ConvertFrom("#3a3a1e")
        }
    } else {
        return @{
            WindowBackground = $bc.ConvertFrom("#f5f5f5")
            PanelBackground = $bc.ConvertFrom("#ffffff")
            CardBackground = $bc.ConvertFrom("#ffffff")
            BorderColor = $bc.ConvertFrom("#e0e0e0")
            TextPrimary = $bc.ConvertFrom("#1e1e1e")
            TextSecondary = $bc.ConvertFrom("#666666")
            AccentColor = $bc.ConvertFrom("#0078d4")
            AccentHover = $bc.ConvertFrom("#106ebe")
            RowBackground = $bc.ConvertFrom("#ffffff")
            RowAlternate = $bc.ConvertFrom("#fafafa")
            HeaderBackground = $bc.ConvertFrom("#0078d4")
            HeaderForeground = $bc.ConvertFrom("#ffffff")
            StatusBarBackground = $bc.ConvertFrom("#333333")
            StatusBarForeground = $bc.ConvertFrom("#ffffff")
            ButtonBackground = $bc.ConvertFrom("#ffffff")
            ButtonForeground = $bc.ConvertFrom("#333333")
            ButtonBorder = $bc.ConvertFrom("#cccccc")
            InputBackground = $bc.ConvertFrom("#ffffff")
            InputForeground = $bc.ConvertFrom("#1e1e1e")
            InputBorder = $bc.ConvertFrom("#cccccc")
            NewRowBackground = $bc.ConvertFrom("#d4edda")
            RemovedRowBackground = $bc.ConvertFrom("#f8d7da")
            UpdatedRowBackground = $bc.ConvertFrom("#fff3cd")
        }
    }
}

function Apply-Theme {
    param([System.Windows.Window]$Window)

    $theme = Get-CurrentTheme
    $script:ThemeColors = Get-ThemeColors -Theme $theme

    # Apply to window
    $Window.Background = $script:ThemeColors.WindowBackground

    # Set dark/light title bar (Windows 10/11)
    try {
        $windowHelper = New-Object System.Windows.Interop.WindowInteropHelper($Window)
        $hwnd = $windowHelper.Handle
        if ($hwnd -ne [IntPtr]::Zero) {
            [DwmApi]::SetDarkTitleBar($hwnd, ($theme -eq "dark"))
        }
    } catch {
        # Ignore if API not available (older Windows)
    }

    # Find and style controls - use named elements now that content is a TabControl
    $toolbarBorder = $Window.FindName("swToolbarBorder")
    $searchBorder  = $Window.FindName("swSearchBorder")
    $dataGrid      = $Window.FindName("dgSoftware")
    $statusBorder  = $Window.FindName("swStatusBorder")
    $statusText    = $Window.FindName("txtStatus")
    $searchText    = $Window.FindName("txtSearch")

    # Toolbar
    if ($toolbarBorder) {
        $toolbarBorder.Background = $script:ThemeColors.PanelBackground
        $toolbarBorder.BorderBrush = $script:ThemeColors.BorderColor
    }

    # Search bar
    if ($searchBorder) {
        $searchBorder.Background = $script:ThemeColors.CardBackground
        $searchGrid = $searchBorder.Child
        if ($searchGrid -and $searchGrid.Children) {
            foreach ($child in $searchGrid.Children) {
                if ($child -is [System.Windows.Controls.TextBlock]) {
                    $child.Foreground = $script:ThemeColors.TextPrimary
                }
            }
        }
    }

    # Search textbox
    if ($searchText) {
        $searchText.Background = $script:ThemeColors.InputBackground
        $searchText.Foreground = $script:ThemeColors.InputForeground
        $searchText.BorderBrush = $script:ThemeColors.InputBorder
    }

    # DataGrid
    if ($dataGrid) {
        $dataGrid.Background = $script:ThemeColors.CardBackground
        $dataGrid.BorderBrush = $script:ThemeColors.BorderColor
        $dataGrid.RowBackground = $script:ThemeColors.RowBackground
        $dataGrid.AlternatingRowBackground = $script:ThemeColors.RowAlternate
        $dataGrid.Foreground = $script:ThemeColors.TextPrimary

        # Style column headers
        $headerStyle = New-Object System.Windows.Style([System.Windows.Controls.Primitives.DataGridColumnHeader])
        $headerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::BackgroundProperty, $script:ThemeColors.HeaderBackground)))
        $headerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::ForegroundProperty, $script:ThemeColors.HeaderForeground)))
        $headerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::BorderBrushProperty, $script:ThemeColors.BorderColor)))
        $headerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::BorderThicknessProperty, [System.Windows.Thickness]::new(0,0,1,0))))
        $headerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::PaddingProperty, [System.Windows.Thickness]::new(8,6,8,6))))
        $headerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::FontWeightProperty, [System.Windows.FontWeights]::SemiBold)))
        $dataGrid.ColumnHeaderStyle = $headerStyle

        # Update row style for change status colors
        $rowStyle = New-Object System.Windows.Style([System.Windows.Controls.DataGridRow])

        $newTrigger = New-Object System.Windows.DataTrigger
        $newTrigger.Binding = New-Object System.Windows.Data.Binding("ChangeStatus")
        $newTrigger.Value = "NEW"
        $newTrigger.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.DataGridRow]::BackgroundProperty, $script:ThemeColors.NewRowBackground)))
        $rowStyle.Triggers.Add($newTrigger)

        $removedTrigger = New-Object System.Windows.DataTrigger
        $removedTrigger.Binding = New-Object System.Windows.Data.Binding("ChangeStatus")
        $removedTrigger.Value = "REMOVED"
        $removedTrigger.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.DataGridRow]::BackgroundProperty, $script:ThemeColors.RemovedRowBackground)))
        $rowStyle.Triggers.Add($removedTrigger)

        $dataGrid.RowStyle = $rowStyle

        $cellStyle = New-Object System.Windows.Style([System.Windows.Controls.DataGridCell])
        $cellStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.DataGridCell]::ForegroundProperty, $script:ThemeColors.TextPrimary)))
        $cellStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.DataGridCell]::BorderThicknessProperty, [System.Windows.Thickness]::new(0))))
        $dataGrid.CellStyle = $cellStyle
    }

    # Status bar
    if ($statusBorder) {
        $statusBorder.Background = $script:ThemeColors.StatusBarBackground
    }
    if ($statusText) {
        $statusText.Foreground = $script:ThemeColors.StatusBarForeground
    }

    # Style buttons
    foreach ($btn in @("btnScan", "btnCompare", "btnExportCsv", "btnExportHtml", "btnExportJson", "btnSettings", "btnClearSearch")) {
        $button = $Window.FindName($btn)
        if ($button) {
            if ($btn -eq "btnScan") {
                $button.Background = $script:ThemeColors.AccentColor
                $button.Foreground = $script:ThemeColors.HeaderForeground
            } else {
                $button.Background = $script:ThemeColors.ButtonBackground
                $button.Foreground = $script:ThemeColors.ButtonForeground
                $button.BorderBrush = $script:ThemeColors.ButtonBorder
            }
        }
    }
}
#endregion

#region Software Scanning Functions
function Get-RegistrySoftware {
    $software = @()

    $regPaths = @(
        @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"; Arch = "x64" },
        @{ Path = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"; Arch = "x86" },
        @{ Path = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"; Arch = "User" }
    )

    foreach ($regInfo in $regPaths) {
        $items = Get-ItemProperty $regInfo.Path -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            if (-not $item.DisplayName) { continue }

            if (-not $script:Config.display.showSystemComponents -and $item.SystemComponent -eq 1) { continue }

            $excluded = $false
            foreach ($pattern in $script:Config.display.excludePatterns) {
                if ($item.DisplayName -match $pattern) {
                    $excluded = $true
                    break
                }
            }
            if ($excluded) { continue }

            $installDate = ""
            if ($item.InstallDate) {
                try {
                    if ($item.InstallDate -match '^\d{8}$') {
                        $installDate = [datetime]::ParseExact($item.InstallDate, "yyyyMMdd", $null).ToString("yyyy-MM-dd")
                    } else {
                        $installDate = $item.InstallDate
                    }
                } catch { $installDate = $item.InstallDate }
            }

            $size = ""
            if ($item.EstimatedSize) {
                $sizeKB = [int]$item.EstimatedSize
                if ($sizeKB -gt 1024000) {
                    $size = "{0:N1} GB" -f ($sizeKB / 1024 / 1024)
                } elseif ($sizeKB -gt 1024) {
                    $size = "{0:N1} MB" -f ($sizeKB / 1024)
                } else {
                    $size = "$sizeKB KB"
                }
            }

            $uniqueId = ""
            if ($item.PSPath) {
                $uniqueId = Split-Path $item.PSPath -Leaf
            }

            $software += [PSCustomObject]@{
                Name = $item.DisplayName
                CustomName = ""
                Version = if ($item.DisplayVersion) { $item.DisplayVersion } else { "" }
                Publisher = if ($item.Publisher) { $item.Publisher } else { "" }
                InstallDate = $installDate
                Source = "Registry"
                InstallLocation = if ($item.InstallLocation) { $item.InstallLocation } else { "" }
                Architecture = $regInfo.Arch
                Size = $size
                UniqueId = $uniqueId
                IsSystemComponent = ($item.SystemComponent -eq 1)
                IsFramework = $false
                UninstallString = if ($item.UninstallString) { $item.UninstallString } else { "" }
                HelpLink = if ($item.HelpLink) { $item.HelpLink } else { "" }
                Comments = if ($item.Comments) { $item.Comments } else { "" }
                ChangeStatus = ""
            }
        }
    }

    return $software
}

function Get-AppxSoftware {
    $software = @()

    try {
        $packages = Get-AppxPackage -ErrorAction Stop

        foreach ($pkg in $packages) {
            if (-not $script:Config.display.showFrameworks -and $pkg.IsFramework) { continue }

            $displayName = $pkg.Name
            try {
                $manifestPath = Join-Path $pkg.InstallLocation "AppxManifest.xml"
                if (Test-Path $manifestPath) {
                    [xml]$manifest = Get-Content $manifestPath -ErrorAction SilentlyContinue
                    if ($manifest.Package.Properties.DisplayName -and
                        $manifest.Package.Properties.DisplayName -notmatch '^ms-resource:') {
                        $displayName = $manifest.Package.Properties.DisplayName
                    }
                }
            } catch { }

            $excluded = $false
            foreach ($pattern in $script:Config.display.excludePatterns) {
                if ($displayName -match $pattern -or $pkg.Name -match $pattern) {
                    $excluded = $true
                    break
                }
            }
            if ($excluded) { continue }

            $publisher = $pkg.Publisher
            if ($publisher -match 'O=([^,]+)') {
                $publisher = $Matches[1]
            } elseif ($publisher -match 'CN=([^,]+)') {
                $publisher = $Matches[1]
            }

            $software += [PSCustomObject]@{
                Name = $displayName
                CustomName = ""
                Version = $pkg.Version.ToString()
                Publisher = $publisher
                InstallDate = ""
                Source = "AppX"
                InstallLocation = $pkg.InstallLocation
                Architecture = $pkg.Architecture.ToString()
                Size = ""
                UniqueId = $pkg.PackageFamilyName
                IsSystemComponent = $pkg.NonRemovable
                IsFramework = $pkg.IsFramework
                UninstallString = ""
                HelpLink = ""
                Comments = ""
                ChangeStatus = ""
            }
        }
    } catch {
        Write-Host "Error getting AppX packages: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    return $software
}

function Get-WingetSoftware {
    $software = @()

    $wingetPath = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $wingetPath) {
        return $software
    }

    try {
        $output = & winget list --disable-interactivity 2>$null

        if (-not $output) { return $software }

        $headerIndex = -1
        for ($i = 0; $i -lt $output.Count; $i++) {
            if ($output[$i] -match '^Name\s+Id\s+Version') {
                $headerIndex = $i
                break
            }
        }

        if ($headerIndex -lt 0) { return $software }

        $separatorLine = $output[$headerIndex + 1]
        if ($separatorLine -notmatch '^-+') {
            $separatorLine = $output[$headerIndex]
        }

        $headerLine = $output[$headerIndex]
        $nameStart = 0
        $idStart = $headerLine.IndexOf("Id")
        $versionStart = $headerLine.IndexOf("Version")
        $availableStart = $headerLine.IndexOf("Available")
        $sourceStart = $headerLine.IndexOf("Source")

        for ($i = $headerIndex + 2; $i -lt $output.Count; $i++) {
            $line = $output[$i]
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if ($line -match '^-+$') { continue }

            $name = ""
            $id = ""
            $version = ""

            if ($line.Length -gt $idStart) {
                $name = $line.Substring($nameStart, [Math]::Min($idStart - $nameStart, $line.Length)).Trim()
            }
            if ($line.Length -gt $versionStart -and $idStart -gt 0) {
                $id = $line.Substring($idStart, [Math]::Min($versionStart - $idStart, $line.Length - $idStart)).Trim()
            }
            if ($sourceStart -gt 0 -and $line.Length -gt $sourceStart) {
                $version = $line.Substring($versionStart, [Math]::Min($sourceStart - $versionStart, $line.Length - $versionStart)).Trim()
            } elseif ($availableStart -gt 0 -and $line.Length -gt $availableStart) {
                $version = $line.Substring($versionStart, [Math]::Min($availableStart - $versionStart, $line.Length - $versionStart)).Trim()
            } else {
                $version = $line.Substring($versionStart).Trim()
            }

            if ([string]::IsNullOrWhiteSpace($name)) { continue }

            $excluded = $false
            foreach ($pattern in $script:Config.display.excludePatterns) {
                if ($name -match $pattern -or $id -match $pattern) {
                    $excluded = $true
                    break
                }
            }
            if ($excluded) { continue }

            $software += [PSCustomObject]@{
                Name = $name
                CustomName = ""
                Version = $version
                Publisher = ""
                InstallDate = ""
                Source = "Winget"
                InstallLocation = ""
                Architecture = ""
                Size = ""
                UniqueId = $id
                IsSystemComponent = $false
                IsFramework = $false
                UninstallString = ""
                HelpLink = ""
                Comments = ""
                ChangeStatus = ""
            }
        }
    } catch {
        Write-Host "Error getting Winget packages: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    return $software
}

function Invoke-SoftwareScan {
    param([System.Windows.Controls.ProgressBar]$ProgressBar, [System.Windows.Controls.TextBlock]$StatusText)

    $script:SoftwareList.Clear()
    $script:ComparisonSnapshots = @()
    $allSoftware = @()

    if ($script:Config.dataSources.registry.enabled) {
        if ($StatusText) { $StatusText.Text = "Scanning Windows Registry..." }
        if ($ProgressBar) { $ProgressBar.Value = 10 }
        [System.Windows.Forms.Application]::DoEvents()

        $regSoftware = Get-RegistrySoftware
        $allSoftware += $regSoftware
    }

    if ($script:Config.dataSources.appx.enabled) {
        if ($StatusText) { $StatusText.Text = "Scanning AppX packages..." }
        if ($ProgressBar) { $ProgressBar.Value = 40 }
        [System.Windows.Forms.Application]::DoEvents()

        $appxSoftware = Get-AppxSoftware
        $allSoftware += $appxSoftware
    }

    if ($script:Config.dataSources.winget.enabled) {
        if ($StatusText) { $StatusText.Text = "Scanning Winget packages..." }
        if ($ProgressBar) { $ProgressBar.Value = 70 }
        [System.Windows.Forms.Application]::DoEvents()

        $wingetSoftware = Get-WingetSoftware
        $allSoftware += $wingetSoftware
    }

    if ($StatusText) { $StatusText.Text = "Applying custom names..." }
    if ($ProgressBar) { $ProgressBar.Value = 90 }
    [System.Windows.Forms.Application]::DoEvents()

    foreach ($item in $allSoftware) {
        $key = "$($item.Source)::$($item.UniqueId)"
        if (-not $key -or $key -eq "::") {
            $key = "$($item.Source)::$($item.Name)"
        }

        if ($script:Config.customNames.PSObject.Properties[$key]) {
            $item.CustomName = $script:Config.customNames.$key
        }
    }

    $allSoftware = $allSoftware | Sort-Object { if ($_.CustomName) { $_.CustomName } else { $_.Name } }

    foreach ($item in $allSoftware) {
        $script:SoftwareList.Add($item)
    }

    if ($StatusText) { $StatusText.Text = "Scan complete. Found $($script:SoftwareList.Count) items." }
    if ($ProgressBar) { $ProgressBar.Value = 100 }

    return $script:SoftwareList.Count
}
#endregion

#region Comparison Functions
function Get-DateFromFilename {
    param([string]$FilePath)

    $filename = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    if ($filename -match '_(\d{8})$') {
        try {
            return [datetime]::ParseExact($Matches[1], "yyyyMMdd", $null).ToString("yyyy-MM-dd")
        } catch { }
    }
    return (Get-Item $FilePath).LastWriteTime.ToString("yyyy-MM-dd")
}

function Compare-WithPrevious {
    param([string]$PreviousFilePath)

    if (-not (Test-Path $PreviousFilePath)) {
        [System.Windows.MessageBox]::Show("Previous file not found.", "Error", "OK", "Error")
        return
    }

    try {
        $extension = [System.IO.Path]::GetExtension($PreviousFilePath).ToLower()
        $previousData = @()

        if ($extension -eq ".csv") {
            $previousData = Import-Csv $PreviousFilePath
        } elseif ($extension -eq ".json") {
            $previousData = Get-Content $PreviousFilePath -Raw | ConvertFrom-Json
        } else {
            [System.Windows.MessageBox]::Show("Unsupported file format. Use CSV or JSON.", "Error", "OK", "Error")
            return
        }

        $previousDate = Get-DateFromFilename -FilePath $PreviousFilePath

        $previousLookup = @{}
        foreach ($item in $previousData) {
            $displayName = if ($item.CustomName) { $item.CustomName } else { $item.Name }
            $key = "$($item.Source)::$displayName"
            $previousLookup[$key] = $item
        }

        $script:ComparisonSnapshots = @(
            @{
                Date = $previousDate
                Data = $previousLookup
            }
        )

        $currentKeys = @{}
        foreach ($item in $script:SoftwareList) {
            $displayName = if ($item.CustomName) { $item.CustomName } else { $item.Name }
            $key = "$($item.Source)::$displayName"
            $currentKeys[$key] = $true

            if ($previousLookup.ContainsKey($key)) {
                $prev = $previousLookup[$key]
                $item | Add-Member -NotePropertyName "PreviousVersion" -NotePropertyValue $prev.Version -Force
                if ($item.Version -ne $prev.Version) {
                    $item.ChangeStatus = "UPDATED"
                } else {
                    $item.ChangeStatus = ""
                }
            } else {
                $item | Add-Member -NotePropertyName "PreviousVersion" -NotePropertyValue "" -Force
                $item.ChangeStatus = "NEW"
            }
        }

        foreach ($key in $previousLookup.Keys) {
            if (-not $currentKeys.ContainsKey($key)) {
                $prev = $previousLookup[$key]
                $removed = [PSCustomObject]@{
                    Name = $prev.Name
                    CustomName = if ($prev.CustomName) { $prev.CustomName } else { "" }
                    Version = ""
                    PreviousVersion = $prev.Version
                    Publisher = if ($prev.Publisher) { $prev.Publisher } else { "" }
                    InstallDate = if ($prev.InstallDate) { $prev.InstallDate } else { "" }
                    Source = $prev.Source
                    InstallLocation = ""
                    Architecture = if ($prev.Architecture) { $prev.Architecture } else { "" }
                    Size = ""
                    UniqueId = if ($prev.UniqueId) { $prev.UniqueId } else { "" }
                    IsSystemComponent = $false
                    IsFramework = $false
                    UninstallString = ""
                    HelpLink = ""
                    Comments = ""
                    ChangeStatus = "REMOVED"
                }
                $script:SoftwareList.Add($removed)
            }
        }

        $temp = @($script:SoftwareList)
        $script:SoftwareList.Clear()
        $sorted = $temp | Sort-Object @{Expression={
            switch ($_.ChangeStatus) {
                "NEW" { 0 }
                "UPDATED" { 1 }
                "REMOVED" { 2 }
                default { 3 }
            }
        }}, { if ($_.CustomName) { $_.CustomName } else { $_.Name } }

        foreach ($item in $sorted) {
            $script:SoftwareList.Add($item)
        }

        $newCount = ($script:SoftwareList | Where-Object { $_.ChangeStatus -eq "NEW" }).Count
        $updatedCount = ($script:SoftwareList | Where-Object { $_.ChangeStatus -eq "UPDATED" }).Count
        $removedCount = ($script:SoftwareList | Where-Object { $_.ChangeStatus -eq "REMOVED" }).Count

        [System.Windows.MessageBox]::Show(
            "Comparison complete.`n`nNew: $newCount`nUpdated: $updatedCount`nRemoved: $removedCount",
            "Comparison Complete", "OK", "Information")

    } catch {
        [System.Windows.MessageBox]::Show("Error comparing files: $($_.Exception.Message)", "Error", "OK", "Error")
    }
}
#endregion

#region Export Functions
function Export-ToCsv {
    param([string]$FilePath)

    $enabledProps = Get-EnabledProperties
    $exportData = $script:SoftwareList | Select-Object $enabledProps
    $exportData | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
}

function Export-ToJson {
    param([string]$FilePath)

    $enabledProps = Get-EnabledProperties
    $exportData = $script:SoftwareList | Select-Object $enabledProps
    $exportData | ConvertTo-Json -Depth 5 | Set-Content $FilePath -Encoding UTF8
}

function Export-ToHtml {
    param([string]$FilePath)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $currentDate = Get-Date -Format "yyyy-MM-dd"
    $computerName = $env:COMPUTERNAME

    $hasComparison = $script:ComparisonSnapshots.Count -gt 0
    $previousDate = if ($hasComparison) { $script:ComparisonSnapshots[0].Date } else { $null }

    $newCount = ($script:SoftwareList | Where-Object { $_.ChangeStatus -eq "NEW" }).Count
    $updatedCount = ($script:SoftwareList | Where-Object { $_.ChangeStatus -eq "UPDATED" }).Count
    $removedCount = ($script:SoftwareList | Where-Object { $_.ChangeStatus -eq "REMOVED" }).Count
    $unchangedCount = ($script:SoftwareList | Where-Object { -not $_.ChangeStatus }).Count

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Software Inventory - $computerName</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0; padding: 20px;
            background: #f5f5f5;
        }
        .container { display: inline-block; min-width: 100%; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #333; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        .meta { color: #666; margin-bottom: 20px; }
        .meta span { margin-right: 20px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; table-layout: auto; border: 1px solid #dee2e6; }
        th { background: #0078d4; color: white; padding: 12px 8px; text-align: left; position: sticky; top: 0; border: 1px solid #0078d4; white-space: nowrap; }
        th.date-col { background: #5c6bc0; min-width: 120px; }
        td { padding: 10px 8px; border: 1px solid #dee2e6; word-wrap: break-word; }
        td.version { font-family: 'Consolas', 'Courier New', monospace; text-align: center; }
        tr:nth-child(even) { background: #f8f9fa; }
        tr:hover { background: #e7f3ff; }
        tr.new td { background: #d4edda; }
        tr.updated td { background: #fff3cd; }
        tr.removed td { background: #f8d7da; }
        tr.removed td.name { text-decoration: line-through; }
        .status { font-weight: bold; text-align: center; }
        .status.new { color: #155724; }
        .status.updated { color: #856404; }
        .status.removed { color: #721c24; }
        .summary { background: #e7f3ff; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .count { font-size: 24px; font-weight: bold; color: #0078d4; }
        .legend { margin-top: 10px; }
        .legend-item { display: inline-block; margin-right: 20px; padding: 3px 10px; border-radius: 3px; font-size: 13px; }
        .legend-new { background: #d4edda; color: #155724; }
        .legend-updated { background: #fff3cd; color: #856404; }
        .legend-removed { background: #f8d7da; color: #721c24; }
        .legend-unchanged { background: #f8f9fa; color: #666; }
        .empty-version { color: #999; font-style: italic; }
        @media print {
            .container { box-shadow: none; }
            tr:hover { background: none; }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Software Inventory Report</h1>
        <div class="meta">
            <span><strong>Computer:</strong> $computerName</span>
            <span><strong>Generated:</strong> $timestamp</span>
        </div>
        <div class="summary">
            <span class="count">$($script:SoftwareList.Count)</span> applications
"@

    if ($hasComparison) {
        $html += @"

            <div class="legend">
                <span class="legend-item legend-new">$newCount New</span>
                <span class="legend-item legend-updated">$updatedCount Updated</span>
                <span class="legend-item legend-removed">$removedCount Removed</span>
                <span class="legend-item legend-unchanged">$unchangedCount Unchanged</span>
            </div>
"@
    }

    $html += @"
        </div>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Publisher</th>
                    <th>Source</th>
"@

    if ($hasComparison) {
        $html += "                    <th class=`"date-col`">$previousDate</th>`n"
        $html += "                    <th class=`"date-col`">$currentDate</th>`n"
        $html += "                    <th>Status</th>`n"
    } else {
        $html += "                    <th>Version</th>`n"
    }

    $html += @"
                </tr>
            </thead>
            <tbody>
"@

    foreach ($item in $script:SoftwareList) {
        $displayName = if ($item.CustomName) { $item.CustomName } else { $item.Name }
        $rowClass = ""
        if ($item.ChangeStatus -eq "NEW") { $rowClass = "new" }
        elseif ($item.ChangeStatus -eq "UPDATED") { $rowClass = "updated" }
        elseif ($item.ChangeStatus -eq "REMOVED") { $rowClass = "removed" }

        $html += "                <tr class=`"$rowClass`">`n"
        $html += "                    <td class=`"name`">$([System.Web.HttpUtility]::HtmlEncode($displayName))</td>`n"
        $html += "                    <td>$([System.Web.HttpUtility]::HtmlEncode($item.Publisher))</td>`n"
        $html += "                    <td>$([System.Web.HttpUtility]::HtmlEncode($item.Source))</td>`n"

        if ($hasComparison) {
            $prevVersion = $item.PreviousVersion
            if ($prevVersion) {
                $html += "                    <td class=`"version`">$([System.Web.HttpUtility]::HtmlEncode($prevVersion))</td>`n"
            } else {
                $html += "                    <td class=`"version empty-version`">-</td>`n"
            }

            $currVersion = $item.Version
            if ($currVersion) {
                $html += "                    <td class=`"version`">$([System.Web.HttpUtility]::HtmlEncode($currVersion))</td>`n"
            } else {
                $html += "                    <td class=`"version empty-version`">-</td>`n"
            }

            $statusClass = switch ($item.ChangeStatus) {
                "NEW" { "new" }
                "UPDATED" { "updated" }
                "REMOVED" { "removed" }
                default { "" }
            }
            $statusText = if ($item.ChangeStatus) { $item.ChangeStatus } else { "" }
            $html += "                    <td class=`"status $statusClass`">$statusText</td>`n"
        } else {
            $html += "                    <td class=`"version`">$([System.Web.HttpUtility]::HtmlEncode($item.Version))</td>`n"
        }

        $html += "                </tr>`n"
    }

    $html += @"
            </tbody>
        </table>
        <div class="meta" style="margin-top: 20px;">
            Generated by SoftwareLister v2.0
        </div>
    </div>
</body>
</html>
"@

    $html | Set-Content $FilePath -Encoding UTF8
}

#endregion

#region Dynamic Column Generation
function Build-DataGridColumns {
    param([System.Windows.Controls.DataGrid]$DataGrid)

    $DataGrid.Columns.Clear()

    $columnWidths = @{
        "Name" = 250
        "CustomName" = 150
        "Version" = 100
        "Publisher" = 150
        "InstallDate" = 100
        "Source" = 80
        "InstallLocation" = 250
        "Architecture" = 90
        "Size" = 80
        "UniqueId" = 200
        "IsSystemComponent" = 100
        "IsFramework" = 80
        "UninstallString" = 300
        "HelpLink" = 200
        "Comments" = 200
        "ChangeStatus" = 130
    }

    foreach ($prop in $script:Config.properties.PSObject.Properties) {
        if ($prop.Value.enabled -eq $true) {
            $propName = $prop.Name
            $displayName = $prop.Value.displayName
            $width = if ($columnWidths[$propName]) { $columnWidths[$propName] } else { 100 }

            $column = New-Object System.Windows.Controls.DataGridTextColumn
            $column.Header = $displayName
            $column.Binding = New-Object System.Windows.Data.Binding($propName)
            $column.Width = $width

            if ($propName -eq "CustomName") {
                $column.IsReadOnly = $false
                $style = New-Object System.Windows.Style([System.Windows.Controls.TextBlock])
                $style.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBlock]::ForegroundProperty, [System.Windows.Media.Brushes]::DodgerBlue)))
                $style.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBlock]::FontStyleProperty, [System.Windows.FontStyles]::Italic)))
                $column.ElementStyle = $style
            } else {
                $column.IsReadOnly = $true
            }

            $DataGrid.Columns.Add($column)
        }
    }

    $statusColumn = New-Object System.Windows.Controls.DataGridTextColumn
    $statusColumn.Header = "Status"
    $statusColumn.Binding = New-Object System.Windows.Data.Binding("ChangeStatus")
    $statusColumn.Width = 130
    $statusColumn.IsReadOnly = $true

    $statusStyle = New-Object System.Windows.Style([System.Windows.Controls.TextBlock])
    $statusStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBlock]::FontWeightProperty, [System.Windows.FontWeights]::Bold)))
    $statusColumn.ElementStyle = $statusStyle

    $DataGrid.Columns.Add($statusColumn)
}
#endregion

#region Driver Manager Functions

function Load-DriverData {
    $script:DrvC['drv_btnRefresh'].IsEnabled = $false
    $script:DrvC['drv_txtStatus'].Text       = "Enumerating PnP devices..."

    Write-DebugLine 'Drivers' "Enumerating PnP devices..."
    $devices  = @(Get-PnpDevice -ErrorAction SilentlyContinue)
    $devTotal = $devices.Count
    $script:DrvC['drv_txtStatus'].Text = "Found $devTotal devices - reading properties..."
    Write-DebugLine 'Drivers' "Found $devTotal PnP devices"

    $script:_drvIdx   = 0
    $script:_drvArr   = $devices
    $script:_drvTotal = $devTotal
    $script:_drvBuf   = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:_drvSw    = [System.Diagnostics.Stopwatch]::StartNew()

    Write-DebugLine 'Drivers' "Kicking off Dispatcher.Background processing..."
    $script:Window.Dispatcher.InvokeAsync(
        [Action]{ Process-DriverBatch },
        [System.Windows.Threading.DispatcherPriority]::Background
    ) | Out-Null
}

function Process-DriverBatch {
    $batchSize = 5
    $end = [Math]::Min($script:_drvIdx + $batchSize, $script:_drvTotal)

    while ($script:_drvIdx -lt $end) {
        $i   = $script:_drvIdx
        $dev = $script:_drvArr[$i]

        try {
            $propsArray = $dev | Get-PnpDeviceProperty -ErrorAction SilentlyContinue
            $props = @{}
            foreach ($p in $propsArray) { $props[$p.KeyName] = $p }

            $driverDateRaw = if ($props.ContainsKey('DEVPKEY_Device_DriverDate')) { $props['DEVPKEY_Device_DriverDate'].Data } else { $null }
            $driverDate    = if ($driverDateRaw -is [datetime]) { $driverDateRaw.ToString('yyyy-MM-dd') } else { $driverDateRaw }

            $script:_drvBuf.Add([PSCustomObject]@{
                FriendlyName   = if ($dev.FriendlyName) { $dev.FriendlyName } else { "(Unknown Device)" }
                Class          = $dev.Class
                Status         = $dev.Status
                Present        = $dev.Present
                DriverVersion  = if ($props.ContainsKey('DEVPKEY_Device_DriverVersion'))  { $props['DEVPKEY_Device_DriverVersion'].Data  } else { $null }
                DriverProvider = if ($props.ContainsKey('DEVPKEY_Device_DriverProvider')) { $props['DEVPKEY_Device_DriverProvider'].Data } else { $null }
                DriverDate     = $driverDate
                DriverInfPath  = if ($props.ContainsKey('DEVPKEY_Device_DriverInfPath'))  { $props['DEVPKEY_Device_DriverInfPath'].Data  } else { $null }
                InstanceId     = $dev.InstanceId
            })

            Write-DebugLine 'Drivers' ("[{0,4}/{1}] {2,-40} [{3}]" -f ($i + 1), $script:_drvTotal, $dev.FriendlyName, $dev.Class)
        }
        catch {
            Write-DebugLine 'Drivers' ("[{0,4}/{1}] ERROR: {2} - {3}" -f ($i + 1), $script:_drvTotal, $dev.FriendlyName, $_)
        }

        $script:_drvIdx++
    }

    $script:DrvC['drv_txtStatus'].Text = "Reading properties... $($script:_drvIdx) of $($script:_drvTotal) devices"

    if ($script:_drvIdx -lt $script:_drvTotal) {
        $script:Window.Dispatcher.InvokeAsync(
            [Action]{ Process-DriverBatch },
            [System.Windows.Threading.DispatcherPriority]::Background
        ) | Out-Null
    } else {
        $script:_drvSw.Stop()
        Write-DebugLine 'Drivers' "Complete - $($script:_drvTotal) devices in $($script:_drvSw.ElapsedMilliseconds)ms"
        $script:AllDevices = $script:_drvBuf

        $classes = @('All Classes') + ($script:AllDevices.Class | Sort-Object -Unique | Where-Object { $_ })
        $script:DrvC['drv_cboClass'].Items.Clear()
        foreach ($cls in $classes) {
            $item = [System.Windows.Controls.ComboBoxItem]::new()
            $item.Content = $cls
            $script:DrvC['drv_cboClass'].Items.Add($item) | Out-Null
        }
        $script:DrvC['drv_cboClass'].SelectedIndex = 0

        Apply-DriverFilters
        $script:DrvC['drv_btnRefresh'].IsEnabled = $true
    }
}

function Apply-DriverFilters {
    $search     = $script:DrvC['drv_txtSearch'].Text.Trim().ToLower()
    $classItem  = $script:DrvC['drv_cboClass'].SelectedItem  -as [System.Windows.Controls.ComboBoxItem]
    $statusItem = $script:DrvC['drv_cboStatus'].SelectedItem -as [System.Windows.Controls.ComboBoxItem]
    $presentItem= $script:DrvC['drv_cboPresent'].SelectedItem -as [System.Windows.Controls.ComboBoxItem]
    $classVal   = if ($classItem)  { $classItem.Content  } else { 'All Classes'  }
    $statusVal  = if ($statusItem) { $statusItem.Content } else { 'All Statuses' }
    $presentVal = if ($presentItem){ $presentItem.Content} else { 'All Devices'  }

    $filtered = $script:AllDevices | Where-Object {
        $row = $_

        $matchSearch = ($search -eq '') -or (
            ($row.FriendlyName   -and $row.FriendlyName.ToLower().Contains($search))  -or
            ($row.Class          -and $row.Class.ToLower().Contains($search))          -or
            ($row.DriverVersion  -and $row.DriverVersion.ToLower().Contains($search))  -or
            ($row.DriverProvider -and $row.DriverProvider.ToLower().Contains($search)) -or
            ($row.DriverInfPath  -and $row.DriverInfPath.ToLower().Contains($search))  -or
            ($row.InstanceId     -and $row.InstanceId.ToLower().Contains($search))
        )

        $matchClass   = ($classVal  -eq 'All Classes')  -or ($row.Class   -eq $classVal)
        $matchStatus  = ($statusVal -eq 'All Statuses') -or ($row.Status  -eq $statusVal)
        $matchPresent = switch ($presentVal) {
            'Present Only' { $row.Present -eq $true  }
            'Absent Only'  { $row.Present -eq $false }
            default        { $true }
        }

        $matchSearch -and $matchClass -and $matchStatus -and $matchPresent
    }

    $script:DrvC['drv_dgDevices'].ItemsSource = $filtered
    $script:DrvC['drv_txtStatus'].Text    = "$($filtered.Count) devices shown  |  $($script:AllDevices.Count) total"
    $script:DrvC['drv_txtSubtitle'].Text  = "$($script:AllDevices.Count) devices loaded"
}

function Backup-Drivers {
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description  = "Select folder to save driver backup"
    $dialog.SelectedPath = [Environment]::GetFolderPath('Desktop')

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $dest = Join-Path $dialog.SelectedPath "DriverBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    $script:DrvC['drv_btnBackup'].IsEnabled = $false
    $script:DrvC['drv_btnBackup'].Content   = "Backing up..."
    $script:DrvC['drv_txtStatus'].Text      = "Exporting drivers to: $dest"

    # Export-WindowsDriver can take several minutes; run it in a new PS runspace
    # so the UI thread stays free. A DispatcherTimer polls for completion.
    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.AddScript({
        param($destination)
        try {
            Export-WindowsDriver -Online -Destination $destination -ErrorAction Stop | Out-Null
            return @{ Success = $true; Path = $destination; Error = $null }
        } catch {
            return @{ Success = $false; Path = $destination; Error = $_.Exception.Message }
        }
    }).AddParameter('destination', $dest) | Out-Null

    $script:_drvBackupPs     = $ps
    $script:_drvBackupHandle = $ps.BeginInvoke()
    $script:_drvBackupDest   = $dest

    # DispatcherTimer runs on the UI thread every 500 ms; safe for UI access.
    $script:_drvBackupTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:_drvBackupTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $script:_drvBackupTimer.Add_Tick({
        if (-not $script:_drvBackupHandle.IsCompleted) { return }

        $script:_drvBackupTimer.Stop()

        try {
            $res = $script:_drvBackupPs.EndInvoke($script:_drvBackupHandle)
            $r   = if ($res -and $res.Count -gt 0) { $res[0] } else { @{ Success = $false; Path = $script:_drvBackupDest; Error = "No result returned" } }
        } catch {
            $r = @{ Success = $false; Path = $script:_drvBackupDest; Error = $_.Exception.Message }
        } finally {
            $script:_drvBackupPs.Dispose()
            $script:_drvBackupPs = $null
        }

        $script:DrvC['drv_btnBackup'].IsEnabled = $true
        $script:DrvC['drv_btnBackup'].Content   = "Backup Drivers"

        if ($r.Success) {
            $count = (Get-ChildItem -Path $r.Path -Directory -ErrorAction SilentlyContinue).Count
            $script:DrvC['drv_txtStatus'].Text = "Backup complete: $count drivers exported to $($r.Path)"
            [System.Windows.MessageBox]::Show(
                "Driver backup complete!`n`n$count driver packages exported to:`n$($r.Path)",
                "Backup Successful",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
        } else {
            $script:DrvC['drv_txtStatus'].Text = "Backup failed: $($r.Error)"
            [System.Windows.MessageBox]::Show(
                "Driver backup failed:`n`n$($r.Error)",
                "Backup Failed",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
    })
    $script:_drvBackupTimer.Start()
}

function Install-InfList {
    param([string[]]$InfPaths)

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($inf in $InfPaths) {
        $entry = [PSCustomObject]@{
            InfFile = $inf
            Success = $false
            Message = ''
        }
        try {
            $pnp  = & pnputil.exe /add-driver $inf /install 2>&1
            $exit = $LASTEXITCODE

            if ($exit -eq 0) {
                $entry.Success = $true
                $entry.Message = 'Installed successfully'
            }
            elseif ($exit -eq 3010) {
                $entry.Success = $true
                $entry.Message = 'Installed - reboot required'
            }
            else {
                $entry.Message = "pnputil exit code $exit - $($pnp -join ' ')"
            }
        }
        catch {
            $entry.Message = $_.Exception.Message
        }
        $results.Add($entry)
    }
    return $results
}

function Show-ImportResults {
    param([System.Collections.Generic.List[PSCustomObject]]$Results)

    $ok     = @($Results | Where-Object Success)
    $failed = @($Results | Where-Object { -not $_.Success })
    $reboot = @($Results | Where-Object { $_.Message -like '*reboot*' })

    $summary  = "Processed $($Results.Count) driver(s):`n"
    $summary += "  Succeeded : $($ok.Count)`n"
    if ($reboot.Count -gt 0) { $summary += "  Need reboot: $($reboot.Count)`n" }
    $summary += "  Failed    : $($failed.Count)"

    $detail = ($Results | ForEach-Object {
        $icon = if ($_.Success) { 'OK' } else { 'FAIL' }
        "$icon  $(Split-Path $_.InfFile -Leaf)`n     $($_.Message)"
    }) -join "`n`n"

    $icon = if ($failed.Count -eq 0) {
        [System.Windows.MessageBoxImage]::Information
    } else {
        [System.Windows.MessageBoxImage]::Warning
    }

    [System.Windows.MessageBox]::Show(
        "$summary`n`n── Details ──`n$detail",
        "Import Results",
        [System.Windows.MessageBoxButton]::OK,
        $icon
    )
}

function Import-SingleDriver {
    $dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $dialog.Title            = "Select a driver INF file to install"
    $dialog.Filter           = "INF Files (*.inf)|*.inf|All Files (*.*)|*.*"
    $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $infPath = $dialog.FileName
    $confirm = [System.Windows.MessageBox]::Show(
        "Install driver from:`n$infPath`n`nThis will call pnputil /add-driver /install. Continue?",
        "Confirm Driver Install",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

    $script:DrvC['drv_btnImportFile'].IsEnabled = $false
    $script:DrvC['drv_btnImportFile'].Content   = "Installing..."
    $script:DrvC['drv_txtStatus'].Text          = "Installing: $(Split-Path $infPath -Leaf) ..."

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.AddScript({
        param($path)
        $entry = [PSCustomObject]@{ InfFile = $path; Success = $false; Message = '' }
        try {
            $out  = & pnputil.exe /add-driver $path /install 2>&1
            $exit = $LASTEXITCODE
            if ($exit -eq 0)    { $entry.Success = $true; $entry.Message = 'Installed successfully' }
            elseif ($exit -eq 3010) { $entry.Success = $true; $entry.Message = 'Installed - reboot required' }
            else                { $entry.Message = "pnputil exit $exit - $($out -join ' ')" }
        } catch { $entry.Message = $_.Exception.Message }
        return [System.Collections.Generic.List[PSCustomObject]]@($entry)
    }).AddParameter('path', $infPath) | Out-Null

    $script:_drvImportPs     = $ps
    $script:_drvImportHandle = $ps.BeginInvoke()
    $script:_drvImportPath   = $infPath

    $script:_drvImportTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:_drvImportTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $script:_drvImportTimer.Add_Tick({
        if (-not $script:_drvImportHandle.IsCompleted) { return }
        $script:_drvImportTimer.Stop()
        try {
            $results = $script:_drvImportPs.EndInvoke($script:_drvImportHandle)
        } catch {
            $r = [PSCustomObject]@{ InfFile = $script:_drvImportPath; Success = $false; Message = $_.Exception.Message }
            $results = [System.Collections.Generic.List[PSCustomObject]]@($r)
        } finally {
            $script:_drvImportPs.Dispose(); $script:_drvImportPs = $null
        }
        $script:DrvC['drv_btnImportFile'].IsEnabled = $true
        $script:DrvC['drv_btnImportFile'].Content   = "Import Driver (.inf)"
        $ok = @($results | Where-Object Success).Count
        $script:DrvC['drv_txtStatus'].Text = if ($ok -gt 0) {
            "Driver installed: $(Split-Path $script:_drvImportPath -Leaf)"
        } else {
            "Driver install failed: $(Split-Path $script:_drvImportPath -Leaf)"
        }
        Show-ImportResults -Results $results
        if ($ok -gt 0) { Load-DriverData }
    })
    $script:_drvImportTimer.Start()
}

function Import-DriverFolder {
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description  = "Select folder containing driver INF files (searched recursively)"
    $dialog.SelectedPath = [Environment]::GetFolderPath('Desktop')

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $folder   = $dialog.SelectedPath
    $infFiles = @(Get-ChildItem -Path $folder -Filter '*.inf' -Recurse -ErrorAction SilentlyContinue |
                  Select-Object -ExpandProperty FullName)

    if ($infFiles.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No .inf files found in:`n$folder",
            "No Drivers Found",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        )
        return
    }

    $confirm = [System.Windows.MessageBox]::Show(
        "Found $($infFiles.Count) INF file(s) in:`n$folder`n`nInstall all of them? This may take a while.",
        "Confirm Bulk Driver Install",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

    $script:DrvC['drv_btnImportFolder'].IsEnabled = $false
    $script:DrvC['drv_btnImportFolder'].Content   = "Installing..."
    $script:DrvC['drv_txtStatus'].Text            = "Installing $($infFiles.Count) drivers from folder..."

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.AddScript({
        param($paths)
        $results = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($inf in $paths) {
            $entry = [PSCustomObject]@{ InfFile = $inf; Success = $false; Message = '' }
            try {
                $out  = & pnputil.exe /add-driver $inf /install 2>&1
                $exit = $LASTEXITCODE
                if ($exit -eq 0)        { $entry.Success = $true; $entry.Message = 'Installed successfully' }
                elseif ($exit -eq 3010) { $entry.Success = $true; $entry.Message = 'Installed - reboot required' }
                else                    { $entry.Message = "pnputil exit $exit - $($out -join ' ')" }
            } catch { $entry.Message = $_.Exception.Message }
            $results.Add($entry)
        }
        return $results
    }).AddParameter('paths', $infFiles) | Out-Null

    $script:_drvFolderPs     = $ps
    $script:_drvFolderHandle = $ps.BeginInvoke()

    $script:_drvFolderTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:_drvFolderTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $script:_drvFolderTimer.Add_Tick({
        if (-not $script:_drvFolderHandle.IsCompleted) { return }
        $script:_drvFolderTimer.Stop()
        try {
            $results = $script:_drvFolderPs.EndInvoke($script:_drvFolderHandle)
        } catch {
            $results = [System.Collections.Generic.List[PSCustomObject]]::new()
        } finally {
            $script:_drvFolderPs.Dispose(); $script:_drvFolderPs = $null
        }
        $script:DrvC['drv_btnImportFolder'].IsEnabled = $true
        $script:DrvC['drv_btnImportFolder'].Content   = "Import Driver Folder"
        $ok     = @($results | Where-Object Success).Count
        $failed = @($results | Where-Object { -not $_.Success }).Count
        $script:DrvC['drv_txtStatus'].Text = "$ok installed  |  $failed failed  (from $($results.Count) total)"
        Show-ImportResults -Results $results
        if ($ok -gt 0) { Load-DriverData }
    })
    $script:_drvFolderTimer.Start()
}

#endregion

#region Service Manager Functions

function Resolve-ServiceExePath {
    param([string]$ImagePath)
    if ([string]::IsNullOrWhiteSpace($ImagePath)) { return $null }

    $ImagePath = $ImagePath.Trim()

    if ($ImagePath -match '^"([^"]+)"') {
        return $Matches[1]
    }

    if ($ImagePath -match '^([A-Za-z]:\\[^\s]+\.exe)') {
        return $Matches[1]
    }

    return ($ImagePath -split ' ')[0]
}

function Get-ServiceControlPermissions {
    param([string]$ServiceName)

    try {
        $sd = & sc.exe sdshow $ServiceName 2>$null | Where-Object { $_ -match 'D:' }
        if (-not $sd) { return @{ Summary = 'Unable to read SD'; Risk = $false } }

        $sidMap = @{
            'WD' = 'Everyone'
            'AU' = 'Authenticated Users'
            'IU' = 'Interactive Users'
            'BU' = 'Built-in Users'
            'BA' = 'Administrators'
            'SY' = 'SYSTEM'
            'LS' = 'Local Service'
            'NS' = 'Network Service'
            'PU' = 'Power Users'
            'NO' = 'Network Configuration Operators'
            'SO' = 'Server Operators'
        }

        # Only rights that are genuine privilege-escalation vectors:
        #   DC = SERVICE_CHANGE_CONFIG (can replace binary path)
        #   WD = WRITE_DAC            (can loosen the service DACL itself)
        #   WO = WRITE_OWNER          (can take ownership then change DACL)
        # Read/query/start rights (CC, LC, RC, RP, SW, LO, CR…) are granted to
        # Authenticated Users on almost every service by default — not risky.
        $controlRights = @('DC', 'WD', 'WO')
        $nonAdminSids  = @('WD','AU','IU','BU','PU','NO','SO')

        $riskyEntries = [System.Collections.Generic.List[string]]::new()

        $acePattern = '\(([^)]+)\)'
        $aces = [regex]::Matches($sd, $acePattern)

        foreach ($ace in $aces) {
            $parts = $ace.Groups[1].Value -split ';'
            if ($parts.Count -lt 6) { continue }

            $aceType = $parts[0]
            $rights  = $parts[2]
            $sid     = $parts[5]

            if ($aceType -eq 'A') {
                $hasControl = $false
                foreach ($r in $controlRights) {
                    if ($rights -like "*$r*") { $hasControl = $true; break }
                }
                if ($hasControl -and $nonAdminSids -contains $sid) {
                    $friendlyId = if ($sidMap.ContainsKey($sid)) { $sidMap[$sid] } else { $sid }
                    $riskyEntries.Add($friendlyId)
                }
            }
        }

        $isRisky = $riskyEntries.Count -gt 0
        $summary = if ($isRisky) {
            "Non-admin control: $($riskyEntries -join ', ')"
        } else {
            "Restricted to privileged accounts"
        }

        return @{ Summary = $summary; Risk = $isRisky }
    }
    catch {
        return @{ Summary = "Error: $_"; Risk = $false }
    }
}

function Get-ExeAclAnalysis {
    param([string]$ExePath)

    $empty = [System.Collections.Generic.List[PSCustomObject]]::new()

    if ([string]::IsNullOrWhiteSpace($ExePath)) {
        return @{ Entries = $empty; IsRisky = $false; Summary = 'No exe path' }
    }

    $ExePath = [System.Environment]::ExpandEnvironmentVariables($ExePath)

    if (-not (Test-Path $ExePath -ErrorAction SilentlyContinue)) {
        return @{ Entries = $empty; IsRisky = $false; Summary = "Exe not found: $ExePath" }
    }

    try {
        $acl     = Get-Acl -Path $ExePath -ErrorAction Stop
        $entries = [System.Collections.Generic.List[PSCustomObject]]::new()
        $isRisky = $false

        # Only rights that actually allow replacing/hijacking the executable.
        # AppendData (4) is excluded: it shares a bit with CreateDirectories,
        # so it appears on virtually every file inherited from a standard folder ACL.
        # WriteExtendedAttributes and WriteAttributes alone are not exploitable.
        # WriteData catches all composite rights (Write, Modify, FullControl) because
        # they all include that bit.
        $writeRights = @(
            [System.Security.AccessControl.FileSystemRights]::WriteData,        # overwrite file content; also matches Write/Modify/FullControl
            [System.Security.AccessControl.FileSystemRights]::TakeOwnership,   # can take ownership then grant self WriteData
            [System.Security.AccessControl.FileSystemRights]::ChangePermissions # can loosen DACL then grant self WriteData
        )

        $privileged = @('Administrator', 'SYSTEM', 'TrustedInstaller', 'Administrators',
                        'NT SERVICE', 'NT AUTHORITY\SYSTEM', 'Creator Owner')

        foreach ($ace in $acl.Access) {
            $identity = $ace.IdentityReference.Value
            $rights   = $ace.FileSystemRights
            $aceType  = $ace.AccessControlType

            $hasWrite = $false
            foreach ($wr in $writeRights) {
                if (($rights -band $wr) -ne 0) { $hasWrite = $true; break }
            }

            $isPriv = $false
            foreach ($p in $privileged) {
                if ($identity -like "*$p*") { $isPriv = $true; break }
            }

            $riskyFlag = if ($hasWrite -and -not $isPriv -and $aceType -eq 'Allow') {
                $isRisky = $true
                '! Yes'
            } elseif ($hasWrite) {
                'Write (privileged)'
            } else {
                '-'
            }

            $entries.Add([PSCustomObject]@{
                Identity = $identity
                Rights   = $rights.ToString()
                AceType  = $aceType.ToString()
                IsRisky  = $riskyFlag
            })
        }

        $riskyAccounts = @($entries | Where-Object { $_.IsRisky -eq '! Yes' } |
                          Select-Object -ExpandProperty Identity)
        $summary = if ($isRisky) {
            "Write access: $($riskyAccounts -join ', ')"
        } else {
            "No non-admin write access detected"
        }

        return @{ Entries = $entries; IsRisky = $isRisky; Summary = $summary }
    }
    catch {
        return @{ Entries = $empty; IsRisky = $false; Summary = "ACL read error: $_" }
    }
}

function Load-ServiceData {
    # -- Disable controls ----------------------------------------------------------
    $script:SvcC['svc_btnRefresh'].IsEnabled = $false
    $script:SvcC['svc_btnStart'].IsEnabled   = $false
    $script:SvcC['svc_btnStop'].IsEnabled    = $false
    $script:SvcC['svc_btnRestart'].IsEnabled = $false
    $script:SvcC['svc_pnlDetail'].Visibility = 'Collapsed'

    $doStartup = ($script:SvcC['svc_chkStartupType'].IsChecked -eq $true)
    $doCtrl    = ($script:SvcC['svc_chkSvcControl'].IsChecked  -eq $true)
    $doAcl     = ($script:SvcC['svc_chkExeAcl'].IsChecked     -eq $true)

    Write-DebugLine 'Services' ("Starting - Options: Startup={0} SvcCtrl={1} ExeAcl={2}" -f $doStartup, $doCtrl, $doAcl)

    # -- Enumerate services (fast) -------------------------------------------------
    $svcsArray = @(Get-Service -ErrorAction SilentlyContinue)
    $svcTotal  = $svcsArray.Count
    $script:SvcC['svc_txtStatus'].Text = "Found $svcTotal services..."
    Write-DebugLine 'Services' "Found $svcTotal services"

    # -- Bulk WMI query for startup types (synchronous on UI thread) ---------------
    # Single WMI call - typically completes in 1-3 s.
    $wmiMap = @{}
    if ($doStartup) {
        $script:SvcC['svc_txtStatus'].Text = "Querying WMI startup types ($svcTotal services)..."
        Write-DebugLine 'Services' "WMI bulk query starting..."
        $wmiSw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $wmiAll = Get-WmiObject Win32_Service -ErrorAction SilentlyContinue
            if ($wmiAll) { foreach ($w in $wmiAll) { $wmiMap[$w.Name] = $w } }
        } catch { }
        $wmiSw.Stop()
        Write-DebugLine 'Services' "WMI complete - $($wmiMap.Count) entries in $($wmiSw.ElapsedMilliseconds)ms"
    }

    $script:SvcC['svc_txtStatus'].Text = "Analyzing services... 0 of $svcTotal"

    # -- Store shared state consumed by Process-ServiceBatch -----------------------
    $script:_svcIdx    = 0
    $script:_svcArr    = $svcsArray
    $script:_svcTotal  = $svcTotal
    $script:_svcWmiMap = $wmiMap
    $script:_svcDoSt   = $doStartup
    $script:_svcDoCtrl = $doCtrl
    $script:_svcDoAcl  = $doAcl
    $script:_svcBuf    = [System.Collections.Generic.List[PSCustomObject]]::new()
    $script:_svcSw     = [System.Diagnostics.Stopwatch]::StartNew()

    Write-DebugLine 'Services' "Kicking off Dispatcher.Background processing..."

    # -- Kick off incremental processing at Background priority -------------------
    # Background priority lets WPF handle input/render between each batch.
    $script:Window.Dispatcher.InvokeAsync(
        [Action]{ Process-ServiceBatch },
        [System.Windows.Threading.DispatcherPriority]::Background
    ) | Out-Null
}

function Process-ServiceBatch {
    # Process up to $batchSize services per UI tick, then re-queue at Background
    # priority so input/render events can run between batches.
    $batchSize = 5
    $end = [Math]::Min($script:_svcIdx + $batchSize, $script:_svcTotal)

    while ($script:_svcIdx -lt $end) {
        $i   = $script:_svcIdx
        $svc = $script:_svcArr[$i]

        $regPath   = "HKLM:\SYSTEM\CurrentControlSet\Services\$($svc.Name)"
        $regKey    = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        $imagePath = $regKey.ImagePath
        $exePath   = Resolve-ServiceExePath -ImagePath $imagePath
        $logOnAs   = $regKey.ObjectName

        Write-DebugLine 'Services' ("[{0,4}/{1}] Checking: {2} ({3})" -f ($i + 1), $script:_svcTotal, $svc.Name, $svc.Status)

        $startupType = if ($script:_svcDoSt) {
            try {
                $w = $script:_svcWmiMap[$svc.Name]
                switch ($w.StartMode) {
                    'Auto'     { if ($regKey.DelayedAutostart -eq 1) { 'Automatic (Delayed)' } else { 'Automatic' } }
                    'Manual'   { 'Manual' }
                    'Disabled' { 'Disabled' }
                    default    { if ($w) { $w.StartMode } else { $svc.StartType.ToString() } }
                }
            } catch { $svc.StartType.ToString() }
        } else {
            $svc.StartType.ToString()
        }

        $svcCtrl = if ($script:_svcDoCtrl) {
            Write-DebugLine 'Services' "      sc.exe sdshow $($svc.Name)"
            Get-ServiceControlPermissions -ServiceName $svc.Name
        } else {
            @{ Summary = 'Not queried'; Risk = $false }
        }
        $svcCtrlRisk = if (-not $script:_svcDoCtrl) { 'N/A' } elseif ($svcCtrl.Risk) { '! Weak' } else { 'OK' }

        $exeAcl = if ($script:_svcDoAcl) {
            Write-DebugLine 'Services' "      Get-Acl: $exePath"
            Get-ExeAclAnalysis -ExePath $exePath
        } else {
            @{ Entries = [System.Collections.Generic.List[PSCustomObject]]::new(); IsRisky = $false; Summary = 'Not queried' }
        }
        $exeRisk = if (-not $script:_svcDoAcl) {
            'N/A'
        } elseif ($exeAcl.IsRisky) {
            '! Risky'
        } elseif ([string]::IsNullOrWhiteSpace($exePath)) {
            'N/A'
        } else {
            'OK'
        }

        $script:_svcBuf.Add([PSCustomObject]@{
            Name                 = $svc.Name
            DisplayName          = $svc.DisplayName
            State                = $svc.Status.ToString()
            StartupType          = $startupType
            LogOnAs              = if ($logOnAs) { $logOnAs } else { 'N/A' }
            ExePath              = if ($exePath) { $exePath } else { $imagePath }
            ExePathRaw           = $exePath
            ServiceControlRisk   = $svcCtrlRisk
            ServiceControlDetail = $svcCtrl.Summary
            ExeWriteRisk         = $exeRisk
            ExeAclSummary        = $exeAcl.Summary
        })

        if ($script:DebugMode) {
            $flag   = if ($svcCtrlRisk -eq '! Weak' -or $exeRisk -eq '! Risky') { '! ' } else { '  ' }
            $exeStr = if ($exePath) { $exePath } else { $imagePath }
            $ts = (Get-Date).ToString('HH:mm:ss.fff')
            [Console]::WriteLine(("[$ts][DBG][Services] $flag[{0,4}/{1}] {2,-28} | {3,-8} | Startup={4,-22} | Ctrl={5,-6} | ACL={6,-8} | {7}" -f ($i + 1), $script:_svcTotal, $svc.Name, $svc.Status, $startupType, $svcCtrlRisk, $exeRisk, $exeStr))
        }

        $script:_svcIdx++
    }

    # Progress update
    $script:SvcC['svc_txtStatus'].Text = "Analyzing... $($script:_svcIdx) of $($script:_svcTotal) services"

    if ($script:_svcIdx -lt $script:_svcTotal) {
        # More services remain -- re-queue at Background priority
        $script:Window.Dispatcher.InvokeAsync(
            [Action]{ Process-ServiceBatch },
            [System.Windows.Threading.DispatcherPriority]::Background
        ) | Out-Null
    } else {
        # All done -- commit results and re-enable UI
        $script:_svcSw.Stop()
        Write-DebugLine 'Services' "Complete - $($script:_svcTotal) services in $($script:_svcSw.ElapsedMilliseconds)ms"
        $script:AllServices = $script:_svcBuf
        $script:AclCache    = @{}
        Apply-ServiceFilters
        $script:SvcC['svc_btnRefresh'].IsEnabled = $true
    }
}

function Apply-ServiceFilters {
    $stateItem   = $script:SvcC['svc_cboState'].SelectedItem   -as [System.Windows.Controls.ComboBoxItem]
    $startupItem = $script:SvcC['svc_cboStartup'].SelectedItem -as [System.Windows.Controls.ComboBoxItem]
    $riskItem    = $script:SvcC['svc_cboRisk'].SelectedItem    -as [System.Windows.Controls.ComboBoxItem]

    $search  = $script:SvcC['svc_txtSearch'].Text.Trim().ToLower()
    $state   = if ($stateItem)   { $stateItem.Content   } else { 'All States' }
    $startup = if ($startupItem) { $startupItem.Content } else { 'All Startup Types' }
    $risk    = if ($riskItem)    { $riskItem.Content    } else { 'All Services' }

    $filtered = $script:AllServices | Where-Object {
        $row = $_

        $matchSearch = ($search -eq '') -or (
            ($row.Name        -and $row.Name.ToLower().Contains($search))        -or
            ($row.DisplayName -and $row.DisplayName.ToLower().Contains($search)) -or
            ($row.ExePath     -and $row.ExePath.ToLower().Contains($search))     -or
            ($row.LogOnAs     -and $row.LogOnAs.ToLower().Contains($search))
        )

        $matchState   = ($state -eq 'All States')         -or ($row.State -eq $state)
        $matchStartup = ($startup -eq 'All Startup Types') -or ($row.StartupType -eq $startup)
        $matchRisk    = switch ($risk) {
            '!  Risky Exe ACL'    { $row.ExeWriteRisk -eq '! Risky' }
            '!  Weak Svc Control' { $row.ServiceControlRisk -eq '! Weak' }
            default               { $true }
        }

        $matchSearch -and $matchState -and $matchStartup -and $matchRisk
    }

    $script:SvcC['svc_dgServices'].ItemsSource = $filtered

    $total  = $script:AllServices.Count
    $shown  = ($filtered | Measure-Object).Count
    $risky  = @($script:AllServices | Where-Object { $_.ExeWriteRisk -eq '! Risky' }).Count
    $weak   = @($script:AllServices | Where-Object { $_.ServiceControlRisk -eq '! Weak' }).Count
    $script:SvcC['svc_txtStatus'].Text   = "$shown of $total services shown   |   ! $risky risky exe ACLs   ! $weak weak svc control"
    $script:SvcC['svc_txtSubtitle'].Text = "$total services loaded"
}

function Invoke-ServiceAction {
    param([string]$Action)

    $row = $script:SvcC['svc_dgServices'].SelectedItem
    if (-not $row) { return }

    $svcName = $row.Name
    $confirm = [System.Windows.MessageBox]::Show(
        "$Action service '$svcName'?",
        "Confirm Action",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

    $script:SvcC['svc_btnStart'].IsEnabled   = $false
    $script:SvcC['svc_btnStop'].IsEnabled    = $false
    $script:SvcC['svc_btnRestart'].IsEnabled = $false
    $script:SvcC['svc_txtStatus'].Text       = "$Action '$svcName'..."

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.AddScript({
        param($action, $name)
        try {
            switch ($action) {
                'Start'   { Start-Service   -Name $name -ErrorAction Stop }
                'Stop'    { Stop-Service    -Name $name -Force -ErrorAction Stop }
                'Restart' { Restart-Service -Name $name -Force -ErrorAction Stop }
            }
            return @{ Success = $true;  Message = "$action completed for '$name'" }
        } catch {
            return @{ Success = $false; Message = $_.Exception.Message }
        }
    }).AddParameter('action', $Action).AddParameter('name', $svcName) | Out-Null

    $script:_svcActionPs     = $ps
    $script:_svcActionHandle = $ps.BeginInvoke()
    $script:_svcActionName   = $svcName

    $script:_svcActionTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:_svcActionTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $script:_svcActionTimer.Add_Tick({
        if (-not $script:_svcActionHandle.IsCompleted) { return }
        $script:_svcActionTimer.Stop()
        try {
            $result = $script:_svcActionPs.EndInvoke($script:_svcActionHandle)
            if ($result -and $result.Count -gt 0) { $r = $result[0] }
            else { $r = @{ Success = $false; Message = "No result returned" } }
        } catch {
            $r = @{ Success = $false; Message = $_.Exception.Message }
        } finally {
            $script:_svcActionPs.Dispose(); $script:_svcActionPs = $null
        }
        $script:SvcC['svc_txtStatus'].Text = if ($r.Success) {
            "[OK]  $($r.Message)"
        } else {
            "[FAIL]  $($r.Message)"
        }
        $script:SvcC['svc_btnStart'].IsEnabled   = $true
        $script:SvcC['svc_btnStop'].IsEnabled    = $true
        $script:SvcC['svc_btnRestart'].IsEnabled = $true
        Load-ServiceData
    })
    $script:_svcActionTimer.Start()
}

function Update-ServiceDetailPanel {
    param($row)

    if ($null -eq $row) {
        $script:SvcC['svc_pnlDetail'].Visibility = 'Collapsed'
        $script:SvcC['svc_btnStart'].IsEnabled   = $false
        $script:SvcC['svc_btnStop'].IsEnabled    = $false
        $script:SvcC['svc_btnRestart'].IsEnabled = $false
        return
    }

    $script:SvcC['svc_pnlDetail'].Visibility    = 'Visible'
    $script:SvcC['svc_btnStart'].IsEnabled      = ($row.State -ne 'Running')
    $script:SvcC['svc_btnStop'].IsEnabled       = ($row.State -eq 'Running')
    $script:SvcC['svc_btnRestart'].IsEnabled    = $true

    $script:SvcC['svc_detailName'].Text        = $row.Name
    $script:SvcC['svc_detailDisplayName'].Text = "Display: $($row.DisplayName)"
    $script:SvcC['svc_detailState'].Text       = "State: $($row.State)"
    $script:SvcC['svc_detailStartup'].Text     = "Startup: $($row.StartupType)"
    $script:SvcC['svc_detailLogOn'].Text       = "Log on as: $($row.LogOnAs)"
    $script:SvcC['svc_detailExePath'].Text     = if ($row.ExePath) { $row.ExePath } else { '(not resolved)' }
    $script:SvcC['svc_detailSvcControl'].Text  = "Service Control: $($row.ServiceControlRisk)`n$($row.ServiceControlDetail)"
    $script:SvcC['svc_detailExeAcl'].Text      = "Exe Write ACL: $($row.ExeWriteRisk)`n$($row.ExeAclSummary)"

    if (-not $script:AclCache.ContainsKey($row.Name)) {
        $aclResult = Get-ExeAclAnalysis -ExePath $row.ExePathRaw
        $script:AclCache[$row.Name] = $aclResult.Entries
    }
    $script:SvcC['svc_dgAcl'].ItemsSource = $script:AclCache[$row.Name]
}

#endregion

#region XAML Definition
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SoftwareLister - System Manager Suite" Height="800" Width="1300"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5">
    <Window.Resources>

        <!-- ── Software tab button styles ── -->
        <Style TargetType="Button" x:Key="ToolbarButton">
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Margin" Value="3"/>
            <Setter Property="Background" Value="#0078d4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#106ebe"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#cccccc"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="Button" x:Key="SecondaryButton">
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Margin" Value="3"/>
            <Setter Property="Background" Value="#ffffff"/>
            <Setter Property="Foreground" Value="#333333"/>
            <Setter Property="BorderBrush" Value="#cccccc"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#e0e0e0"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- ── Driver / Service tab shared styles (Catppuccin Mocha palette) ── -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#313244"/>
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="5" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#45475A"/></Trigger>
                            <Trigger Property="IsPressed"   Value="True"><Setter Property="Background" Value="#585B70"/></Trigger>
                            <Trigger Property="IsEnabled"   Value="False"><Setter Property="Opacity"   Value="0.45"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="AccentButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#89B4FA"/>
            <Setter Property="Foreground" Value="#1E1E2E"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#74C7EC"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#A6E3A1"/>
            <Setter Property="Foreground" Value="#1E1E2E"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#94E2D5"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#F38BA8"/>
            <Setter Property="Foreground" Value="#1E1E2E"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#EBA0AC"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="SearchBox" TargetType="TextBox">
            <Setter Property="Background" Value="#313244"/>
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="CaretBrush" Value="#CDD6F4"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}" CornerRadius="5" Padding="{TemplateBinding Padding}">
                            <ScrollViewer x:Name="PART_ContentHost" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="ModernCombo" TargetType="ComboBox">
            <Setter Property="Background" Value="#313244"/>
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style x:Key="ModernGrid" TargetType="DataGrid">
            <Setter Property="Background" Value="#181825"/>
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="RowBackground" Value="#181825"/>
            <Setter Property="AlternatingRowBackground" Value="#1E1E2E"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="SelectionMode" Value="Single"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="CanUserReorderColumns" Value="True"/>
            <Setter Property="CanUserResizeColumns" Value="True"/>
            <Setter Property="CanUserSortColumns" Value="True"/>
            <Setter Property="HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>

    </Window.Resources>

    <TabControl Name="tcMain">

        <!-- ════════════════════════════════════════════════════ SOFTWARE TAB ══ -->
        <TabItem Header="Software" Name="tabSoftware">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Toolbar -->
                <Border Name="swToolbarBorder" Grid.Row="0" Background="White" BorderBrush="#e0e0e0" BorderThickness="0,0,0,1" Padding="10">
                    <StackPanel Orientation="Horizontal">
                        <Button Name="btnScan" Content="Scan Software" Style="{StaticResource ToolbarButton}"/>
                        <Separator Style="{StaticResource {x:Static ToolBar.SeparatorStyleKey}}" Margin="10,0"/>
                        <Button Name="btnCompare" Content="Compare with Previous..." Style="{StaticResource SecondaryButton}" IsEnabled="False"/>
                        <Separator Style="{StaticResource {x:Static ToolBar.SeparatorStyleKey}}" Margin="10,0"/>
                        <Button Name="btnExportCsv" Content="Export CSV" Style="{StaticResource SecondaryButton}" IsEnabled="False"/>
                        <Button Name="btnExportHtml" Content="Export HTML" Style="{StaticResource SecondaryButton}" IsEnabled="False"/>
                        <Button Name="btnExportJson" Content="Export JSON" Style="{StaticResource SecondaryButton}" IsEnabled="False"/>
                        <Separator Style="{StaticResource {x:Static ToolBar.SeparatorStyleKey}}" Margin="10,0"/>
                        <Button Name="btnSettings" Content="Settings" Style="{StaticResource SecondaryButton}"/>
                    </StackPanel>
                </Border>

                <!-- Search and Filter Bar -->
                <Border Name="swSearchBorder" Grid.Row="1" Background="#fafafa" Padding="10">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="300"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <TextBlock Grid.Column="0" Text="Search:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                        <TextBox Name="txtSearch" Grid.Column="1" Padding="5" VerticalContentAlignment="Center"/>
                        <Button Name="btnClearSearch" Grid.Column="2" Content="Clear" Style="{StaticResource SecondaryButton}" Padding="10,5"/>

                        <StackPanel Grid.Column="4" Orientation="Horizontal">
                            <TextBlock Text="Items:" VerticalAlignment="Center" Margin="0,0,5,0"/>
                            <TextBlock Name="txtItemCount" Text="0" VerticalAlignment="Center" FontWeight="Bold"/>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- Main DataGrid - columns generated dynamically from config -->
                <DataGrid Name="dgSoftware" Grid.Row="2" Margin="10"
                          AutoGenerateColumns="False"
                          IsReadOnly="False"
                          CanUserAddRows="False"
                          CanUserDeleteRows="False"
                          SelectionMode="Extended"
                          GridLinesVisibility="Horizontal"
                          HeadersVisibility="Column"
                          Background="White"
                          BorderBrush="#e0e0e0"
                          RowBackground="White"
                          AlternatingRowBackground="#fafafa">
                    <DataGrid.RowStyle>
                        <Style TargetType="DataGridRow">
                            <Style.Triggers>
                                <DataTrigger Binding="{Binding ChangeStatus}" Value="NEW">
                                    <Setter Property="Background" Value="#d4edda"/>
                                </DataTrigger>
                                <DataTrigger Binding="{Binding ChangeStatus}" Value="REMOVED">
                                    <Setter Property="Background" Value="#f8d7da"/>
                                </DataTrigger>
                            </Style.Triggers>
                        </Style>
                    </DataGrid.RowStyle>
                </DataGrid>

                <!-- Status Bar -->
                <Border Name="swStatusBorder" Grid.Row="3" Background="#333" Padding="10">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="200"/>
                        </Grid.ColumnDefinitions>

                        <TextBlock Name="txtStatus" Grid.Column="0" Text="Ready. Click 'Scan Software' to begin." Foreground="White"/>
                        <ProgressBar Name="pbProgress" Grid.Column="1" Height="15" Minimum="0" Maximum="100" Value="0"/>
                    </Grid>
                </Border>
            </Grid>
        </TabItem>

        <!-- ════════════════════════════════════════════════════ DRIVERS TAB ══ -->
        <TabItem Header="Drivers" Name="tabDrivers">
            <Grid Margin="16" Background="#1E1E2E">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Title bar -->
                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,14">
                    <TextBlock Text="&#x2699;" FontSize="22" Foreground="#89B4FA" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <StackPanel>
                        <TextBlock Text="Driver Manager" FontSize="20" FontWeight="Bold" Foreground="#CDD6F4"/>
                        <TextBlock Name="drv_txtSubtitle" Text="" FontSize="11" Foreground="#6C7086"/>
                    </StackPanel>
                </StackPanel>

                <!-- Search + filter toolbar -->
                <Grid Grid.Row="1" Margin="0,0,0,6">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <Grid Grid.Column="0" Margin="0,0,8,0">
                        <TextBox Name="drv_txtSearch" Style="{StaticResource SearchBox}" Height="32"/>
                        <TextBlock Text="Search devices, drivers, classes..." Foreground="#45475A" FontSize="12"
                                   IsHitTestVisible="False" VerticalAlignment="Center" Margin="12,0,0,0">
                            <TextBlock.Style>
                                <Style TargetType="TextBlock">
                                    <Setter Property="Visibility" Value="Collapsed"/>
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding Text, ElementName=drv_txtSearch}" Value="">
                                            <Setter Property="Visibility" Value="Visible"/>
                                        </DataTrigger>
                                    </Style.Triggers>
                                </Style>
                            </TextBlock.Style>
                        </TextBlock>
                    </Grid>

                    <ComboBox Name="drv_cboClass" Grid.Column="1" Style="{StaticResource ModernCombo}" Width="160" Height="32" Margin="0,0,8,0"/>

                    <ComboBox Name="drv_cboStatus" Grid.Column="2" Style="{StaticResource ModernCombo}" Width="120" Height="32" Margin="0,0,8,0">
                        <ComboBoxItem Content="All Statuses" IsSelected="True"/>
                        <ComboBoxItem Content="OK"/>
                        <ComboBoxItem Content="Error"/>
                        <ComboBoxItem Content="Degraded"/>
                        <ComboBoxItem Content="Unknown"/>
                    </ComboBox>

                    <ComboBox Name="drv_cboPresent" Grid.Column="3" Style="{StaticResource ModernCombo}" Width="130" Height="32" Margin="0,0,8,0">
                        <ComboBoxItem Content="All Devices" IsSelected="True"/>
                        <ComboBoxItem Content="Present Only"/>
                        <ComboBoxItem Content="Absent Only"/>
                    </ComboBox>

                    <Button Name="drv_btnRefresh" Grid.Column="4" Content="Refresh" Style="{StaticResource ModernButton}" Height="32"/>
                </Grid>

                <!-- Action buttons toolbar -->
                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,8">
                    <Button Name="drv_btnBackup"       Content="Backup Drivers"       Style="{StaticResource SuccessButton}" Height="30" Margin="0,0,8,0"/>
                    <Button Name="drv_btnImportFile"   Content="Import Driver (.inf)"  Style="{StaticResource AccentButton}"  Height="30" Margin="0,0,8,0"/>
                    <Button Name="drv_btnImportFolder" Content="Import Driver Folder"  Style="{StaticResource AccentButton}"  Height="30" Margin="0,0,16,0"/>
                    <Button Name="drv_btnExportCsv"    Content="Export CSV"            Style="{StaticResource ModernButton}"  Height="30" Margin="0,0,8,0"/>
                    <Button Name="drv_btnExportHtml"   Content="Export HTML"           Style="{StaticResource ModernButton}"  Height="30" Margin="0,0,8,0"/>
                    <Button Name="drv_btnCompare"      Content="Compare Baseline"      Style="{StaticResource ModernButton}"  Height="30" Margin="0,0,8,0"/>
                    <Button Name="drv_btnCopySelected" Content="Copy Row"              Style="{StaticResource ModernButton}"  Height="30"/>
                </StackPanel>

                <!-- Main DataGrid -->
                <DataGrid Name="drv_dgDevices" Grid.Row="3" Style="{StaticResource ModernGrid}" Margin="0,0,0,8">
                    <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                            <Setter Property="Background" Value="#313244"/>
                            <Setter Property="Foreground" Value="#89B4FA"/>
                            <Setter Property="FontWeight" Value="SemiBold"/>
                            <Setter Property="FontSize"   Value="11"/>
                            <Setter Property="Padding"    Value="10,6"/>
                            <Setter Property="BorderThickness" Value="0,0,1,0"/>
                            <Setter Property="BorderBrush" Value="#45475A"/>
                        </Style>
                    </DataGrid.ColumnHeaderStyle>
                    <DataGrid.RowStyle>
                        <Style TargetType="DataGridRow">
                            <Setter Property="Background" Value="Transparent"/>
                            <Style.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Background" Value="#313244"/>
                                </Trigger>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter Property="Background" Value="#45475A"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </DataGrid.RowStyle>
                    <DataGrid.CellStyle>
                        <Style TargetType="DataGridCell">
                            <Setter Property="BorderThickness" Value="0"/>
                            <Setter Property="Padding"         Value="6,4"/>
                            <Setter Property="Foreground"      Value="#CDD6F4"/>
                            <Style.Triggers>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter Property="Background"      Value="Transparent"/>
                                    <Setter Property="BorderThickness" Value="0"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </DataGrid.CellStyle>
                    <DataGrid.Columns>
                        <DataGridTemplateColumn Header=" " Width="30" CanUserSort="True" SortMemberPath="Status">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <Ellipse Width="8" Height="8" HorizontalAlignment="Center">
                                        <Ellipse.Style>
                                            <Style TargetType="Ellipse">
                                                <Setter Property="Fill" Value="#6C7086"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Status}" Value="OK">
                                                        <Setter Property="Fill" Value="#A6E3A1"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Error">
                                                        <Setter Property="Fill" Value="#F38BA8"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Degraded">
                                                        <Setter Property="Fill" Value="#FAB387"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </Ellipse.Style>
                                    </Ellipse>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                        <DataGridTextColumn Header="Device Name"     Binding="{Binding FriendlyName}"  Width="220"/>
                        <DataGridTextColumn Header="Class"           Binding="{Binding Class}"          Width="110"/>
                        <DataGridTextColumn Header="Status"          Binding="{Binding Status}"         Width="70"/>
                        <DataGridTextColumn Header="Driver Version"  Binding="{Binding DriverVersion}"  Width="110"/>
                        <DataGridTextColumn Header="Driver Provider" Binding="{Binding DriverProvider}" Width="130"/>
                        <DataGridTextColumn Header="Driver Date"     Binding="{Binding DriverDate}"     Width="95"/>
                        <DataGridTextColumn Header="INF File"        Binding="{Binding DriverInfPath}"  Width="160"/>
                        <DataGridTextColumn Header="Present"         Binding="{Binding Present}"        Width="65"/>
                        <DataGridTextColumn Header="Instance ID"     Binding="{Binding InstanceId}"     Width="*"/>
                    </DataGrid.Columns>
                </DataGrid>

                <!-- Detail panel -->
                <Border Grid.Row="4" Background="#181825" CornerRadius="6" Padding="14,10"
                        Margin="0,0,0,10" Visibility="Collapsed" Name="drv_pnlDetail">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0">
                            <TextBlock Text="DEVICE" Foreground="#6C7086" FontSize="10" FontWeight="Bold" Margin="0,0,0,4"/>
                            <TextBlock Name="drv_detailName"     Foreground="#CDD6F4" FontSize="12" TextWrapping="Wrap"/>
                            <TextBlock Name="drv_detailClass"    Foreground="#A6ADC8" FontSize="11" Margin="0,2,0,0"/>
                            <TextBlock Name="drv_detailStatus"   Foreground="#A6ADC8" FontSize="11" Margin="0,2,0,0"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Margin="16,0">
                            <TextBlock Text="DRIVER" Foreground="#6C7086" FontSize="10" FontWeight="Bold" Margin="0,0,0,4"/>
                            <TextBlock Name="drv_detailVersion"  Foreground="#CDD6F4" FontSize="12"/>
                            <TextBlock Name="drv_detailProvider" Foreground="#A6ADC8" FontSize="11" Margin="0,2,0,0"/>
                            <TextBlock Name="drv_detailDate"     Foreground="#A6ADC8" FontSize="11" Margin="0,2,0,0"/>
                            <TextBlock Name="drv_detailInf"      Foreground="#A6ADC8" FontSize="11" Margin="0,2,0,0" TextWrapping="Wrap"/>
                        </StackPanel>
                        <StackPanel Grid.Column="2">
                            <TextBlock Text="INSTANCE ID" Foreground="#6C7086" FontSize="10" FontWeight="Bold" Margin="0,0,0,4"/>
                            <TextBlock Name="drv_detailInstanceId" Foreground="#89B4FA" FontSize="11" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- Status bar -->
                <TextBlock Name="drv_txtStatus" Grid.Row="5"
                           Foreground="#6C7086" FontSize="11" VerticalAlignment="Center"
                           Text="Click 'Refresh' to load drivers."/>
            </Grid>
        </TabItem>

        <!-- ════════════════════════════════════════════════════ SERVICES TAB ══ -->
        <TabItem Header="Services" Name="tabServices">
            <Grid Margin="16" Background="#1E1E2E">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>   <!-- Title -->
                    <RowDefinition Height="Auto"/>   <!-- Search + filters -->
                    <RowDefinition Height="Auto"/>   <!-- Action buttons -->
                    <RowDefinition Height="*"/>      <!-- Main grid -->
                    <RowDefinition Height="Auto"/>   <!-- Detail panel -->
                    <RowDefinition Height="Auto"/>   <!-- Status bar -->
                </Grid.RowDefinitions>

                <!-- ── TITLE ── -->
                <Grid Grid.Row="0" Margin="0,0,0,14">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                        <TextBlock Text="&#x1F6E1;" FontSize="22" Foreground="#F38BA8" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <StackPanel>
                            <TextBlock Text="Service Manager" FontSize="20" FontWeight="Bold" Foreground="#CDD6F4"/>
                            <TextBlock Name="svc_txtSubtitle" Text="" FontSize="11" Foreground="#6C7086"/>
                        </StackPanel>
                    </StackPanel>
                    <Border Grid.Column="1" Background="#181825" CornerRadius="6" Padding="10,6"
                            HorizontalAlignment="Right" VerticalAlignment="Center">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBlock Text="Query:" Foreground="#6C7086" FontSize="11" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <CheckBox Name="svc_chkStartupType" Content="Startup type (WMI)" IsChecked="True"
                                      Foreground="#CDD6F4" FontSize="11" VerticalAlignment="Center" Margin="0,0,14,0"/>
                            <CheckBox Name="svc_chkSvcControl" Content="Service control permissions" IsChecked="True"
                                      Foreground="#CDD6F4" FontSize="11" VerticalAlignment="Center" Margin="0,0,14,0"/>
                            <CheckBox Name="svc_chkExeAcl" Content="Executable ACL analysis" IsChecked="True"
                                      Foreground="#CDD6F4" FontSize="11" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                </Grid>

                <!-- ── SEARCH + FILTERS ── -->
                <Grid Grid.Row="1" Margin="0,0,0,6">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <Grid Grid.Column="0" Margin="0,0,8,0">
                        <TextBox Name="svc_txtSearch" Style="{StaticResource SearchBox}" Height="32" VerticalContentAlignment="Center"/>
                        <TextBlock Text="Search services, executables, accounts..."
                                   Foreground="#45475A" FontSize="12" IsHitTestVisible="False"
                                   VerticalAlignment="Center" Margin="12,0,0,0">
                            <TextBlock.Style>
                                <Style TargetType="TextBlock">
                                    <Setter Property="Visibility" Value="Collapsed"/>
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding Text, ElementName=svc_txtSearch}" Value="">
                                            <Setter Property="Visibility" Value="Visible"/>
                                        </DataTrigger>
                                    </Style.Triggers>
                                </Style>
                            </TextBlock.Style>
                        </TextBlock>
                    </Grid>

                    <ComboBox Name="svc_cboState" Grid.Column="1" Style="{StaticResource ModernCombo}"
                              Width="130" Height="32" Margin="0,0,8,0">
                        <ComboBoxItem Content="All States"  IsSelected="True"/>
                        <ComboBoxItem Content="Running"/>
                        <ComboBoxItem Content="Stopped"/>
                        <ComboBoxItem Content="Paused"/>
                    </ComboBox>

                    <ComboBox Name="svc_cboStartup" Grid.Column="2" Style="{StaticResource ModernCombo}"
                              Width="140" Height="32" Margin="0,0,8,0">
                        <ComboBoxItem Content="All Startup Types" IsSelected="True"/>
                        <ComboBoxItem Content="Automatic"/>
                        <ComboBoxItem Content="Automatic (Delayed)"/>
                        <ComboBoxItem Content="Manual"/>
                        <ComboBoxItem Content="Disabled"/>
                    </ComboBox>

                    <ComboBox Name="svc_cboRisk" Grid.Column="3" Style="{StaticResource ModernCombo}"
                              Width="145" Height="32" Margin="0,0,8,0">
                        <ComboBoxItem Content="All Services"     IsSelected="True"/>
                        <ComboBoxItem Content="! Risky Exe ACL"/>
                        <ComboBoxItem Content="! Weak Svc Control"/>
                    </ComboBox>

                    <Button Name="svc_btnRefresh" Grid.Column="4" Content="&#x21BA;  Refresh"
                            Style="{StaticResource ModernButton}" Height="32"/>
                </Grid>

                <!-- ── ACTION BUTTONS ── -->
                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,8">
                    <Button Name="svc_btnStart"      Content="&#x25B6;  Start"            Style="{StaticResource ModernButton}"  Height="30" Margin="0,0,8,0"/>
                    <Button Name="svc_btnStop"       Content="&#x25A0;  Stop"             Style="{StaticResource DangerButton}"  Height="30" Margin="0,0,8,0"/>
                    <Button Name="svc_btnRestart"    Content="&#x21BA;  Restart"          Style="{StaticResource ModernButton}"  Height="30" Margin="0,0,16,0"/>
                    <Button Name="svc_btnExportCsv"  Content="Export CSV"                 Style="{StaticResource ModernButton}"  Height="30" Margin="0,0,8,0"/>
                    <Button Name="svc_btnExportHtml" Content="Export HTML"                Style="{StaticResource ModernButton}"  Height="30" Margin="0,0,8,0"/>
                    <Button Name="svc_btnCompare"    Content="&#x21C4;  Compare Baseline" Style="{StaticResource ModernButton}"  Height="30" Margin="0,0,8,0"/>
                    <Button Name="svc_btnCopyRow"    Content="Copy Row"                   Style="{StaticResource ModernButton}"  Height="30"/>
                </StackPanel>

                <!-- ── MAIN GRID ── -->
                <DataGrid Name="svc_dgServices" Grid.Row="3" Style="{StaticResource ModernGrid}" Margin="0,0,0,8">
                    <DataGrid.ColumnHeaderStyle>
                        <Style TargetType="DataGridColumnHeader">
                            <Setter Property="Background"   Value="#313244"/>
                            <Setter Property="Foreground"   Value="#F38BA8"/>
                            <Setter Property="FontWeight"   Value="SemiBold"/>
                            <Setter Property="FontSize"     Value="11"/>
                            <Setter Property="Padding"      Value="10,6"/>
                            <Setter Property="BorderThickness" Value="0,0,1,0"/>
                            <Setter Property="BorderBrush"  Value="#45475A"/>
                        </Style>
                    </DataGrid.ColumnHeaderStyle>
                    <DataGrid.RowStyle>
                        <Style TargetType="DataGridRow">
                            <Setter Property="Background" Value="Transparent"/>
                            <Style.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Background" Value="#313244"/>
                                </Trigger>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter Property="Background" Value="#45475A"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </DataGrid.RowStyle>
                    <DataGrid.CellStyle>
                        <Style TargetType="DataGridCell">
                            <Setter Property="BorderThickness" Value="0"/>
                            <Setter Property="Padding"         Value="6,4"/>
                            <Setter Property="Foreground"      Value="#CDD6F4"/>
                            <Style.Triggers>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter Property="Background"      Value="Transparent"/>
                                    <Setter Property="BorderThickness" Value="0"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </DataGrid.CellStyle>

                    <DataGrid.Columns>
                        <!-- State dot -->
                        <DataGridTemplateColumn Header="  " Width="28" SortMemberPath="State" CanUserSort="True">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <Ellipse Width="8" Height="8" HorizontalAlignment="Center">
                                        <Ellipse.Style>
                                            <Style TargetType="Ellipse">
                                                <Setter Property="Fill" Value="#6C7086"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding State}" Value="Running">
                                                        <Setter Property="Fill" Value="#A6E3A1"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding State}" Value="Stopped">
                                                        <Setter Property="Fill" Value="#F38BA8"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding State}" Value="Paused">
                                                        <Setter Property="Fill" Value="#FAB387"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </Ellipse.Style>
                                    </Ellipse>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>

                        <DataGridTextColumn Header="Service Name"    Binding="{Binding Name}"           Width="160"/>
                        <DataGridTextColumn Header="Display Name"   Binding="{Binding DisplayName}"    Width="200"/>
                        <DataGridTextColumn Header="State"          Binding="{Binding State}"          Width="75"/>
                        <DataGridTextColumn Header="Startup Type"   Binding="{Binding StartupType}"    Width="135"/>
                        <DataGridTextColumn Header="Log On As"      Binding="{Binding LogOnAs}"        Width="140"/>

                        <!-- Svc Control Risk -->
                        <DataGridTemplateColumn Header="Svc Control" Width="95" SortMemberPath="ServiceControlRisk" CanUserSort="True">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding ServiceControlRisk}" HorizontalAlignment="Center" FontSize="11">
                                        <TextBlock.Style>
                                            <Style TargetType="TextBlock">
                                                <Setter Property="Foreground" Value="#A6E3A1"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding ServiceControlRisk}" Value="! Weak">
                                                        <Setter Property="Foreground" Value="#FAB387"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>

                        <!-- Exe Write Risk -->
                        <DataGridTemplateColumn Header="Exe Write ACL" Width="105" SortMemberPath="ExeWriteRisk" CanUserSort="True">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBlock Text="{Binding ExeWriteRisk}" HorizontalAlignment="Center" FontSize="11">
                                        <TextBlock.Style>
                                            <Style TargetType="TextBlock">
                                                <Setter Property="Foreground" Value="#A6E3A1"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding ExeWriteRisk}" Value="! Risky">
                                                        <Setter Property="Foreground" Value="#F38BA8"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding ExeWriteRisk}" Value="N/A">
                                                        <Setter Property="Foreground" Value="#6C7086"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>

                        <DataGridTextColumn Header="Executable Path"  Binding="{Binding ExePath}"       Width="*"/>
                    </DataGrid.Columns>
                </DataGrid>

                <!-- ── DETAIL PANEL ── -->
                <Border Name="svc_pnlDetail" Grid.Row="4" Background="#181825" CornerRadius="6"
                        Padding="16,12" Margin="0,0,0,10" Visibility="Collapsed">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                        <!-- Service info -->
                        <StackPanel Grid.Row="0" Grid.Column="0" Margin="0,0,16,8">
                            <TextBlock Text="SERVICE" Foreground="#6C7086" FontSize="10" FontWeight="Bold" Margin="0,0,0,4"/>
                            <TextBlock Name="svc_detailName"        Foreground="#CDD6F4" FontSize="13" FontWeight="SemiBold"/>
                            <TextBlock Name="svc_detailDisplayName" Foreground="#A6ADC8" FontSize="11" Margin="0,2,0,0"/>
                            <TextBlock Name="svc_detailState"       Foreground="#A6ADC8" FontSize="11" Margin="0,2,0,0"/>
                            <TextBlock Name="svc_detailStartup"     Foreground="#A6ADC8" FontSize="11" Margin="0,2,0,0"/>
                            <TextBlock Name="svc_detailLogOn"       Foreground="#A6ADC8" FontSize="11" Margin="0,2,0,0"/>
                        </StackPanel>

                        <!-- Exe path -->
                        <StackPanel Grid.Row="0" Grid.Column="1" Margin="0,0,16,8">
                            <TextBlock Text="EXECUTABLE" Foreground="#6C7086" FontSize="10" FontWeight="Bold" Margin="0,0,0,4"/>
                            <TextBlock Name="svc_detailExePath" Foreground="#89B4FA" FontSize="11" TextWrapping="Wrap"/>
                        </StackPanel>

                        <!-- Security summary -->
                        <StackPanel Grid.Row="0" Grid.Column="2" Margin="0,0,0,8">
                            <TextBlock Text="SECURITY" Foreground="#6C7086" FontSize="10" FontWeight="Bold" Margin="0,0,0,4"/>
                            <TextBlock Name="svc_detailSvcControl" Foreground="#CDD6F4" FontSize="11" TextWrapping="Wrap"/>
                            <TextBlock Name="svc_detailExeAcl"     Foreground="#CDD6F4" FontSize="11" TextWrapping="Wrap" Margin="0,6,0,0"/>
                        </StackPanel>

                        <!-- Full ACL table -->
                        <StackPanel Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="3">
                            <TextBlock Text="EXECUTABLE FILE PERMISSIONS" Foreground="#6C7086" FontSize="10" FontWeight="Bold" Margin="0,4,0,6"/>
                            <DataGrid Name="svc_dgAcl"
                                      Background="#1E1E2E" Foreground="#CDD6F4" BorderThickness="0"
                                      RowBackground="#1E1E2E" AlternatingRowBackground="#181825"
                                      GridLinesVisibility="None" HeadersVisibility="Column"
                                      AutoGenerateColumns="False" IsReadOnly="True"
                                      MaxHeight="130" FontSize="11"
                                      HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto">
                                <DataGrid.ColumnHeaderStyle>
                                    <Style TargetType="DataGridColumnHeader">
                                        <Setter Property="Background" Value="#313244"/>
                                        <Setter Property="Foreground" Value="#F38BA8"/>
                                        <Setter Property="FontWeight" Value="SemiBold"/>
                                        <Setter Property="Padding"    Value="8,4"/>
                                        <Setter Property="BorderThickness" Value="0"/>
                                    </Style>
                                </DataGrid.ColumnHeaderStyle>
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Identity"        Binding="{Binding Identity}"    Width="220"/>
                                    <DataGridTextColumn Header="Rights"          Binding="{Binding Rights}"      Width="200"/>
                                    <DataGridTextColumn Header="Type"            Binding="{Binding AceType}"     Width="80"/>
                                    <DataGridTextColumn Header="Write/Modify ?"  Binding="{Binding IsRisky}"     Width="110"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- ── STATUS BAR ── -->
                <Grid Grid.Row="5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Name="svc_txtStatus" Grid.Column="0"
                               Foreground="#6C7086" FontSize="11" VerticalAlignment="Center"
                               Text="Click 'Refresh' to load services."/>
                    <TextBlock Name="svc_txtLegend" Grid.Column="1"
                               Foreground="#6C7086" FontSize="10" VerticalAlignment="Center"
                               Text="&#x25CF; Running  &#x25CF; Stopped  &#x25CF; Paused   |   ! Weak = non-Admins can control   ! Risky = non-Admins can write exe"/>
                </Grid>
            </Grid>
        </TabItem>

    </TabControl>
</Window>
"@
#endregion

#region Settings Window
function Show-SettingsWindow {
    [xml]$settingsXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SoftwareLister Settings" Height="600" Width="500"
        WindowStartupLocation="CenterOwner"
        Background="#F5F5F5">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Settings" FontSize="20" FontWeight="Bold" Margin="0,0,0,15"/>

        <TabControl Grid.Row="1">
            <TabItem Header="Data Sources">
                <StackPanel Margin="15">
                    <TextBlock Text="Select which sources to scan:" FontWeight="Bold" Margin="0,0,0,10"/>
                    <CheckBox Name="chkRegistry" Content="Windows Registry (Traditional Programs)" Margin="0,5"/>
                    <CheckBox Name="chkAppx" Content="AppX Packages (Microsoft Store / UWP)" Margin="0,5"/>
                    <CheckBox Name="chkWinget" Content="Winget (Windows Package Manager)" Margin="0,5"/>

                    <Separator Margin="0,20"/>

                    <TextBlock Text="Display Options:" FontWeight="Bold" Margin="0,10,0,10"/>
                    <CheckBox Name="chkShowSystem" Content="Show System Components" Margin="0,5"/>
                    <CheckBox Name="chkShowFrameworks" Content="Show Frameworks/Runtimes" Margin="0,5"/>

                    <Separator Margin="0,20"/>

                    <TextBlock Text="Theme:" FontWeight="Bold" Margin="0,10,0,10"/>
                    <RadioButton Name="rbThemeSystem" Content="Follow Windows System Setting" Margin="0,5" GroupName="Theme"/>
                    <RadioButton Name="rbThemeLight" Content="Light Mode" Margin="0,5" GroupName="Theme"/>
                    <RadioButton Name="rbThemeDark" Content="Dark Mode" Margin="0,5" GroupName="Theme"/>
                </StackPanel>
            </TabItem>

            <TabItem Header="Properties">
                <Grid Margin="15">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Text="Select properties to include in reports:" FontWeight="Bold" Margin="0,0,0,10"/>
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <StackPanel Name="spProperties">
                            <!-- Properties will be added dynamically -->
                        </StackPanel>
                    </ScrollViewer>
                </Grid>
            </TabItem>

            <TabItem Header="Exclude Patterns">
                <Grid Margin="15">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Text="Regex patterns to exclude (one per line):" FontWeight="Bold" Margin="0,0,0,10"/>
                    <TextBox Name="txtExcludePatterns" Grid.Row="1" AcceptsReturn="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" FontFamily="Consolas"/>
                    <TextBlock Grid.Row="2" Text="Examples: ^Microsoft Visual C\+\+.*Redistributable"
                               Foreground="Gray" FontSize="11" Margin="0,5,0,0"/>
                </Grid>
            </TabItem>
        </TabControl>

        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button Name="btnSaveSettings" Content="Save" Padding="20,8" Margin="5,0" Background="#0078d4" Foreground="White" BorderThickness="0"/>
            <Button Name="btnCancelSettings" Content="Cancel" Padding="20,8" Margin="5,0"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $settingsReader = New-Object System.Xml.XmlNodeReader $settingsXaml
    $settingsWindow = [Windows.Markup.XamlReader]::Load($settingsReader)
    $settingsWindow.Owner = $script:Window

    $chkRegistry = $settingsWindow.FindName("chkRegistry")
    $chkAppx = $settingsWindow.FindName("chkAppx")
    $chkWinget = $settingsWindow.FindName("chkWinget")
    $chkShowSystem = $settingsWindow.FindName("chkShowSystem")
    $chkShowFrameworks = $settingsWindow.FindName("chkShowFrameworks")
    $rbThemeSystem = $settingsWindow.FindName("rbThemeSystem")
    $rbThemeLight = $settingsWindow.FindName("rbThemeLight")
    $rbThemeDark = $settingsWindow.FindName("rbThemeDark")
    $spProperties = $settingsWindow.FindName("spProperties")
    $txtExcludePatterns = $settingsWindow.FindName("txtExcludePatterns")
    $btnSaveSettings = $settingsWindow.FindName("btnSaveSettings")
    $btnCancelSettings = $settingsWindow.FindName("btnCancelSettings")

    $chkRegistry.IsChecked = $script:Config.dataSources.registry.enabled
    $chkAppx.IsChecked = $script:Config.dataSources.appx.enabled
    $chkWinget.IsChecked = $script:Config.dataSources.winget.enabled
    $chkShowSystem.IsChecked = $script:Config.display.showSystemComponents
    $chkShowFrameworks.IsChecked = $script:Config.display.showFrameworks

    $currentTheme = if ($script:Config.display.theme) { $script:Config.display.theme.ToLower() } else { "system" }
    switch ($currentTheme) {
        "light" { $rbThemeLight.IsChecked = $true }
        "dark" { $rbThemeDark.IsChecked = $true }
        default { $rbThemeSystem.IsChecked = $true }
    }

    $propertyCheckboxes = @{}
    foreach ($prop in $script:Config.properties.PSObject.Properties) {
        $checkbox = New-Object System.Windows.Controls.CheckBox
        $checkbox.Content = "$($prop.Value.displayName) - $($prop.Value.description)"
        $checkbox.IsChecked = $prop.Value.enabled
        $checkbox.Margin = "0,5,0,0"
        $checkbox.Tag = $prop.Name

        if ($prop.Value.required) {
            $checkbox.IsEnabled = $false
            $checkbox.Content += " (Required)"
        }

        $spProperties.Children.Add($checkbox)
        $propertyCheckboxes[$prop.Name] = $checkbox
    }

    $txtExcludePatterns.Text = $script:Config.display.excludePatterns -join "`r`n"

    $btnSaveSettings.Add_Click({
        $script:Config.dataSources.registry.enabled = $chkRegistry.IsChecked
        $script:Config.dataSources.appx.enabled = $chkAppx.IsChecked
        $script:Config.dataSources.winget.enabled = $chkWinget.IsChecked

        $script:Config.display.showSystemComponents = $chkShowSystem.IsChecked
        $script:Config.display.showFrameworks = $chkShowFrameworks.IsChecked

        if ($rbThemeDark.IsChecked) {
            $script:Config.display.theme = "dark"
        } elseif ($rbThemeLight.IsChecked) {
            $script:Config.display.theme = "light"
        } else {
            $script:Config.display.theme = "system"
        }

        foreach ($propName in $propertyCheckboxes.Keys) {
            $script:Config.properties.$propName.enabled = $propertyCheckboxes[$propName].IsChecked
        }

        $patterns = $txtExcludePatterns.Text -split "`r?`n" | Where-Object { $_.Trim() }
        $script:Config.display.excludePatterns = @($patterns)

        if (Save-Configuration) {
            [System.Windows.MessageBox]::Show("Settings saved successfully.", "Settings", "OK", "Information")
            $settingsWindow.DialogResult = $true
            $settingsWindow.Close()
        }
    })

    $btnCancelSettings.Add_Click({
        $settingsWindow.DialogResult = $false
        $settingsWindow.Close()
    })

    return $settingsWindow.ShowDialog()
}
#endregion

#region Main Application
# Load configuration
if (-not (Load-Configuration)) {
    exit
}

# Create window
$reader = New-Object System.Xml.XmlNodeReader $xaml
$script:Window = [Windows.Markup.XamlReader]::Load($reader)

# ── Software tab controls ──
$btnScan        = $script:Window.FindName("btnScan")
$btnCompare     = $script:Window.FindName("btnCompare")
$btnExportCsv   = $script:Window.FindName("btnExportCsv")
$btnExportHtml  = $script:Window.FindName("btnExportHtml")
$btnExportJson  = $script:Window.FindName("btnExportJson")
$btnSettings    = $script:Window.FindName("btnSettings")
$txtSearch      = $script:Window.FindName("txtSearch")
$btnClearSearch = $script:Window.FindName("btnClearSearch")
$txtItemCount   = $script:Window.FindName("txtItemCount")
$dgSoftware     = $script:Window.FindName("dgSoftware")
$txtStatus      = $script:Window.FindName("txtStatus")
$pbProgress     = $script:Window.FindName("pbProgress")

# ── Driver tab controls ──
$drvControlNames = @(
    'drv_txtSubtitle','drv_txtSearch','drv_cboClass','drv_cboStatus','drv_cboPresent',
    'drv_btnRefresh','drv_btnBackup','drv_btnImportFile','drv_btnImportFolder',
    'drv_btnExportCsv','drv_btnExportHtml','drv_btnCompare','drv_btnCopySelected',
    'drv_dgDevices','drv_pnlDetail',
    'drv_detailName','drv_detailClass','drv_detailStatus','drv_detailVersion',
    'drv_detailProvider','drv_detailDate','drv_detailInf','drv_detailInstanceId',
    'drv_txtStatus'
)
foreach ($name in $drvControlNames) {
    $script:DrvC[$name] = $script:Window.FindName($name)
}

# ── Service tab controls ──
$svcControlNames = @(
    'svc_txtSubtitle','svc_txtSearch','svc_cboState','svc_cboStartup','svc_cboRisk',
    'svc_chkStartupType','svc_chkSvcControl','svc_chkExeAcl',
    'svc_btnRefresh','svc_btnStart','svc_btnStop','svc_btnRestart',
    'svc_btnExportCsv','svc_btnExportHtml','svc_btnCompare','svc_btnCopyRow',
    'svc_dgServices','svc_pnlDetail',
    'svc_detailName','svc_detailDisplayName','svc_detailState','svc_detailStartup','svc_detailLogOn',
    'svc_detailExePath','svc_detailSvcControl','svc_detailExeAcl','svc_dgAcl',
    'svc_txtStatus','svc_txtLegend'
)
foreach ($name in $svcControlNames) {
    $script:SvcC[$name] = $script:Window.FindName($name)
}

# Bind DataGrid and build columns from config
$dgSoftware.ItemsSource = $script:SoftwareList
Build-DataGridColumns -DataGrid $dgSoftware

# Apply theme
Apply-Theme -Window $script:Window

# ── Software tab event handlers ──
$btnScan.Add_Click({
    $btnScan.IsEnabled = $false
    $pbProgress.Value = 0

    try {
        $count = Invoke-SoftwareScan -ProgressBar $pbProgress -StatusText $txtStatus
        $txtItemCount.Text = $count.ToString()

        $btnCompare.IsEnabled = $true
        $btnExportCsv.IsEnabled = $true
        $btnExportHtml.IsEnabled = $true
        $btnExportJson.IsEnabled = $true
    } finally {
        $btnScan.IsEnabled = $true
    }
})

$btnCompare.Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Filter = "Supported Files|*.csv;*.json|CSV Files|*.csv|JSON Files|*.json"
    $dialog.Title = "Select Previous Scan Results"

    if ($dialog.ShowDialog()) {
        Compare-WithPrevious -PreviousFilePath $dialog.FileName
        $dgSoftware.Items.Refresh()
        $txtItemCount.Text = $script:SoftwareList.Count.ToString()
    }
})

$btnExportCsv.Add_Click({
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = "CSV Files|*.csv"
    $dialog.FileName = "SoftwareInventory_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd').csv"

    if ($dialog.ShowDialog()) {
        Export-ToCsv -FilePath $dialog.FileName
        $txtStatus.Text = "Exported to: $($dialog.FileName)"
        [System.Windows.MessageBox]::Show("Exported successfully to:`n$($dialog.FileName)", "Export Complete", "OK", "Information")
    }
})

$btnExportJson.Add_Click({
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = "JSON Files|*.json"
    $dialog.FileName = "SoftwareInventory_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd').json"

    if ($dialog.ShowDialog()) {
        Export-ToJson -FilePath $dialog.FileName
        $txtStatus.Text = "Exported to: $($dialog.FileName)"
        [System.Windows.MessageBox]::Show("Exported successfully to:`n$($dialog.FileName)", "Export Complete", "OK", "Information")
    }
})

$btnExportHtml.Add_Click({
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = "HTML Files|*.html"
    $dialog.FileName = "SoftwareInventory_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd').html"

    if ($dialog.ShowDialog()) {
        Export-ToHtml -FilePath $dialog.FileName
        $txtStatus.Text = "Exported to: $($dialog.FileName)"

        $result = [System.Windows.MessageBox]::Show("Exported successfully to:`n$($dialog.FileName)`n`nWould you like to open it?", "Export Complete", "YesNo", "Information")
        if ($result -eq "Yes") {
            Start-Process $dialog.FileName
        }
    }
})

$btnSettings.Add_Click({
    $result = Show-SettingsWindow
    if ($result -eq $true) {
        Build-DataGridColumns -DataGrid $dgSoftware
        Apply-Theme -Window $script:Window
    }
})

$txtSearch.Add_TextChanged({
    $searchText = $txtSearch.Text.ToLower()
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:SoftwareList)

    if ([string]::IsNullOrWhiteSpace($searchText)) {
        $view.Filter = $null
    } else {
        $view.Filter = {
            param($item)
            $name = if ($item.CustomName) { $item.CustomName } else { $item.Name }
            return (
                $name.ToLower().Contains($searchText) -or
                $item.Name.ToLower().Contains($searchText) -or
                ($item.Publisher -and $item.Publisher.ToLower().Contains($searchText)) -or
                ($item.Version -and $item.Version.ToLower().Contains($searchText))
            )
        }
    }

    $count = 0
    foreach ($item in $view) { $count++ }
    $txtItemCount.Text = $count.ToString()
})

$btnClearSearch.Add_Click({
    $txtSearch.Text = ""
})

$dgSoftware.Add_CellEditEnding({
    param($sender, $e)

    if ($e.Column.Header -eq "Custom Name") {
        $item = $e.Row.Item
        $textBox = $e.EditingElement
        $newName = $textBox.Text

        $key = "$($item.Source)::$($item.UniqueId)"
        if (-not $key -or $key -eq "::") {
            $key = "$($item.Source)::$($item.Name)"
        }

        if ([string]::IsNullOrWhiteSpace($newName)) {
            if ($script:Config.customNames.PSObject.Properties[$key]) {
                $script:Config.customNames.PSObject.Properties.Remove($key)
            }
        } else {
            if (-not $script:Config.customNames) {
                $script:Config.customNames = [PSCustomObject]@{}
            }
            $script:Config.customNames | Add-Member -NotePropertyName $key -NotePropertyValue $newName -Force
        }

        Save-Configuration
    }
})

# ── Driver tab event handlers ──
$script:DrvC['drv_txtSearch'].Add_TextChanged({   Apply-DriverFilters })
$script:DrvC['drv_cboClass'].Add_SelectionChanged({   Apply-DriverFilters })
$script:DrvC['drv_cboStatus'].Add_SelectionChanged({  Apply-DriverFilters })
$script:DrvC['drv_cboPresent'].Add_SelectionChanged({ Apply-DriverFilters })

$script:DrvC['drv_btnRefresh'].Add_Click({ Load-DriverData })
$script:DrvC['drv_btnBackup'].Add_Click({ Backup-Drivers })
$script:DrvC['drv_btnImportFile'].Add_Click({ Import-SingleDriver })
$script:DrvC['drv_btnImportFolder'].Add_Click({ Import-DriverFolder })

$script:DrvC['drv_dgDevices'].Add_SelectionChanged({
    $row = $script:DrvC['drv_dgDevices'].SelectedItem
    if ($null -eq $row) {
        $script:DrvC['drv_pnlDetail'].Visibility = 'Collapsed'
        return
    }
    $script:DrvC['drv_pnlDetail'].Visibility       = 'Visible'
    $script:DrvC['drv_detailName'].Text             = $row.FriendlyName
    $script:DrvC['drv_detailClass'].Text            = "Class: $($row.Class)"
    $script:DrvC['drv_detailStatus'].Text           = "Status: $($row.Status)  |  Present: $($row.Present)"
    $script:DrvC['drv_detailVersion'].Text          = "Version: $($row.DriverVersion)"
    $script:DrvC['drv_detailProvider'].Text         = "Provider: $($row.DriverProvider)"
    $script:DrvC['drv_detailDate'].Text             = "Date: $($row.DriverDate)"
    $script:DrvC['drv_detailInf'].Text              = "INF: $($row.DriverInfPath)"
    $script:DrvC['drv_detailInstanceId'].Text       = $row.InstanceId
})

$script:DrvC['drv_btnExportCsv'].Add_Click({
    $defaultName = "Drivers_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $path = Export-ReportCsv -Data $script:DrvC['drv_dgDevices'].ItemsSource -DefaultName $defaultName
    if ($path) { $script:DrvC['drv_txtStatus'].Text = "Exported CSV: $path" }
})

$script:DrvC['drv_btnExportHtml'].Add_Click({
    $defaultName = "Drivers_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $path = Export-ReportHtml `
        -Data            $script:DrvC['drv_dgDevices'].ItemsSource `
        -Title           "Driver Manager Report" `
        -DefaultName     $defaultName `
        -HighlightField  'Status' `
        -HighlightValues @('Error','Degraded','Unknown')
    if ($path) {
        $script:DrvC['drv_txtStatus'].Text = "Exported HTML: $path"
        Start-Process $path
    }
})

$script:DrvC['drv_btnCompare'].Add_Click({
    if ($script:AllDevices.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Device list has not loaded yet - please wait and try again.",
            "Not Ready", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    Invoke-CompareFlow `
        -CurrentData   $script:AllDevices `
        -KeyField      'InstanceId' `
        -CompareFields @('FriendlyName','Class','Status','Present','DriverVersion','DriverProvider','DriverDate','DriverInfPath') `
        -WindowTitle   'Driver Manager - Baseline Comparison'
})

$script:DrvC['drv_btnCopySelected'].Add_Click({
    $row = $script:DrvC['drv_dgDevices'].SelectedItem
    if ($row) {
        $text = "Name: $($row.FriendlyName)`nClass: $($row.Class)`nStatus: $($row.Status)`nDriver Version: $($row.DriverVersion)`nDriver Provider: $($row.DriverProvider)`nDriver Date: $($row.DriverDate)`nINF: $($row.DriverInfPath)`nInstance ID: $($row.InstanceId)"
        [System.Windows.Clipboard]::SetText($text)
        $script:DrvC['drv_txtStatus'].Text = "Row copied to clipboard."
    }
})

# ── Service tab event handlers ──
$script:SvcC['svc_txtSearch'].Add_TextChanged({   Apply-ServiceFilters })
$script:SvcC['svc_cboState'].Add_SelectionChanged({   Apply-ServiceFilters })
$script:SvcC['svc_cboStartup'].Add_SelectionChanged({ Apply-ServiceFilters })
$script:SvcC['svc_cboRisk'].Add_SelectionChanged({    Apply-ServiceFilters })

$script:SvcC['svc_btnRefresh'].Add_Click({ Load-ServiceData })
$script:SvcC['svc_btnStart'].Add_Click({   Invoke-ServiceAction -Action 'Start' })
$script:SvcC['svc_btnStop'].Add_Click({    Invoke-ServiceAction -Action 'Stop' })
$script:SvcC['svc_btnRestart'].Add_Click({ Invoke-ServiceAction -Action 'Restart' })

$script:SvcC['svc_dgServices'].Add_SelectionChanged({
    Update-ServiceDetailPanel -row $script:SvcC['svc_dgServices'].SelectedItem
})

$script:SvcC['svc_btnExportCsv'].Add_Click({
    $defaultName = "Services_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $exportData  = $script:SvcC['svc_dgServices'].ItemsSource |
        Select-Object Name, DisplayName, State, StartupType, LogOnAs,
                      ServiceControlRisk, ServiceControlDetail,
                      ExeWriteRisk, ExeAclSummary, ExePath
    $path = Export-ReportCsv -Data $exportData -DefaultName $defaultName
    if ($path) { $script:SvcC['svc_txtStatus'].Text = "✅  Exported CSV: $path" }
})

$script:SvcC['svc_btnExportHtml'].Add_Click({
    $defaultName = "Services_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $exportData  = $script:SvcC['svc_dgServices'].ItemsSource |
        Select-Object Name, DisplayName, State, StartupType, LogOnAs,
                      ServiceControlRisk, ServiceControlDetail,
                      ExeWriteRisk, ExeAclSummary, ExePath
    $path = Export-ReportHtml `
        -Data            $exportData `
        -Title           "Service Manager Report" `
        -DefaultName     $defaultName `
        -HighlightField  'ExeWriteRisk' `
        -HighlightValues @('! Risky')
    if ($path) {
        $script:SvcC['svc_txtStatus'].Text = "✅  Exported HTML: $path"
        Start-Process $path
    }
})

$script:SvcC['svc_btnCompare'].Add_Click({
    if ($script:AllServices.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Service list has not loaded yet - please wait and try again.",
            "Not Ready", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    $flatCurrent = $script:AllServices |
        Select-Object Name, DisplayName, State, StartupType, LogOnAs,
                      ServiceControlRisk, ServiceControlDetail,
                      ExeWriteRisk, ExeAclSummary, ExePath
    Invoke-CompareFlow `
        -CurrentData   $flatCurrent `
        -KeyField      'Name' `
        -CompareFields @('State','StartupType','LogOnAs','ServiceControlRisk','ExeWriteRisk','ExePath') `
        -WindowTitle   'Service Manager - Baseline Comparison'
})

$script:SvcC['svc_btnCopyRow'].Add_Click({
    $row = $script:SvcC['svc_dgServices'].SelectedItem
    if ($row) {
        $text = @"
Service Name    : $($row.Name)
Display Name    : $($row.DisplayName)
State           : $($row.State)
Startup Type    : $($row.StartupType)
Log On As       : $($row.LogOnAs)
Executable      : $($row.ExePath)
Svc Control     : $($row.ServiceControlRisk) - $($row.ServiceControlDetail)
Exe Write ACL   : $($row.ExeWriteRisk) - $($row.ExeAclSummary)
"@
        [System.Windows.Clipboard]::SetText($text)
        $script:SvcC['svc_txtStatus'].Text = "Row copied to clipboard."
    }
})


# Apply title bar theme when window handle is available
$script:Window.Add_Loaded({
    $theme = Get-CurrentTheme
    try {
        $windowHelper = New-Object System.Windows.Interop.WindowInteropHelper($script:Window)
        $hwnd = $windowHelper.Handle
        if ($hwnd -ne [IntPtr]::Zero) {
            [DwmApi]::SetDarkTitleBar($hwnd, ($theme -eq "dark"))
        }
    } catch { }
})

# Show window
$script:Window.ShowDialog()
#endregion
