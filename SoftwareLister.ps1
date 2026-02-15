#Requires -Version 5.1
<#
.SYNOPSIS
    SoftwareLister - Comprehensive software inventory tool for Windows
.DESCRIPTION
    Scans and lists all installed software from:
    - Windows Registry (traditional programs)
    - AppX packages (Microsoft Store / UWP apps)
    - Winget (Windows Package Manager)

    Features:
    - Export to CSV, HTML, JSON formats
    - Compare with previous scans to detect changes
    - Custom name overrides for better reporting
    - Configurable properties via config.json
.NOTES
    Author: SoftwareLister
    Version: 2.0
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

#region Global Variables
$script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ConfigPath = Join-Path $script:ScriptPath "config.json"
$script:Config = $null
$script:SoftwareList = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
$script:ComparisonData = $null
$script:ComparisonSnapshots = @()  # Stores version history: @{Date, Data} for each compared scan
$script:Window = $null
#endregion

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
    # Detect Windows system theme setting
    # Returns "dark" or "light"
    try {
        $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $value = Get-ItemPropertyValue -Path $regPath -Name "AppsUseLightTheme" -ErrorAction Stop
        if ($value -eq 0) { return "dark" }
        else { return "light" }
    } catch {
        return "light"  # Default to light if can't detect
    }
}

function Get-CurrentTheme {
    # Get the theme to use based on config setting
    $themeSetting = $script:Config.display.theme
    if (-not $themeSetting) { $themeSetting = "system" }

    switch ($themeSetting.ToLower()) {
        "dark" { return "dark" }
        "light" { return "light" }
        default { return Get-WindowsTheme }  # "system" or any other value
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

    # Find and style controls
    $toolbarBorder = $Window.Content.Children[0]
    $searchBorder = $Window.Content.Children[1]
    $dataGrid = $Window.FindName("dgSoftware")
    $statusBorder = $Window.Content.Children[3]
    $statusText = $Window.FindName("txtStatus")
    $searchText = $Window.FindName("txtSearch")

    # Toolbar
    if ($toolbarBorder) {
        $toolbarBorder.Background = $script:ThemeColors.PanelBackground
        $toolbarBorder.BorderBrush = $script:ThemeColors.BorderColor
    }

    # Search bar
    if ($searchBorder) {
        $searchBorder.Background = $script:ThemeColors.CardBackground
        # Find text elements in search bar and style them
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

        # NEW status trigger
        $newTrigger = New-Object System.Windows.DataTrigger
        $newTrigger.Binding = New-Object System.Windows.Data.Binding("ChangeStatus")
        $newTrigger.Value = "NEW"
        $newTrigger.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.DataGridRow]::BackgroundProperty, $script:ThemeColors.NewRowBackground)))
        $rowStyle.Triggers.Add($newTrigger)

        # REMOVED status trigger
        $removedTrigger = New-Object System.Windows.DataTrigger
        $removedTrigger.Binding = New-Object System.Windows.Data.Binding("ChangeStatus")
        $removedTrigger.Value = "REMOVED"
        $removedTrigger.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.DataGridRow]::BackgroundProperty, $script:ThemeColors.RemovedRowBackground)))
        $rowStyle.Triggers.Add($removedTrigger)

        $dataGrid.RowStyle = $rowStyle

        # Style DataGrid cells for text color
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
                # Primary button keeps accent color
                $button.Background = $script:ThemeColors.AccentColor
                $button.Foreground = $script:ThemeColors.HeaderForeground
            } else {
                # Secondary buttons
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

            # Skip system components if configured
            if (-not $script:Config.display.showSystemComponents -and $item.SystemComponent -eq 1) { continue }

            # Check exclude patterns
            $excluded = $false
            foreach ($pattern in $script:Config.display.excludePatterns) {
                if ($item.DisplayName -match $pattern) {
                    $excluded = $true
                    break
                }
            }
            if ($excluded) { continue }

            # Parse install date
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

            # Format size
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

            # Get registry key name (GUID or name)
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
            # Skip frameworks if configured
            if (-not $script:Config.display.showFrameworks -and $pkg.IsFramework) { continue }

            # Try to get friendly name from manifest
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

            # Check exclude patterns
            $excluded = $false
            foreach ($pattern in $script:Config.display.excludePatterns) {
                if ($displayName -match $pattern -or $pkg.Name -match $pattern) {
                    $excluded = $true
                    break
                }
            }
            if ($excluded) { continue }

            # Parse publisher
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

    # Check if winget exists
    $wingetPath = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $wingetPath) {
        return $software
    }

    try {
        # Get winget list output
        $output = & winget list --disable-interactivity 2>$null

        if (-not $output) { return $software }

        # Find header line
        $headerIndex = -1
        for ($i = 0; $i -lt $output.Count; $i++) {
            if ($output[$i] -match '^Name\s+Id\s+Version') {
                $headerIndex = $i
                break
            }
        }

        if ($headerIndex -lt 0) { return $software }

        # Parse separator line to get column positions
        $separatorLine = $output[$headerIndex + 1]
        if ($separatorLine -notmatch '^-+') {
            $separatorLine = $output[$headerIndex]
        }

        # Find column positions from header
        $headerLine = $output[$headerIndex]
        $nameStart = 0
        $idStart = $headerLine.IndexOf("Id")
        $versionStart = $headerLine.IndexOf("Version")
        $availableStart = $headerLine.IndexOf("Available")
        $sourceStart = $headerLine.IndexOf("Source")

        # Parse data lines
        for ($i = $headerIndex + 2; $i -lt $output.Count; $i++) {
            $line = $output[$i]
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if ($line -match '^-+$') { continue }

            # Extract fields based on positions
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

            # Check exclude patterns
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
    $script:ComparisonSnapshots = @()  # Clear comparison data for fresh scan
    $allSoftware = @()

    # Scan Registry
    if ($script:Config.dataSources.registry.enabled) {
        if ($StatusText) { $StatusText.Text = "Scanning Windows Registry..." }
        if ($ProgressBar) { $ProgressBar.Value = 10 }
        [System.Windows.Forms.Application]::DoEvents()

        $regSoftware = Get-RegistrySoftware
        $allSoftware += $regSoftware
    }

    # Scan AppX
    if ($script:Config.dataSources.appx.enabled) {
        if ($StatusText) { $StatusText.Text = "Scanning AppX packages..." }
        if ($ProgressBar) { $ProgressBar.Value = 40 }
        [System.Windows.Forms.Application]::DoEvents()

        $appxSoftware = Get-AppxSoftware
        $allSoftware += $appxSoftware
    }

    # Scan Winget
    if ($script:Config.dataSources.winget.enabled) {
        if ($StatusText) { $StatusText.Text = "Scanning Winget packages..." }
        if ($ProgressBar) { $ProgressBar.Value = 70 }
        [System.Windows.Forms.Application]::DoEvents()

        $wingetSoftware = Get-WingetSoftware
        $allSoftware += $wingetSoftware
    }

    # Apply custom names
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

    # Sort and add to collection
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
    # Try to extract date from filename pattern like SoftwareInventory_COMPUTERNAME_20240115
    if ($filename -match '_(\d{8})$') {
        try {
            return [datetime]::ParseExact($Matches[1], "yyyyMMdd", $null).ToString("yyyy-MM-dd")
        } catch { }
    }
    # Fallback to file modification date
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

        # Get date for snapshot
        $previousDate = Get-DateFromFilename -FilePath $PreviousFilePath

        # Build lookup from previous data (by display name for matching)
        $previousLookup = @{}
        foreach ($item in $previousData) {
            $displayName = if ($item.CustomName) { $item.CustomName } else { $item.Name }
            $key = "$($item.Source)::$displayName"
            $previousLookup[$key] = $item
        }

        # Store snapshots for HTML comparison export
        $script:ComparisonSnapshots = @(
            @{
                Date = $previousDate
                Data = $previousLookup
            }
        )

        # Compare current with previous
        $currentKeys = @{}
        foreach ($item in $script:SoftwareList) {
            $displayName = if ($item.CustomName) { $item.CustomName } else { $item.Name }
            $key = "$($item.Source)::$displayName"
            $currentKeys[$key] = $true

            if ($previousLookup.ContainsKey($key)) {
                $prev = $previousLookup[$key]
                # Store previous version for comparison
                $item | Add-Member -NotePropertyName "PreviousVersion" -NotePropertyValue $prev.Version -Force
                # Check for changes
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

        # Find removed items and add them with empty current version
        foreach ($key in $previousLookup.Keys) {
            if (-not $currentKeys.ContainsKey($key)) {
                $prev = $previousLookup[$key]
                $removed = [PSCustomObject]@{
                    Name = $prev.Name
                    CustomName = if ($prev.CustomName) { $prev.CustomName } else { "" }
                    Version = ""  # No current version
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

        # Resort with status (changes first, then alphabetical)
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

    # Check if we have comparison data
    $hasComparison = $script:ComparisonSnapshots.Count -gt 0
    $previousDate = if ($hasComparison) { $script:ComparisonSnapshots[0].Date } else { $null }

    # Count changes
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
            # Previous version column
            $prevVersion = $item.PreviousVersion
            if ($prevVersion) {
                $html += "                    <td class=`"version`">$([System.Web.HttpUtility]::HtmlEncode($prevVersion))</td>`n"
            } else {
                $html += "                    <td class=`"version empty-version`">-</td>`n"
            }

            # Current version column
            $currVersion = $item.Version
            if ($currVersion) {
                $html += "                    <td class=`"version`">$([System.Web.HttpUtility]::HtmlEncode($currVersion))</td>`n"
            } else {
                $html += "                    <td class=`"version empty-version`">-</td>`n"
            }

            # Status column
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

    # Define column widths for each property
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

    # Build columns based on enabled properties in config
    foreach ($prop in $script:Config.properties.PSObject.Properties) {
        if ($prop.Value.enabled -eq $true) {
            $propName = $prop.Name
            $displayName = $prop.Value.displayName
            $width = if ($columnWidths[$propName]) { $columnWidths[$propName] } else { 100 }

            $column = New-Object System.Windows.Controls.DataGridTextColumn
            $column.Header = $displayName
            $column.Binding = New-Object System.Windows.Data.Binding($propName)
            $column.Width = $width

            # CustomName column is editable, others are read-only
            if ($propName -eq "CustomName") {
                $column.IsReadOnly = $false
                # Apply italic blue style
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

    # Always add ChangeStatus column at the end (for comparison feature)
    $statusColumn = New-Object System.Windows.Controls.DataGridTextColumn
    $statusColumn.Header = "Status"
    $statusColumn.Binding = New-Object System.Windows.Data.Binding("ChangeStatus")
    $statusColumn.Width = 130
    $statusColumn.IsReadOnly = $true

    # Create style with triggers for status column
    $statusStyle = New-Object System.Windows.Style([System.Windows.Controls.TextBlock])
    $statusStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBlock]::FontWeightProperty, [System.Windows.FontWeights]::Bold)))
    $statusColumn.ElementStyle = $statusStyle

    $DataGrid.Columns.Add($statusColumn)
}
#endregion

#region XAML Definition
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SoftwareLister - Software Inventory Tool" Height="800" Width="1300"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5">
    <Window.Resources>
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
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Toolbar -->
        <Border Grid.Row="0" Background="White" BorderBrush="#e0e0e0" BorderThickness="0,0,0,1" Padding="10">
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
        <Border Grid.Row="1" Background="#fafafa" Padding="10">
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

        <!-- Main DataGrid - Columns generated dynamically from config -->
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
        <Border Grid.Row="3" Background="#333" Padding="10">
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

    # Get controls
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

    # Populate current values
    $chkRegistry.IsChecked = $script:Config.dataSources.registry.enabled
    $chkAppx.IsChecked = $script:Config.dataSources.appx.enabled
    $chkWinget.IsChecked = $script:Config.dataSources.winget.enabled
    $chkShowSystem.IsChecked = $script:Config.display.showSystemComponents
    $chkShowFrameworks.IsChecked = $script:Config.display.showFrameworks

    # Set current theme
    $currentTheme = if ($script:Config.display.theme) { $script:Config.display.theme.ToLower() } else { "system" }
    switch ($currentTheme) {
        "light" { $rbThemeLight.IsChecked = $true }
        "dark" { $rbThemeDark.IsChecked = $true }
        default { $rbThemeSystem.IsChecked = $true }
    }

    # Populate properties checkboxes
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

    # Populate exclude patterns
    $txtExcludePatterns.Text = $script:Config.display.excludePatterns -join "`r`n"

    $btnSaveSettings.Add_Click({
        # Save data sources
        $script:Config.dataSources.registry.enabled = $chkRegistry.IsChecked
        $script:Config.dataSources.appx.enabled = $chkAppx.IsChecked
        $script:Config.dataSources.winget.enabled = $chkWinget.IsChecked

        # Save display options
        $script:Config.display.showSystemComponents = $chkShowSystem.IsChecked
        $script:Config.display.showFrameworks = $chkShowFrameworks.IsChecked

        # Save theme
        if ($rbThemeDark.IsChecked) {
            $script:Config.display.theme = "dark"
        } elseif ($rbThemeLight.IsChecked) {
            $script:Config.display.theme = "light"
        } else {
            $script:Config.display.theme = "system"
        }

        # Save properties
        foreach ($propName in $propertyCheckboxes.Keys) {
            $script:Config.properties.$propName.enabled = $propertyCheckboxes[$propName].IsChecked
        }

        # Save exclude patterns
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

# Add System.Web for HTML encoding
Add-Type -AssemblyName System.Web

# Create window
$reader = New-Object System.Xml.XmlNodeReader $xaml
$script:Window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$btnScan = $script:Window.FindName("btnScan")
$btnCompare = $script:Window.FindName("btnCompare")
$btnExportCsv = $script:Window.FindName("btnExportCsv")
$btnExportHtml = $script:Window.FindName("btnExportHtml")
$btnExportJson = $script:Window.FindName("btnExportJson")
$btnSettings = $script:Window.FindName("btnSettings")
$txtSearch = $script:Window.FindName("txtSearch")
$btnClearSearch = $script:Window.FindName("btnClearSearch")
$txtItemCount = $script:Window.FindName("txtItemCount")
$dgSoftware = $script:Window.FindName("dgSoftware")
$txtStatus = $script:Window.FindName("txtStatus")
$pbProgress = $script:Window.FindName("pbProgress")

# Bind DataGrid and build columns from config
$dgSoftware.ItemsSource = $script:SoftwareList
Build-DataGridColumns -DataGrid $dgSoftware

# Apply theme
Apply-Theme -Window $script:Window

# Event Handlers
$btnScan.Add_Click({
    $btnScan.IsEnabled = $false
    $pbProgress.Value = 0

    try {
        $count = Invoke-SoftwareScan -ProgressBar $pbProgress -StatusText $txtStatus
        $txtItemCount.Text = $count.ToString()

        # Enable export buttons
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
        # Rebuild columns and reapply theme based on updated settings
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

    # Count visible items after filter
    $count = 0
    foreach ($item in $view) { $count++ }
    $txtItemCount.Text = $count.ToString()
})

$btnClearSearch.Add_Click({
    $txtSearch.Text = ""
})

# Handle custom name editing
$dgSoftware.Add_CellEditEnding({
    param($sender, $e)

    if ($e.Column.Header -eq "Custom Name") {
        $item = $e.Row.Item
        $textBox = $e.EditingElement
        $newName = $textBox.Text

        # Save custom name to config
        $key = "$($item.Source)::$($item.UniqueId)"
        if (-not $key -or $key -eq "::") {
            $key = "$($item.Source)::$($item.Name)"
        }

        if ([string]::IsNullOrWhiteSpace($newName)) {
            # Remove custom name
            if ($script:Config.customNames.PSObject.Properties[$key]) {
                $script:Config.customNames.PSObject.Properties.Remove($key)
            }
        } else {
            # Add/update custom name
            if (-not $script:Config.customNames) {
                $script:Config.customNames = [PSCustomObject]@{}
            }
            $script:Config.customNames | Add-Member -NotePropertyName $key -NotePropertyValue $newName -Force
        }

        Save-Configuration
    }
})

# Apply title bar theme when window is loaded (handle available)
$script:Window.Add_Loaded({
    $theme = Get-CurrentTheme
    try {
        $windowHelper = New-Object System.Windows.Interop.WindowInteropHelper($script:Window)
        $hwnd = $windowHelper.Handle
        if ($hwnd -ne [IntPtr]::Zero) {
            [DwmApi]::SetDarkTitleBar($hwnd, ($theme -eq "dark"))
        }
    } catch {
        # Ignore if API not available
    }
})

# Show window
$script:Window.ShowDialog()
#endregion
