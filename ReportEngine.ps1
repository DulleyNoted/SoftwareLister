<#
.SYNOPSIS
    ReportEngine.ps1 — Shared export / import / compare engine for Driver Manager and Service Manager.

.DESCRIPTION
    Dot-source this file from DriverManager.ps1 and ServiceManager.ps1 (and the combined tabbed app):

        . "$PSScriptRoot\ReportEngine.ps1"

    It exposes the following public functions:

        Export-ReportCsv   -Data <list> -DefaultName <string>
        Export-ReportHtml  -Data <list> -Title <string> -DefaultName <string> [-HighlightField <string>] [-HighlightValues <string[]>]
        Import-ReportCsv   -DefaultName <string>   → returns [PSCustomObject[]] or $null
        Compare-Reports    -Baseline <PSCustomObject[]> -Current <PSCustomObject[]> -KeyField <string> -CompareFields <string[]>
        Show-CompareWindow -Results <PSCustomObject[]> -Title <string>

    The Compare window is a self-contained WPF dialog — no extra controls needed in the host window.

.NOTES
    Requires: PresentationFramework already loaded by the calling script.
#>

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT — CSV
# ─────────────────────────────────────────────────────────────────────────────
function Export-ReportCsv {
    param(
        [Parameter(Mandatory)][System.Collections.IEnumerable]$Data,
        [string]$DefaultName = "Report_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    )

    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Title    = "Export report as CSV"
    $dialog.Filter   = "CSV Files (*.csv)|*.csv"
    $dialog.FileName = "$DefaultName.csv"

    if (-not $dialog.ShowDialog()) { return $null }

    try {
        $Data | Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8
        return $dialog.FileName
    }
    catch {
        [System.Windows.MessageBox]::Show("CSV export failed:`n$_", "Export Error",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT — HTML
# Produces a self-contained dark-themed HTML report that opens in any browser.
# HighlightField / HighlightValues: rows where HighlightField matches any value
# in HighlightValues get a warning row color.
# ─────────────────────────────────────────────────────────────────────────────
function Export-ReportHtml {
    param(
        [Parameter(Mandatory)][System.Collections.IEnumerable]$Data,
        [Parameter(Mandatory)][string]$Title,
        [string]$DefaultName   = "Report_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
        [string]$HighlightField  = '',
        [string[]]$HighlightValues = @()
    )

    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Title    = "Export report as HTML"
    $dialog.Filter   = "HTML Files (*.html)|*.html"
    $dialog.FileName = "$DefaultName.html"

    if (-not $dialog.ShowDialog()) { return $null }

    try {
        $rows      = @($Data)
        $columns   = if ($rows.Count -gt 0) { $rows[0].PSObject.Properties.Name } else { @() }
        $generated = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

        # Build table rows
        $rowsHtml = foreach ($row in $rows) {
            $isWarning = $false
            if ($HighlightField -and $HighlightValues.Count -gt 0) {
                $val = $row.$HighlightField
                if ($val -and $HighlightValues -contains $val.ToString()) { $isWarning = $true }
            }
            $rowClass = if ($isWarning) { ' class="warn"' } else { '' }

            $cells = foreach ($col in $columns) {
                $val = $row.$col
                $cellVal = if ($null -eq $val) { '' } else { [System.Web.HttpUtility]::HtmlEncode($val.ToString()) }
                "<td>$cellVal</td>"
            }
            "<tr$rowClass>$($cells -join '')</tr>"
        }

        $headerCells = ($columns | ForEach-Object { "<th>$_</th>" }) -join ''

        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>$([System.Web.HttpUtility]::HtmlEncode($Title))</title>
<style>
  :root {
    --bg:       #1e1e2e;
    --surface:  #181825;
    --surface2: #313244;
    --border:   #45475a;
    --text:     #cdd6f4;
    --muted:    #6c7086;
    --accent:   #89b4fa;
    --green:    #a6e3a1;
    --red:      #f38ba8;
    --yellow:   #f9e2af;
    --warn-bg:  #2d1f1f;
    --warn-fg:  #f38ba8;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: var(--bg); color: var(--text); font-family: 'Segoe UI', system-ui, sans-serif; font-size: 13px; padding: 24px; }
  h1   { font-size: 20px; font-weight: 700; color: var(--accent); margin-bottom: 4px; }
  .meta { color: var(--muted); font-size: 11px; margin-bottom: 20px; }
  .summary { display: flex; gap: 16px; margin-bottom: 20px; flex-wrap: wrap; }
  .chip { background: var(--surface2); border-radius: 6px; padding: 8px 14px; font-size: 12px; }
  .chip span { font-weight: 700; color: var(--accent); }
  input#search { background: var(--surface2); border: none; border-radius: 6px; padding: 8px 14px;
                 color: var(--text); font-size: 13px; width: 320px; margin-bottom: 14px; outline: none; }
  input#search::placeholder { color: var(--muted); }
  .table-wrap { overflow-x: auto; border-radius: 8px; border: 1px solid var(--border); }
  table { width: 100%; border-collapse: collapse; }
  thead th { background: var(--surface2); color: var(--accent); font-size: 11px; font-weight: 600;
             text-transform: uppercase; letter-spacing: .5px; padding: 10px 12px; text-align: left;
             position: sticky; top: 0; cursor: pointer; user-select: none; white-space: nowrap; }
  thead th:hover { background: var(--border); }
  thead th.asc::after  { content: ' ↑'; }
  thead th.desc::after { content: ' ↓'; }
  tbody tr { border-bottom: 1px solid var(--border); }
  tbody tr:last-child { border-bottom: none; }
  tbody tr:nth-child(even) { background: var(--surface); }
  tbody tr:hover { background: var(--surface2); }
  tbody tr.warn  { background: var(--warn-bg); color: var(--warn-fg); }
  tbody tr.warn:hover { background: #3a1f1f; }
  td { padding: 8px 12px; vertical-align: top; word-break: break-word; max-width: 340px; }
  .hidden { display: none; }
  footer { margin-top: 20px; color: var(--muted); font-size: 11px; }
</style>
</head>
<body>
<h1>$([System.Web.HttpUtility]::HtmlEncode($Title))</h1>
<div class="meta">Generated: $generated &nbsp;|&nbsp; $($rows.Count) records</div>

<div class="summary">
  <div class="chip">Total: <span>$($rows.Count)</span></div>
</div>

<input id="search" type="text" placeholder="🔍  Filter rows..." oninput="filterTable(this.value)">

<div class="table-wrap">
<table id="mainTable">
  <thead><tr>$headerCells</tr></thead>
  <tbody>
$($rowsHtml -join "`n")
  </tbody>
</table>
</div>

<footer>Exported by System Manager Suite &nbsp;|&nbsp; $generated</footer>

<script>
function filterTable(q) {
  q = q.toLowerCase();
  document.querySelectorAll('#mainTable tbody tr').forEach(tr => {
    tr.classList.toggle('hidden', q && !tr.textContent.toLowerCase().includes(q));
  });
}

// Column sort
let sortCol = -1, sortAsc = true;
document.querySelectorAll('#mainTable thead th').forEach((th, i) => {
  th.addEventListener('click', () => {
    const tbody = document.querySelector('#mainTable tbody');
    const rows  = Array.from(tbody.querySelectorAll('tr'));
    sortAsc = (sortCol === i) ? !sortAsc : true;
    sortCol = i;
    rows.sort((a, b) => {
      const av = a.cells[i]?.textContent.trim() ?? '';
      const bv = b.cells[i]?.textContent.trim() ?? '';
      return sortAsc ? av.localeCompare(bv, undefined, {numeric:true}) : bv.localeCompare(av, undefined, {numeric:true});
    });
    rows.forEach(r => tbody.appendChild(r));
    document.querySelectorAll('#mainTable thead th').forEach(h => h.classList.remove('asc','desc'));
    th.classList.add(sortAsc ? 'asc' : 'desc');
  });
});
</script>
</body>
</html>
"@
        # Need System.Web for HtmlEncode — load it if not present
        Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

        # Re-render with proper encoding now that assembly is loaded
        $html | Set-Content -Path $dialog.FileName -Encoding UTF8
        return $dialog.FileName
    }
    catch {
        [System.Windows.MessageBox]::Show("HTML export failed:`n$_", "Export Error",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# IMPORT — CSV
# Returns array of PSCustomObjects matching the exported schema, or $null.
# ─────────────────────────────────────────────────────────────────────────────
function Import-ReportCsv {
    param(
        [string]$DefaultName = ''
    )

    $dialog = [Microsoft.Win32.OpenFileDialog]::new()
    $dialog.Title  = "Import previously exported report (CSV)"
    $dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"

    if (-not $dialog.ShowDialog()) { return $null }

    try {
        $data = Import-Csv -Path $dialog.FileName -Encoding UTF8
        if (-not $data -or $data.Count -eq 0) {
            [System.Windows.MessageBox]::Show("The selected file is empty or could not be parsed.",
                "Import Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return $null
        }
        return $data
    }
    catch {
        [System.Windows.MessageBox]::Show("CSV import failed:`n$_", "Import Error",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# COMPARE ENGINE
# Compares a baseline (imported CSV) against current live data.
#
# KeyField     — the property used to match rows between snapshots (e.g. "Name", "InstanceId")
# CompareFields — the properties to diff if a key is found in both snapshots
#
# Returns a list of [PSCustomObject] diff rows with these fields:
#   ChangeType  : Added | Removed | Changed | Unchanged
#   Key         : the key field value
#   Field       : which field changed (empty for Added/Removed)
#   OldValue    : value in baseline
#   NewValue    : value in current
#   Summary     : human-readable one-liner
# ─────────────────────────────────────────────────────────────────────────────
function Compare-Reports {
    param(
        [Parameter(Mandatory)][object[]]$Baseline,
        [Parameter(Mandatory)][object[]]$Current,
        [Parameter(Mandatory)][string]$KeyField,
        [Parameter(Mandatory)][string[]]$CompareFields
    )

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    $baselineMap = @{}
    foreach ($row in $Baseline) {
        $key = $row.$KeyField
        if ($key) { $baselineMap[$key.ToString()] = $row }
    }

    $currentMap = @{}
    foreach ($row in $Current) {
        $key = $row.$KeyField
        if ($key) { $currentMap[$key.ToString()] = $row }
    }

    # Find added (in current, not in baseline)
    foreach ($key in $currentMap.Keys) {
        if (-not $baselineMap.ContainsKey($key)) {
            $cur = $currentMap[$key]
            $results.Add([PSCustomObject]@{
                ChangeType = 'Added'
                Key        = $key
                Field      = ''
                OldValue   = ''
                NewValue   = ($CompareFields | ForEach-Object { "$_=$($cur.$_)" }) -join '; '
                Summary    = "NEW: $key"
            })
        }
    }

    # Find removed (in baseline, not in current)
    foreach ($key in $baselineMap.Keys) {
        if (-not $currentMap.ContainsKey($key)) {
            $base = $baselineMap[$key]
            $results.Add([PSCustomObject]@{
                ChangeType = 'Removed'
                Key        = $key
                Field      = ''
                OldValue   = ($CompareFields | ForEach-Object { "$_=$($base.$_)" }) -join '; '
                NewValue   = ''
                Summary    = "REMOVED: $key"
            })
        }
    }

    # Find changed (in both — diff each CompareField)
    foreach ($key in $baselineMap.Keys) {
        if (-not $currentMap.ContainsKey($key)) { continue }

        $base = $baselineMap[$key]
        $cur  = $currentMap[$key]
        $anyChange = $false

        foreach ($field in $CompareFields) {
            $oldVal = if ($null -eq $base.$field) { '' } else { $base.$field.ToString().Trim() }
            $newVal = if ($null -eq $cur.$field)  { '' } else { $cur.$field.ToString().Trim() }

            if ($oldVal -ne $newVal) {
                $anyChange = $true
                $results.Add([PSCustomObject]@{
                    ChangeType = 'Changed'
                    Key        = $key
                    Field      = $field
                    OldValue   = $oldVal
                    NewValue   = $newVal
                    Summary    = "$key -> ${field}: '$oldVal' -> '$newVal'"
                })
            }
        }

        if (-not $anyChange) {
            $results.Add([PSCustomObject]@{
                ChangeType = 'Unchanged'
                Key        = $key
                Field      = ''
                OldValue   = ''
                NewValue   = ''
                Summary    = "No changes"
            })
        }
    }

    return $results
}

# ─────────────────────────────────────────────────────────────────────────────
# COMPARE WINDOW — self-contained WPF diff viewer dialog
# ─────────────────────────────────────────────────────────────────────────────
function Show-CompareWindow {
    param(
        [Parameter(Mandatory)][System.Collections.Generic.List[PSCustomObject]]$Results,
        [string]$Title = "Report Comparison",
        [string]$BaselineFile = '',
        [string]$CurrentLabel = 'Current (Live)'
    )

    [xml]$cmpXAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Compare Reports"
    Height="680" Width="1050"
    MinHeight="420" MinWidth="700"
    WindowStartupLocation="CenterScreen"
    Background="#1E1E2E">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,12">
            <TextBlock Text="⇄" FontSize="22" Foreground="#CBA6F7" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <StackPanel>
                <TextBlock Text="$([System.Web.HttpUtility]::HtmlEncode($Title))" FontSize="18" FontWeight="Bold" Foreground="#CDD6F4"/>
                <TextBlock x:Name="cmpSubtitle" FontSize="11" Foreground="#6C7086"/>
            </StackPanel>
        </StackPanel>

        <!-- Summary chips -->
        <WrapPanel Grid.Row="1" Margin="0,0,0,10">
            <Border x:Name="chipAdded"     Background="#1a2e1a" CornerRadius="5" Padding="12,5" Margin="0,0,8,0">
                <TextBlock x:Name="lblAdded"   Foreground="#A6E3A1" FontSize="12" FontWeight="SemiBold"/>
            </Border>
            <Border x:Name="chipRemoved"   Background="#2e1a1a" CornerRadius="5" Padding="12,5" Margin="0,0,8,0">
                <TextBlock x:Name="lblRemoved" Foreground="#F38BA8" FontSize="12" FontWeight="SemiBold"/>
            </Border>
            <Border x:Name="chipChanged"   Background="#2e2a1a" CornerRadius="5" Padding="12,5" Margin="0,0,8,0">
                <TextBlock x:Name="lblChanged" Foreground="#F9E2AF" FontSize="12" FontWeight="SemiBold"/>
            </Border>
            <Border x:Name="chipUnchanged" Background="#1a1a2e" CornerRadius="5" Padding="12,5" Margin="0,0,8,0">
                <TextBlock x:Name="lblUnchanged" Foreground="#89B4FA" FontSize="12" FontWeight="SemiBold"/>
            </Border>
        </WrapPanel>

        <!-- Filter bar -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,8">
            <ComboBox x:Name="cboFilter" Width="160" Height="30"
                      Background="#313244" Foreground="#CDD6F4" BorderThickness="0"
                      FontSize="12" Margin="0,0,8,0">
                <ComboBoxItem Content="All Changes"   IsSelected="True"/>
                <ComboBoxItem Content="Added"/>
                <ComboBoxItem Content="Removed"/>
                <ComboBoxItem Content="Changed"/>
                <ComboBoxItem Content="Unchanged"/>
            </ComboBox>
            <Button x:Name="btnCmpExportCsv"  Content="Export CSV"  Height="30"
                    Background="#313244" Foreground="#CDD6F4" BorderThickness="0"
                    Cursor="Hand" Padding="12,0" FontSize="12" Margin="0,0,8,0"/>
            <Button x:Name="btnCmpExportHtml" Content="Export HTML" Height="30"
                    Background="#313244" Foreground="#CDD6F4" BorderThickness="0"
                    Cursor="Hand" Padding="12,0" FontSize="12"/>
        </StackPanel>

        <!-- Diff grid -->
        <DataGrid x:Name="dgCompare" Grid.Row="3"
                  Background="#181825" Foreground="#CDD6F4" BorderThickness="0"
                  RowBackground="#181825" AlternatingRowBackground="#1E1E2E"
                  GridLinesVisibility="None" HeadersVisibility="Column"
                  AutoGenerateColumns="False" IsReadOnly="True"
                  CanUserReorderColumns="True" CanUserResizeColumns="True" CanUserSortColumns="True"
                  HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto"
                  FontSize="12" SelectionMode="Single">
            <DataGrid.ColumnHeaderStyle>
                <Style TargetType="DataGridColumnHeader">
                    <Setter Property="Background"   Value="#313244"/>
                    <Setter Property="Foreground"   Value="#CBA6F7"/>
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
                    </Style.Triggers>
                </Style>
            </DataGrid.RowStyle>
            <DataGrid.CellStyle>
                <Style TargetType="DataGridCell">
                    <Setter Property="BorderThickness" Value="0"/>
                    <Setter Property="Padding"         Value="6,4"/>
                </Style>
            </DataGrid.CellStyle>
            <DataGrid.Columns>
                <!-- ChangeType with color coding -->
                <DataGridTemplateColumn Header="Change" Width="90" SortMemberPath="ChangeType" CanUserSort="True">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <TextBlock Text="{Binding ChangeType}" FontWeight="SemiBold" FontSize="11">
                                <TextBlock.Style>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="Foreground" Value="#89B4FA"/>
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding ChangeType}" Value="Added">
                                                <Setter Property="Foreground" Value="#A6E3A1"/>
                                            </DataTrigger>
                                            <DataTrigger Binding="{Binding ChangeType}" Value="Removed">
                                                <Setter Property="Foreground" Value="#F38BA8"/>
                                            </DataTrigger>
                                            <DataTrigger Binding="{Binding ChangeType}" Value="Changed">
                                                <Setter Property="Foreground" Value="#F9E2AF"/>
                                            </DataTrigger>
                                            <DataTrigger Binding="{Binding ChangeType}" Value="Unchanged">
                                                <Setter Property="Foreground" Value="#6C7086"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </TextBlock.Style>
                            </TextBlock>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>

                <DataGridTextColumn Header="Key / Name" Binding="{Binding Key}"        Width="200"/>
                <DataGridTextColumn Header="Field"      Binding="{Binding Field}"      Width="130"/>
                <DataGridTextColumn Header="Baseline Value" Binding="{Binding OldValue}" Width="200">
                    <DataGridTextColumn.ElementStyle>
                        <Style TargetType="TextBlock">
                            <Setter Property="Foreground" Value="#F38BA8"/>
                            <Setter Property="TextWrapping" Value="Wrap"/>
                        </Style>
                    </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Current Value"  Binding="{Binding NewValue}" Width="200">
                    <DataGridTextColumn.ElementStyle>
                        <Style TargetType="TextBlock">
                            <Setter Property="Foreground" Value="#A6E3A1"/>
                            <Setter Property="TextWrapping" Value="Wrap"/>
                        </Style>
                    </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Summary" Binding="{Binding Summary}" Width="*">
                    <DataGridTextColumn.ElementStyle>
                        <Style TargetType="TextBlock">
                            <Setter Property="TextWrapping" Value="Wrap"/>
                        </Style>
                    </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Status -->
        <TextBlock x:Name="cmpStatus" Grid.Row="4" Foreground="#6C7086"
                   FontSize="11" Margin="0,8,0,0"/>
    </Grid>
</Window>
"@

    $cmpReader = [System.Xml.XmlNodeReader]::new($cmpXAML)
    $cmpWindow = [Windows.Markup.XamlReader]::Load($cmpReader)

    $cmpSubtitle  = $cmpWindow.FindName('cmpSubtitle')
    $lblAdded     = $cmpWindow.FindName('lblAdded')
    $lblRemoved   = $cmpWindow.FindName('lblRemoved')
    $lblChanged   = $cmpWindow.FindName('lblChanged')
    $lblUnchanged = $cmpWindow.FindName('lblUnchanged')
    $cboFilter    = $cmpWindow.FindName('cboFilter')
    $dgCompare    = $cmpWindow.FindName('dgCompare')
    $cmpStatus    = $cmpWindow.FindName('cmpStatus')
    $btnCsvExp    = $cmpWindow.FindName('btnCmpExportCsv')
    $btnHtmlExp   = $cmpWindow.FindName('btnCmpExportHtml')

    # Summary counts
    $added     = @($Results | Where-Object { $_.ChangeType -eq 'Added'     }).Count
    $removed   = @($Results | Where-Object { $_.ChangeType -eq 'Removed'   }).Count
    $changed   = @($Results | Where-Object { $_.ChangeType -eq 'Changed'   }).Count
    $unchanged = @($Results | Where-Object { $_.ChangeType -eq 'Unchanged' }).Count

    $lblAdded.Text     = "✚  Added: $added"
    $lblRemoved.Text   = "✖  Removed: $removed"
    $lblChanged.Text   = "↕  Changed: $changed"
    $lblUnchanged.Text = "=  Unchanged: $unchanged"

    $baseLabel = if ($BaselineFile) { [System.IO.Path]::GetFileName($BaselineFile) } else { 'Baseline' }
    $cmpSubtitle.Text = "Baseline: $baseLabel   →   $CurrentLabel"
    $cmpStatus.Text   = "$($Results.Count) total diff entries"

    # Set initial data
    $dgCompare.ItemsSource = $Results

    # Filter handler
    $cboFilter.Add_SelectionChanged({
        $cboFilterItem = $cboFilter.SelectedItem -as [System.Windows.Controls.ComboBoxItem]
        $sel = if ($cboFilterItem) { $cboFilterItem.Content } else { 'All Changes' }
        $filtered = if ($sel -eq 'All Changes') {
            $Results
        } else {
            [System.Linq.Enumerable]::Where($Results, [Func[object,bool]]{ param($r) $r.ChangeType -eq $sel })
        }
        $dgCompare.ItemsSource = $filtered
        $cmpStatus.Text = "$(@($filtered).Count) entries shown"
    })

    # Export CSV from compare window
    $btnCsvExp.Add_Click({
        $path = Export-ReportCsv -Data $dgCompare.ItemsSource -DefaultName "Compare_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        if ($path) { $cmpStatus.Text = "Exported: $path" }
    })

    # Export HTML from compare window
    $btnHtmlExp.Add_Click({
        $path = Export-ReportHtml `
            -Data            $dgCompare.ItemsSource `
            -Title           "Comparison Report — $Title" `
            -DefaultName     "Compare_$(Get-Date -Format 'yyyyMMdd_HHmmss')" `
            -HighlightField  'ChangeType' `
            -HighlightValues @('Removed','Changed')
        if ($path) {
            $cmpStatus.Text = "Exported: $path"
            Start-Process $path
        }
    })

    $cmpWindow.ShowDialog() | Out-Null
}

# ─────────────────────────────────────────────────────────────────────────────
# CONVENIENCE WRAPPER
# Called from the host script's "Compare" button.
# Handles the full flow: prompt for baseline CSV → run Compare-Reports → show window.
#
#   $currentData   : the live $script:AllServices or $script:AllDevices list
#   $keyField      : field used to match rows (e.g. "Name" for services, "InstanceId" for devices)
#   $compareFields : fields to diff
#   $windowTitle   : title shown in the compare dialog
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-CompareFlow {
    param(
        [Parameter(Mandatory)][System.Collections.IEnumerable]$CurrentData,
        [Parameter(Mandatory)][string]$KeyField,
        [Parameter(Mandatory)][string[]]$CompareFields,
        [string]$WindowTitle = 'Report Comparison'
    )

    $baseline = Import-ReportCsv
    if ($null -eq $baseline) { return }   # user cancelled or file was empty

    $current = @($CurrentData)
    if ($current.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No live data loaded yet. Please wait for the list to finish loading.",
            "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $results = [System.Collections.Generic.List[PSCustomObject]]$(
        Compare-Reports -Baseline $baseline -Current $current `
                        -KeyField $KeyField -CompareFields $CompareFields
    )

    Show-CompareWindow -Results $results -Title $WindowTitle
}
