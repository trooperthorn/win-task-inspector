<#
.SYNOPSIS
  Windows Task Inspector - a WPF viewer for every Scheduled Task on this machine.

.DESCRIPTION
  Lists all tasks with state, last/next run, last result, a plain-English trigger
  summary, and what the action launches: the resolved executable (COM handlers are
  resolved to their DLL via the registry), the file's publisher/description, and any
  URLs or hosts found in the arguments or task XML.

  Enable/Disable buttons act on the selected rows. SYSTEM-owned tasks need an
  elevated session; use "Relaunch as Admin".

.NOTES
  Requires Windows PowerShell 5.1+ (STA). Read-only until you press Enable/Disable.
#>
[CmdletBinding()]
param()

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

$script:IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$script:ExeCache = @{}
$script:UrlRegex  = [regex]'(?i)\b(?:https?|ftp)://[^\s"<>]+'
$script:HostRegex = [regex]'(?i)\b(?:[a-z0-9-]+\.)+(?:com|net|org|io|microsoft|windows|live|office|cloud|us|gov|edu|club)\b'
$script:IgnoreUrl = [regex]'(?i)^https?://(?:schemas\.microsoft\.com|www\.w3\.org)/'

function Convert-IsoDuration {
    param([string]$Iso)
    if (-not $Iso) { return '' }
    if ($Iso -match '^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$') {
        $parts = @()
        if ($Matches[1]) { $parts += "$($Matches[1])d" }
        if ($Matches[2]) { $parts += "$($Matches[2])h" }
        if ($Matches[3]) { $parts += "$($Matches[3])m" }
        if ($Matches[4]) { $parts += "$($Matches[4])s" }
        return ($parts -join ' ')
    }
    return $Iso
}

function Convert-DaysOfWeek {
    param([int]$Mask)
    if ($Mask -eq 127) { return 'every day' }
    $names = 'Sun','Mon','Tue','Wed','Thu','Fri','Sat'
    $out = for ($i = 0; $i -lt 7; $i++) { if ($Mask -band (1 -shl $i)) { $names[$i] } }
    return ($out -join ',')
}

function Get-TriggerSummary {
    param($Task)
    $out = @()
    foreach ($t in @($Task.Triggers)) {
        if ($null -eq $t) { continue }
        $kind = $t.CimClass.CimClassName -replace '^MSFT_Task', '' -replace 'Trigger$', ''
        $desc = switch ($kind) {
            'Daily'  { "Daily every $($t.DaysInterval)d" }
            'Weekly' { "Weekly ($(Convert-DaysOfWeek $t.DaysOfWeek)) every $($t.WeeksInterval)w" }
            'Time'   { 'Once' }
            'Logon'  { if ($t.UserId) { "At logon of $($t.UserId)" } else { 'At any logon' } }
            'Boot'   { 'At startup' }
            'Idle'   { 'When idle' }
            'Event'  { 'On event' }
            'Registration' { 'On registration' }
            'SessionStateChange' { "Session change ($($t.StateChange))" }
            'Monthly' { 'Monthly' }
            'MonthlyDOW' { 'Monthly (day of week)' }
            default { $kind }
        }
        if ($t.StartBoundary) {
            try { $desc += ' @ ' + ([datetime]$t.StartBoundary).ToString('yyyy-MM-dd HH:mm') } catch {}
        }
        if ($t.Repetition -and $t.Repetition.Interval) {
            $desc += ' repeat ' + (Convert-IsoDuration $t.Repetition.Interval)
        }
        if ($t.Delay) { $desc += ' delay ' + (Convert-IsoDuration $t.Delay) }
        if ($t.Enabled -eq $false) { $desc += ' [trigger disabled]' }
        $out += $desc
    }
    if ($Task.Settings.MaintenanceSettings -and $Task.Settings.MaintenanceSettings.Period) {
        $out += 'Automatic maintenance every ' + (Convert-IsoDuration $Task.Settings.MaintenanceSettings.Period)
    }
    if ($out.Count -eq 0) { return 'None (on demand)' }
    return ($out -join '; ')
}

$script:ComCache = @{}
function Resolve-ComHandler {
    # Direct .NET registry reads: the HKLM: provider is ~0.2s per miss under Classes\.
    # Returns Path (DLL/EXE), Name, and Service when the class is hosted by a Windows
    # service via its AppID (no in-process server; that is not a defect).
    param([string]$ClassId)
    if ($script:ComCache.ContainsKey($ClassId)) { return $script:ComCache[$ClassId] }
    $result = $null
    $hklm = [Microsoft.Win32.Registry]::LocalMachine
    $roots = @(
        @{ Hive = $hklm; Sub = 'SOFTWARE\Classes\CLSID\' },
        @{ Hive = $hklm; Sub = 'SOFTWARE\WOW6432Node\Classes\CLSID\' },
        @{ Hive = [Microsoft.Win32.Registry]::CurrentUser; Sub = 'SOFTWARE\Classes\CLSID\' }
    )
    foreach ($r in $roots) {
        $key = $null
        try { $key = $r.Hive.OpenSubKey($r.Sub + $ClassId) } catch {}
        if ($null -eq $key) { continue }
        try {
            $name = [string]$key.GetValue('')
            foreach ($server in 'InprocServer32', 'LocalServer32') {
                $sk = $key.OpenSubKey($server)
                if ($null -eq $sk) { continue }
                $path = $sk.GetValue('')
                $sk.Close()
                if ($path) {
                    $result = [pscustomobject]@{ Path = [Environment]::ExpandEnvironmentVariables([string]$path); Name = $name; Service = '' }
                    break
                }
            }
            if (-not $result) {
                # Service-hosted class: CLSID -> AppID -> LocalService -> service ImagePath
                $appId = [string]$key.GetValue('AppID')
                if ($appId) {
                    $ak = $null
                    try { $ak = $hklm.OpenSubKey('SOFTWARE\Classes\AppID\' + $appId) } catch {}
                    if ($ak) {
                        $svc = [string]$ak.GetValue('LocalService')
                        $ak.Close()
                        if ($svc) {
                            $img = ''
                            $svk = $null
                            try { $svk = $hklm.OpenSubKey('SYSTEM\CurrentControlSet\Services\' + $svc) } catch {}
                            if ($svk) { $img = [string]$svk.GetValue('ImagePath'); $svk.Close() }
                            if ($img -match '^\s*"?([^"]+?\.exe)') { $img = $Matches[1] }
                            $result = [pscustomobject]@{ Path = [Environment]::ExpandEnvironmentVariables($img); Name = $name; Service = $svc }
                        }
                    }
                }
            }
            if (-not $result -and $name) {
                # Registered but with no server we can find; report the name, do not flag
                $result = [pscustomobject]@{ Path = ''; Name = $name; Service = '' }
            }
        } finally { $key.Close() }
        if ($result) { break }
    }
    $script:ComCache[$ClassId] = $result
    return $result
}

function Get-ExeInfo {
    param([string]$Path)
    if (-not $Path) { return $null }
    $clean = [Environment]::ExpandEnvironmentVariables($Path.Trim().Trim('"'))
    if ($script:ExeCache.ContainsKey($clean)) { return $script:ExeCache[$clean] }
    $info = [pscustomobject]@{ Publisher = ''; Description = ''; Signed = ''; Exists = $false; Resolved = '' }
    $resolved = $clean
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        $cmd = Get-Command $clean -ErrorAction SilentlyContinue
        if ($cmd -and $cmd.Source) { $resolved = $cmd.Source }
    }
    if (Test-Path -LiteralPath $resolved -PathType Leaf) {
        $info.Exists = $true
        try {
            $vi = (Get-Item -LiteralPath $resolved).VersionInfo
            $info.Description = $vi.FileDescription
            $info.Publisher = $vi.CompanyName
        } catch {}
        $info.Resolved = $resolved
    }
    $script:ExeCache[$clean] = $info
    return $info
}

function Get-Signer {
    # Lazy: Authenticode + catalog lookup is slow (~0.5s/file), so only on demand.
    param([string]$Path)
    $ei = Get-ExeInfo $Path
    if (-not $ei -or -not $ei.Exists) { return 'file not found' }
    if ($ei.Signed) { return $ei.Signed }
    try {
        $sig = Get-AuthenticodeSignature -LiteralPath $ei.Resolved -ErrorAction Stop
        if ($sig.Status -eq 'Valid' -and $sig.SignerCertificate) {
            $cn = ($sig.SignerCertificate.Subject -split ',' | Where-Object { $_.Trim() -like 'CN=*' } | Select-Object -First 1)
            $ei.Signed = ($cn -replace '^\s*CN=', '').Trim().Trim('"')
        } else { $ei.Signed = $sig.Status.ToString() }
    } catch { $ei.Signed = 'unknown' }
    return $ei.Signed
}

$script:SidCache = @{}
$script:SidRegex = [regex]'S-1-\d+(?:-\d+){1,14}'
function Test-SidResolves {
    param([string]$Sid)
    if ($script:SidCache.ContainsKey($Sid)) { return $script:SidCache[$Sid] }
    $ok = $false
    try {
        $acct = (New-Object System.Security.Principal.SecurityIdentifier($Sid)).Translate([System.Security.Principal.NTAccount])
        $ok = [bool]$acct
    } catch { $ok = $false }
    $script:SidCache[$Sid] = $ok
    return $ok
}

function Get-TaskInventory {
    $rows = New-Object System.Collections.Generic.List[object]
    $tasks = Get-ScheduledTask -ErrorAction SilentlyContinue
    foreach ($task in $tasks) {
        $info = $null
        try { $info = Get-ScheduledTaskInfo -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction Stop } catch {}
        $xml = ''
        try { $xml = Export-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -ErrorAction Stop } catch {}

        $execs = @(); $argList = @(); $targets = @(); $publishers = @(); $problems = @()
        foreach ($a in @($task.Actions)) {
            if ($null -eq $a) { continue }
            $aKind = $a.CimClass.CimClassName
            if ($aKind -eq 'MSFT_TaskExecAction') {
                $execs += $a.Execute
                if ($a.Arguments) { $argList += $a.Arguments }
                $ei = Get-ExeInfo $a.Execute
                if ($ei) {
                    if ($ei.Description) { $publishers += $ei.Description }
                    if ($ei.Publisher) { $publishers += $ei.Publisher }
                    if (-not $ei.Exists) { $publishers += 'NOT FOUND'; $problems += "missing file: $($a.Execute)" }
                }
            } elseif ($aKind -eq 'MSFT_TaskComHandlerAction') {
                $com = Resolve-ComHandler $a.ClassId
                if ($com) {
                    if ($com.Service) { $execs += "[COM svc $($com.Service)] $($com.Path)" }
                    elseif ($com.Path) { $execs += "[COM] $($com.Path)" }
                    else { $execs += "[COM] $($com.Name) $($a.ClassId)" }
                    if ($com.Name) { $publishers += $com.Name }
                    if ($com.Path) {
                        $ei = Get-ExeInfo $com.Path
                        if ($ei -and $ei.Publisher) { $publishers += $ei.Publisher }
                        if ($ei -and -not $ei.Exists) { $problems += "missing COM server: $($com.Path)" }
                    }
                } else { $execs += "[COM] $($a.ClassId)"; $problems += "unregistered COM class $($a.ClassId)" }
                if ($a.Data) { $argList += $a.Data }
            } else {
                $execs += "[$($aKind -replace '^MSFT_Task','')]"
            }
        }
        # Unknown SIDs in the task name, folder, or run-as principal (leftovers from deleted accounts)
        $sidSources = @($task.TaskName, $task.TaskPath, $task.Principal.UserId) | Where-Object { $_ }
        foreach ($src in $sidSources) {
            foreach ($m in $script:SidRegex.Matches($src)) {
                if (-not (Test-SidResolves $m.Value)) { $problems += "unknown SID $($m.Value)" }
            }
        }
        $problems = @($problems | Sort-Object -Unique)
        $argText = ($argList -join ' ')
        foreach ($m in $script:UrlRegex.Matches($argText + ' ' + $xml)) { if (-not $script:IgnoreUrl.IsMatch($m.Value)) { $targets += $m.Value } }
        foreach ($m in $script:HostRegex.Matches($argText)) { $targets += $m.Value }
        $targets = @($targets | Sort-Object -Unique)

        $lastResult = ''
        if ($info) {
            if ($info.LastTaskResult -eq 0) { $lastResult = 'OK' }
            elseif ($info.LastTaskResult -eq 267011) { $lastResult = 'Never ran' }
            else { $lastResult = ('0x{0:X8}' -f $info.LastTaskResult) }
        }
        $lastRun = $null; $nextRun = $null
        if ($info) {
            if ($info.LastRunTime -and $info.LastRunTime.Year -gt 2000) { $lastRun = $info.LastRunTime }
            if ($info.NextRunTime -and $info.NextRunTime.Year -gt 2000) { $nextRun = $info.NextRunTime }
        }

        $kinds = @(foreach ($t in @($task.Triggers)) { if ($t) { $t.CimClass.CimClassName -replace '^MSFT_Task', '' -replace 'Trigger$', '' } })
        $rows.Add([pscustomobject]@{
            TaskObject  = $task
            TriggerKinds = $kinds
            Folder      = $task.TaskPath
            Name        = $task.TaskName
            State       = $task.State.ToString()
            Author      = $task.Author
            RunAs       = $task.Principal.UserId
            LastRun     = $lastRun
            LastResult  = $lastResult
            NextRun     = $nextRun
            Triggers    = Get-TriggerSummary $task
            Action      = ($execs -join ' | ')
            Arguments   = $argText
            Publisher   = (@($publishers | Sort-Object -Unique) -join '; ')
            Accesses    = ($targets -join ', ')
            Description = $task.Description
            Problem     = ($problems -join '; ')
            HasProblem  = ($problems.Count -gt 0)
            ExePaths    = @($execs | ForEach-Object { $_ -replace '^\[COM[^\]]*\] ', '' } | Where-Object { $_ -and $_ -notmatch '^[\[{]' })
            Xml         = $xml
        })
    }
    return $rows
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Task Inspector" Height="820" Width="1500" WindowStartupLocation="CenterScreen">
  <DockPanel>
    <ToolBar DockPanel.Dock="Top">
      <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="4,0"/>
      <TextBox x:Name="Filter" Width="260" Margin="0,0,8,0"/>
      <CheckBox x:Name="HideMs" Content="Hide \Microsoft\ tasks" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <CheckBox x:Name="OnlyNet" Content="Only with URLs/hosts" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <Separator/>
      <Button x:Name="Refresh" Content="Refresh" Margin="4,0"/>
      <Button x:Name="Disable" Content="Disable selected" Margin="4,0"/>
      <Button x:Name="Enable" Content="Enable selected" Margin="4,0"/>
      <Button x:Name="RunNow" Content="Run now" Margin="4,0"/>
      <Separator/>
      <Button x:Name="ExportCsv" Content="Export CSV" Margin="4,0"/>
      <Button x:Name="Elevate" Content="Relaunch as Admin" Margin="4,0"/>
      <TextBlock x:Name="Status" VerticalAlignment="Center" Margin="12,0" Foreground="Gray"/>
    </ToolBar>
    <Grid>
     <Grid.ColumnDefinitions>
       <ColumnDefinition Width="260" MinWidth="120"/>
       <ColumnDefinition Width="6"/>
       <ColumnDefinition Width="*"/>
     </Grid.ColumnDefinitions>
    <Grid Grid.Column="0" Margin="0,0,2,0">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <TextBlock Grid.Row="0" Text="Views" FontWeight="Bold" Margin="6,4,0,2"/>
      <ListBox x:Name="Nav" Grid.Row="1" FontSize="13" SelectedIndex="0" BorderThickness="0">
        <ListBoxItem Content="All tasks" Tag="all"/>
        <ListBoxItem Content="Enabled" Tag="enabled"/>
        <ListBoxItem Content="Disabled" Tag="disabled"/>
        <ListBoxItem Content="Running now" Tag="running"/>
        <Separator/>
        <ListBoxItem Content="Run at logon" Tag="Logon"/>
        <ListBoxItem Content="Run at startup" Tag="Boot"/>
        <ListBoxItem Content="Run when idle" Tag="Idle"/>
        <ListBoxItem Content="Run on a schedule" Tag="schedule"/>
        <ListBoxItem Content="Run on an event" Tag="Event"/>
        <ListBoxItem Content="No trigger (on demand)" Tag="none"/>
        <Separator/>
        <ListBoxItem Content="Problems" Tag="problems" Foreground="#C62828"/>
        <ListBoxItem Content="Non-Microsoft" Tag="nonms"/>
      </ListBox>
      <TextBlock Grid.Row="2" Text="Task Scheduler Library" FontWeight="Bold" Margin="6,10,0,2"/>
      <TreeView x:Name="Tree" Grid.Row="3" FontSize="13" BorderThickness="0"/>
    </Grid>
    <GridSplitter Grid.Column="1" Width="6" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="#DDD"/>
    <Grid Grid.Column="2">
      <Grid.RowDefinitions>
        <RowDefinition Height="3*"/>
        <RowDefinition Height="5"/>
        <RowDefinition Height="1*"/>
      </Grid.RowDefinitions>
      <DataGrid x:Name="Grid" Grid.Row="0" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Extended"
                CanUserSortColumns="True" GridLinesVisibility="Horizontal" HeadersVisibility="Column" FontSize="12">
        <DataGrid.ContextMenu>
          <ContextMenu>
            <MenuItem x:Name="CmEnable" Header="Enable"/>
            <MenuItem x:Name="CmDisable" Header="Disable"/>
            <MenuItem x:Name="CmRun" Header="Run now"/>
            <MenuItem x:Name="CmStop" Header="Stop"/>
            <Separator/>
            <MenuItem x:Name="CmBrowse" Header="Open file location"/>
            <MenuItem x:Name="CmCopyPath" Header="Copy task path"/>
            <MenuItem x:Name="CmCopyCmd" Header="Copy command line"/>
            <Separator/>
            <MenuItem x:Name="CmDelete" Header="Delete task..." Foreground="DarkRed"/>
            <Separator/>
            <MenuItem x:Name="CmProps" Header="Properties" FontWeight="Bold"/>
          </ContextMenu>
        </DataGrid.ContextMenu>
        <DataGrid.RowStyle>
          <Style TargetType="DataGridRow">
            <Style.Triggers>
              <DataTrigger Binding="{Binding State}" Value="Disabled"><Setter Property="Foreground" Value="Gray"/></DataTrigger>
              <DataTrigger Binding="{Binding State}" Value="Running"><Setter Property="Foreground" Value="DarkGreen"/></DataTrigger>
              <DataTrigger Binding="{Binding HasProblem}" Value="True">
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="Background" Value="#C62828"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </DataTrigger>
            </Style.Triggers>
          </Style>
        </DataGrid.RowStyle>
        <DataGrid.Columns>
          <DataGridTextColumn Header="Folder" Binding="{Binding Folder}" Width="200"/>
          <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="220"/>
          <DataGridTextColumn Header="State" Binding="{Binding State}" Width="70"/>
          <DataGridTextColumn Header="Problem" Binding="{Binding Problem}" Width="220"/>
          <DataGridTextColumn Header="Last run" Binding="{Binding LastRun, StringFormat=yyyy-MM-dd HH:mm}" Width="120"/>
          <DataGridTextColumn Header="Result" Binding="{Binding LastResult}" Width="80"/>
          <DataGridTextColumn Header="Next run" Binding="{Binding NextRun, StringFormat=yyyy-MM-dd HH:mm}" Width="120"/>
          <DataGridTextColumn Header="Triggers" Binding="{Binding Triggers}" Width="260"/>
          <DataGridTextColumn Header="Action" Binding="{Binding Action}" Width="300"/>
          <DataGridTextColumn Header="Publisher / description" Binding="{Binding Publisher}" Width="240"/>
          <DataGridTextColumn Header="Accesses (URLs/hosts)" Binding="{Binding Accesses}" Width="220"/>
          <DataGridTextColumn Header="Run as" Binding="{Binding RunAs}" Width="120"/>
          <DataGridTextColumn Header="Author" Binding="{Binding Author}" Width="160"/>
        </DataGrid.Columns>
      </DataGrid>
      <GridSplitter Grid.Row="1" HorizontalAlignment="Stretch" Height="5"/>
      <TabControl Grid.Row="2">
        <TabItem Header="Details">
          <TextBox x:Name="Details" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12"/>
        </TabItem>
        <TabItem Header="Task XML">
          <TextBox x:Name="XmlBox" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="11"/>
        </TabItem>
      </TabControl>
    </Grid>
    </Grid>
  </DockPanel>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
$ui = @{}
foreach ($n in 'Filter','HideMs','OnlyNet','Nav','Tree','Refresh','Disable','Enable','RunNow','ExportCsv','Elevate','Status','Grid','Details','XmlBox') { $ui[$n] = $window.FindName($n) }
$cm = $ui.Grid.ContextMenu
foreach ($item in $cm.Items) { if ($item -is [System.Windows.Controls.MenuItem] -and $item.Name) { $ui[$item.Name] = $item } }
$ui.HideMs.IsChecked = $true
if ($script:IsAdmin) { $window.Title = 'Task Inspector (Administrator)'; $ui.Elevate.IsEnabled = $false }

$script:All = @()
$script:FolderFilter = '\'

function Build-FolderTree {
    # One node per folder that exists in the task list (Task Scheduler Library layout);
    # header shows the number of tasks directly in that folder.
    # Direct counts per folder, then cumulative counts (folder + every subfolder)
    $counts = @{}
    foreach ($r in $script:All) { $counts[$r.Folder] = 1 + $(if ($counts.ContainsKey($r.Folder)) { $counts[$r.Folder] } else { 0 }) }
    $totals = @{}
    foreach ($folder in $counts.Keys) {
        $parts = $folder.Trim('\') -split '\\' | Where-Object { $_ }
        $path = '\'
        foreach ($part in $parts) {
            $path = $path + $part + '\'
            $totals[$path] = $counts[$folder] + $(if ($totals.ContainsKey($path)) { $totals[$path] } else { 0 })
        }
    }
    $nodes = @{}
    $root = New-Object System.Windows.Controls.TreeViewItem
    $root.Header = "Task Scheduler Library ($($script:All.Count))"
    $root.Tag = '\'
    $root.IsExpanded = $true
    $nodes['\'] = $root
    foreach ($folder in ($counts.Keys | Sort-Object)) {
        $parts = $folder.Trim('\') -split '\\' | Where-Object { $_ }
        $path = '\'
        foreach ($part in $parts) {
            $parentNode = $nodes[$path]
            $path = $path + $part + '\'
            if (-not $nodes.ContainsKey($path)) {
                $n = New-Object System.Windows.Controls.TreeViewItem
                $c = 0; if ($totals.ContainsKey($path)) { $c = $totals[$path] }
                $n.Header = "$part ($c)"
                $n.Tag = $path
                $n.IsExpanded = ($path -eq '\Microsoft\')
                $parentNode.Items.Add($n) | Out-Null
                $nodes[$path] = $n
            }
        }
    }
    $ui.Tree.Items.Clear()
    $ui.Tree.Items.Add($root) | Out-Null
    # Re-select the folder that was active before a refresh
    if ($nodes.ContainsKey($script:FolderFilter)) { $nodes[$script:FolderFilter].IsSelected = $true } else { $script:FolderFilter = '\'; $root.IsSelected = $true }
}

function Apply-Filter {
    $f = $ui.Filter.Text
    $hideMs = $ui.HideMs.IsChecked
    $onlyNet = $ui.OnlyNet.IsChecked
    $nav = 'all'
    if ($ui.Nav.SelectedItem -and $ui.Nav.SelectedItem.Tag) { $nav = [string]$ui.Nav.SelectedItem.Tag }
    $filtered = $script:All | Where-Object {
        $row = $_
        ($row.Folder -like ($script:FolderFilter + '*')) -and
        ($(if ($hideMs -and $nav -ne 'nonms' -and $script:FolderFilter -eq '\') { $row.Folder -notlike '\Microsoft\*' } else { $true })) -and
        ($(if ($onlyNet) { [bool]$row.Accesses } else { $true })) -and
        ($(switch ($nav) {
            'all'      { $true }
            'enabled'  { $row.State -ne 'Disabled' }
            'disabled' { $row.State -eq 'Disabled' }
            'running'  { $row.State -eq 'Running' }
            'schedule' { @($row.TriggerKinds | Where-Object { $_ -in 'Daily','Weekly','Monthly','MonthlyDOW','Time' }).Count -gt 0 }
            'none'     { @($row.TriggerKinds).Count -eq 0 }
            'problems' { $row.HasProblem }
            'nonms'    { $row.Folder -notlike '\Microsoft\*' }
            default    { $nav -in @($row.TriggerKinds) }
        })) -and
        ($(if ($f) { ($_.Folder + $_.Name + $_.Action + $_.Arguments + $_.Publisher + $_.Accesses + $_.Triggers + $_.Author) -like "*$f*" } else { $true }))
    }
    # Remember sort and selection so an action does not reshuffle the grid
    $sorts = @()
    if ($ui.Grid.ItemsSource) {
        $oldView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($ui.Grid.ItemsSource)
        if ($oldView) { $sorts = @($oldView.SortDescriptions | ForEach-Object { New-Object System.ComponentModel.SortDescription($_.PropertyName, $_.Direction) }) }
    }
    $selectedKeys = @($ui.Grid.SelectedItems | ForEach-Object { $_.Folder + $_.Name })

    $view = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
    foreach ($r in @($filtered)) { $view.Add($r) }
    $ui.Grid.ItemsSource = $view

    $newView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($view)
    $newView.SortDescriptions.Clear()
    foreach ($sd in $sorts) {
        $newView.SortDescriptions.Add($sd)
        foreach ($col in $ui.Grid.Columns) { if ($col.SortMemberPath -eq $sd.PropertyName) { $col.SortDirection = $sd.Direction } }
    }
    if ($selectedKeys.Count) {
        $ui.Grid.SelectedItems.Clear()
        foreach ($r in $view) { if (($r.Folder + $r.Name) -in $selectedKeys) { $ui.Grid.SelectedItems.Add($r) } }
        if ($ui.Grid.SelectedItem) { $ui.Grid.ScrollIntoView($ui.Grid.SelectedItem) }
    }
    $note = ''
    if (-not $script:IsAdmin) { $note = '  (not elevated: SYSTEM tasks are read-only)' }
    $probCount = @($script:All | Where-Object HasProblem).Count
    $where = ''; if ($script:FolderFilter -ne '\') { $where = " in $($script:FolderFilter)" }
    $ui.Status.Text = "$($view.Count) of $($script:All.Count) tasks$where, $probCount with problems$note"
}

function Load-Tasks {
    $ui.Status.Text = 'Loading...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
    $script:All = @(Get-TaskInventory | Sort-Object Folder, Name)
    $window.Cursor = $null
    $script:SuppressTree = $true
    Build-FolderTree
    $script:SuppressTree = $false
    Apply-Filter
}

function Set-SelectedState {
    param([bool]$Enable)
    $rows = @($ui.Grid.SelectedItems)
    if ($rows.Count -eq 0) { $ui.Status.Text = 'Nothing selected'; return }
    $verb = 'Disable'; if ($Enable) { $verb = 'Enable' }
    $ok = 0; $fail = @()
    foreach ($r in $rows) {
        try {
            if ($Enable) { Enable-ScheduledTask -TaskPath $r.Folder -TaskName $r.Name -ErrorAction Stop | Out-Null }
            else { Disable-ScheduledTask -TaskPath $r.Folder -TaskName $r.Name -ErrorAction Stop | Out-Null }
            $ok++
        } catch { $fail += "$($r.Name): $($_.Exception.Message.Trim())" }
    }
    Load-Tasks
    $msg = "${verb}: $ok succeeded"
    if ($fail.Count) {
        $msg += ", $($fail.Count) failed"
        [System.Windows.MessageBox]::Show(($fail -join "`n"), "$verb failed", 'OK', 'Warning') | Out-Null
    }
    $ui.Status.Text = $msg
}

$ui.Refresh.Add_Click({ Load-Tasks })
$ui.Filter.Add_TextChanged({ Apply-Filter })
$ui.HideMs.Add_Click({ Apply-Filter })
$ui.OnlyNet.Add_Click({ Apply-Filter })
$ui.Tree.Add_SelectedItemChanged({
    if ($script:SuppressTree) { return }
    $n = $ui.Tree.SelectedItem
    if ($n -and $n.Tag) { $script:FolderFilter = [string]$n.Tag; Apply-Filter }
})
$ui.Nav.Add_SelectionChanged({ if ($ui.Nav.SelectedItem -is [System.Windows.Controls.ListBoxItem]) { Apply-Filter } })
$ui.Disable.Add_Click({ Set-SelectedState -Enable:$false })
$ui.Enable.Add_Click({ Set-SelectedState -Enable:$true })
$ui.RunNow.Add_Click({
    foreach ($r in @($ui.Grid.SelectedItems)) {
        try { Start-ScheduledTask -TaskPath $r.Folder -TaskName $r.Name -ErrorAction Stop }
        catch { [System.Windows.MessageBox]::Show($_.Exception.Message, 'Run failed') | Out-Null }
    }
    Load-Tasks
})
$ui.ExportCsv.Add_Click({
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV|*.csv'
    $dlg.FileName = "scheduled-tasks-$(Get-Date -Format yyyyMMdd-HHmm).csv"
    if ($dlg.ShowDialog()) {
        $script:All | Select-Object Folder,Name,State,Problem,LastRun,LastResult,NextRun,Triggers,Action,Arguments,Publisher,Accesses,RunAs,Author,Description |
            Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        $ui.Status.Text = "Exported to $($dlg.FileName)"
    }
})
$ui.Elevate.Add_Click({
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`""
    $window.Close()
})
# Right-click on a row selects it (keeps an existing multi-selection if the row is already in it)
$ui.Grid.Add_PreviewMouseRightButtonDown({
    param($sender, $e)
    $dep = $e.OriginalSource
    while ($dep -and -not ($dep -is [System.Windows.Controls.DataGridRow])) { $dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep) }
    if ($dep -and -not $dep.IsSelected) { $ui.Grid.SelectedItems.Clear(); $dep.IsSelected = $true }
})

function Invoke-OnSelected {
    param([scriptblock]$Action, [string]$Verb)
    $rows = @($ui.Grid.SelectedItems)
    if ($rows.Count -eq 0) { $ui.Status.Text = 'Nothing selected'; return }
    $ok = 0; $fail = @()
    foreach ($r in $rows) {
        try { & $Action $r; $ok++ } catch { $fail += "$($r.Name): $($_.Exception.Message.Trim())" }
    }
    Load-Tasks
    $msg = "${Verb}: $ok succeeded"
    if ($fail.Count) {
        $msg += ", $($fail.Count) failed"
        [System.Windows.MessageBox]::Show(($fail -join "`n"), "$Verb failed", 'OK', 'Warning') | Out-Null
    }
    $ui.Status.Text = $msg
}

function Remove-SelectedTasks {
    $rows = @($ui.Grid.SelectedItems)
    if ($rows.Count -eq 0) { $ui.Status.Text = 'Nothing selected'; return }
    $names = ($rows | ForEach-Object { "$($_.Folder)$($_.Name)" }) -join "`n"
    $answer = [System.Windows.MessageBox]::Show("Permanently delete $($rows.Count) task(s)?`n`n$names`n`nThis cannot be undone. Export CSV first if you want a record.", 'Delete scheduled task', 'YesNo', 'Warning')
    if ($answer -ne 'Yes') { return }
    Invoke-OnSelected -Verb 'Delete' -Action { param($r) Unregister-ScheduledTask -TaskPath $r.Folder -TaskName $r.Name -Confirm:$false -ErrorAction Stop }
}

function Open-SelectedLocation {
    $r = $ui.Grid.SelectedItem
    if ($null -eq $r) { return }
    $opened = $false
    foreach ($p in @($r.ExePaths)) {
        $ei = Get-ExeInfo $p
        if ($ei -and $ei.Exists) { Start-Process explorer.exe -ArgumentList "/select,`"$($ei.Resolved)`""; $opened = $true; break }
    }
    if (-not $opened) {
        # Fall back to the task definition file itself
        $def = Join-Path $env:SystemRoot ('System32\Tasks' + $r.Folder + $r.Name)
        if (Test-Path -LiteralPath $def) { Start-Process explorer.exe -ArgumentList "/select,`"$def`"" }
        else { [System.Windows.MessageBox]::Show("No file found for this task's action.", 'Open file location') | Out-Null }
    }
}


function Get-PropertyLines {
    param($Obj, [string[]]$Skip = @('CimClass','CimInstanceProperties','CimSystemProperties','PSComputerName'))
    if ($null -eq $Obj) { return @() }
    $out = @()
    foreach ($p in $Obj.PSObject.Properties) {
        if ($p.Name -in $Skip) { continue }
        $v = $p.Value
        if ($null -eq $v -or $v -eq '') { continue }
        if ($v -is [ciminstance]) {
            $inner = (Get-PropertyLines $v) -join ', '
            if ($inner) { $out += ('{0}={1}' -f $p.Name, $inner) }
        } elseif ($v -is [array]) {
            $out += ('{0}={1}' -f $p.Name, ($v -join ', '))
        } else {
            $out += ('{0}={1}' -f $p.Name, $v)
        }
    }
    return $out
}

function Get-TaskHistory {
    param([string]$TaskPath)
    try {
        $log = Get-WinEvent -ListLog 'Microsoft-Windows-TaskScheduler/Operational' -ErrorAction Stop
        if (-not $log.IsEnabled) {
            return @([pscustomobject]@{ Time = $null; Id = ''; Level = ''; Message = 'Task history is disabled. Enable it in Task Scheduler (Actions > Enable All Tasks History).' })
        }
        $xpath = "*[EventData[Data[@Name='TaskName']='$TaskPath']]"
        $events = Get-WinEvent -LogName 'Microsoft-Windows-TaskScheduler/Operational' -FilterXPath $xpath -MaxEvents 200 -ErrorAction Stop
        return @($events | ForEach-Object { [pscustomobject]@{ Time = $_.TimeCreated; Id = $_.Id; Level = $_.LevelDisplayName; Message = ($_.Message -split "`r?`n")[0] } })
    } catch {
        return @([pscustomobject]@{ Time = $null; Id = ''; Level = ''; Message = "No history: $($_.Exception.Message)" })
    }
}

function Show-TaskProperties {
    param($Row)
    if ($null -eq $Row) { return }
    $task = $Row.TaskObject
    [xml]$px = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Properties" Height="640" Width="880" WindowStartupLocation="CenterOwner">
  <DockPanel>
    <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="8">
      <Button x:Name="PClose" Content="Close" Width="80" IsCancel="True"/>
    </StackPanel>
    <TabControl Margin="8">
      <TabItem Header="General"><TextBox x:Name="PGeneral" IsReadOnly="True" FontFamily="Consolas" FontSize="12" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"/></TabItem>
      <TabItem Header="Triggers"><DataGrid x:Name="PTriggers" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column">
        <DataGrid.Columns>
          <DataGridTextColumn Header="Trigger" Binding="{Binding Type}" Width="130"/>
          <DataGridTextColumn Header="Details" Binding="{Binding Details}" Width="*"/>
          <DataGridTextColumn Header="Enabled" Binding="{Binding Enabled}" Width="70"/>
        </DataGrid.Columns></DataGrid></TabItem>
      <TabItem Header="Actions"><DataGrid x:Name="PActions" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column">
        <DataGrid.Columns>
          <DataGridTextColumn Header="Action" Binding="{Binding Type}" Width="130"/>
          <DataGridTextColumn Header="Details" Binding="{Binding Details}" Width="*"/>
        </DataGrid.Columns></DataGrid></TabItem>
      <TabItem Header="Conditions"><TextBox x:Name="PConditions" IsReadOnly="True" FontFamily="Consolas" FontSize="12" VerticalScrollBarVisibility="Auto"/></TabItem>
      <TabItem Header="Settings"><TextBox x:Name="PSettings" IsReadOnly="True" FontFamily="Consolas" FontSize="12" VerticalScrollBarVisibility="Auto"/></TabItem>
      <TabItem Header="History"><DataGrid x:Name="PHistory" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column">
        <DataGrid.Columns>
          <DataGridTextColumn Header="Time" Binding="{Binding Time, StringFormat=yyyy-MM-dd HH:mm:ss}" Width="150"/>
          <DataGridTextColumn Header="Event" Binding="{Binding Id}" Width="60"/>
          <DataGridTextColumn Header="Level" Binding="{Binding Level}" Width="90"/>
          <DataGridTextColumn Header="Message" Binding="{Binding Message}" Width="*"/>
        </DataGrid.Columns></DataGrid></TabItem>
      <TabItem Header="XML"><TextBox x:Name="PXml" IsReadOnly="True" FontFamily="Consolas" FontSize="11" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/></TabItem>
    </TabControl>
  </DockPanel>
</Window>
"@
    $pw = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $px))
    $pw.Owner = $window
    $pw.Title = "$($Row.Name) Properties"
    $g = @{}; foreach ($n in 'PClose','PGeneral','PTriggers','PActions','PConditions','PSettings','PHistory','PXml') { $g[$n] = $pw.FindName($n) }

    $pr = $task.Principal
    $st = $task.Settings
    $g.PGeneral.Text = @(
        "Name:          $($task.TaskName)"
        "Location:      $($task.TaskPath)"
        "Author:        $($task.Author)"
        "Description:   $($task.Description)"
        ""
        "State:         $($task.State)"
        "Last run:      $($Row.LastRun)  ($($Row.LastResult))"
        "Next run:      $($Row.NextRun)"
        ""
        "Security options"
        "  Run as:      $($pr.UserId)  (group: $($pr.GroupId))"
        "  Logon type:  $($pr.LogonType)"
        "  Run level:   $($pr.RunLevel)"
        "  Hidden:      $($st.Hidden)"
        "  Version:     $($task.Version)   Source: $($task.Source)   URI: $($task.URI)"
        ""
        "Problem:       $($Row.Problem)"
    ) -join "`n"

    $g.PTriggers.ItemsSource = @(foreach ($t in @($task.Triggers)) { if ($t) {
        [pscustomobject]@{
            Type    = ($t.CimClass.CimClassName -replace '^MSFT_Task', '' -replace 'Trigger$', '')
            Details = ((Get-PropertyLines $t -Skip @('CimClass','CimInstanceProperties','CimSystemProperties','PSComputerName','Enabled','Id')) -join ' | ')
            Enabled = $t.Enabled
        } } })
    $g.PActions.ItemsSource = @(foreach ($a in @($task.Actions)) { if ($a) {
        [pscustomobject]@{
            Type    = ($a.CimClass.CimClassName -replace '^MSFT_Task', '' -replace 'Action$', '')
            Details = ((Get-PropertyLines $a -Skip @('CimClass','CimInstanceProperties','CimSystemProperties','PSComputerName','Id')) -join ' | ')
        } } })

    $g.PConditions.Text = @(
        "Idle"
        "  Run only if idle:              $($st.RunOnlyIfIdle)"
        "  Idle duration / wait timeout:  $(Convert-IsoDuration $st.IdleSettings.IdleDuration) / $(Convert-IsoDuration $st.IdleSettings.WaitTimeout)"
        "  Stop if idle ends:             $($st.IdleSettings.StopOnIdleEnd)"
        "  Restart on idle:               $($st.IdleSettings.RestartOnIdle)"
        ""
        "Power"
        "  Start only on AC power:        $($st.DisallowStartIfOnBatteries)"
        "  Stop if switching to battery:  $($st.StopIfGoingOnBatteries)"
        "  Wake the computer to run:      $($st.WakeToRun)"
        ""
        "Network"
        "  Only if network available:     $($st.RunOnlyIfNetworkAvailable)"
        "  Network id / name:             $($st.NetworkSettings.Id) $($st.NetworkSettings.Name)"
    ) -join "`n"

    $maint = 'none'
    if ($st.MaintenanceSettings) { $maint = 'period ' + (Convert-IsoDuration $st.MaintenanceSettings.Period) + ', deadline ' + (Convert-IsoDuration $st.MaintenanceSettings.Deadline) + ', exclusive ' + $st.MaintenanceSettings.Exclusive }
    $g.PSettings.Text = @(
        "Allow task to be run on demand:             $($st.AllowDemandStart)"
        "Run as soon as possible after missed start: $($st.StartWhenAvailable)"
        "If the task fails, restart every:           $(Convert-IsoDuration $st.RestartInterval)  (up to $($st.RestartCount) times)"
        "Stop the task if it runs longer than:       $(Convert-IsoDuration $st.ExecutionTimeLimit)"
        "Force stop if it does not end:              $($st.AllowHardTerminate)"
        "Delete if not scheduled to run again after: $(Convert-IsoDuration $st.DeleteExpiredTaskAfter)"
        "If already running:                         $($st.MultipleInstances)"
        "Priority:                                   $($st.Priority)"
        "Compatibility:                              $($st.Compatibility)"
        "Enabled:                                    $($st.Enabled)"
        "Run only when user is logged on:            $($pr.LogonType -eq 'Interactive')"
        "Do not store password:                      $($pr.LogonType -eq 'S4U')"
        "Disallow start on remote app session:       $($st.DisallowStartOnRemoteAppSession)"
        "Use unified scheduling engine:              $($st.UseUnifiedSchedulingEngine)"
        "Volatile:                                   $($st.Volatile)"
        "Maintenance:                                $maint"
    ) -join "`n"

    $g.PHistory.ItemsSource = Get-TaskHistory ($task.TaskPath + $task.TaskName)
    $g.PXml.Text = $Row.Xml
    $g.PClose.Add_Click({ $pw.Close() })
    $pw.ShowDialog() | Out-Null
}

$ui.CmProps.Add_Click({ Show-TaskProperties $ui.Grid.SelectedItem })
$ui.Grid.Add_MouseDoubleClick({ Show-TaskProperties $ui.Grid.SelectedItem })
$ui.CmEnable.Add_Click({ Set-SelectedState -Enable:$true })
$ui.CmDisable.Add_Click({ Set-SelectedState -Enable:$false })
$ui.CmRun.Add_Click({ Invoke-OnSelected -Verb 'Run' -Action { param($r) Start-ScheduledTask -TaskPath $r.Folder -TaskName $r.Name -ErrorAction Stop } })
$ui.CmStop.Add_Click({ Invoke-OnSelected -Verb 'Stop' -Action { param($r) Stop-ScheduledTask -TaskPath $r.Folder -TaskName $r.Name -ErrorAction Stop } })
$ui.CmBrowse.Add_Click({ Open-SelectedLocation })
$ui.CmCopyPath.Add_Click({ $r = $ui.Grid.SelectedItem; if ($r) { [System.Windows.Clipboard]::SetText("$($r.Folder)$($r.Name)"); $ui.Status.Text = 'Task path copied' } })
$ui.CmCopyCmd.Add_Click({ $r = $ui.Grid.SelectedItem; if ($r) { [System.Windows.Clipboard]::SetText(($r.Action + ' ' + $r.Arguments).Trim()); $ui.Status.Text = 'Command line copied' } })
$ui.CmDelete.Add_Click({ Remove-SelectedTasks })

$ui.Grid.Add_SelectionChanged({
    $r = $ui.Grid.SelectedItem
    if ($null -eq $r) { $ui.Details.Text = ''; $ui.XmlBox.Text = ''; return }
    $ui.Details.Text = @(
        "Task:        $($r.Folder)$($r.Name)"
        "State:       $($r.State)    Run as: $($r.RunAs)    Author: $($r.Author)"
        "Problem:     $($r.Problem)"
        "Last run:    $($r.LastRun)  ($($r.LastResult))"
        "Next run:    $($r.NextRun)"
        "Triggers:    $($r.Triggers)"
        "Action:      $($r.Action)"
        "Arguments:   $($r.Arguments)"
        "Publisher:   $($r.Publisher)"
        "Signer:      " + ((@($r.ExePaths) | ForEach-Object { "$_ => $(Get-Signer $_)" }) -join '; ')
        "Accesses:    $($r.Accesses)"
        "Description: $($r.Description)"
    ) -join "`n"
    $ui.XmlBox.Text = $r.Xml
})

$window.Add_ContentRendered({ Load-Tasks })
$window.ShowDialog() | Out-Null
