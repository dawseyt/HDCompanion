# ============================================================================
# UI.ComputerActions.psm1 - Active Directory Computer Interface Logic
# ============================================================================

function Register-ComputerUIEvents {
    param($Window, $Config, $State)

    $AppRoot = Split-Path -Path $PSScriptRoot -Parent

    $lvData = $Window.FindName("lvData")
    $tabQuickAsset = $Window.FindName("tabQuickAsset")
    $txtTabSysInfo = $Window.FindName("txtTabSysInfo")
    $lvTabProcesses = $Window.FindName("lvTabProcesses")
    $lvTabServices = $Window.FindName("lvTabServices")
    
    $ctxGPResult = $Window.FindName("ctxGPResult")
    $ctxUptime = $Window.FindName("ctxUptime")
    $ctxViewProcesses = $Window.FindName("ctxViewProcesses")
    $ctxPrinterMenu = $Window.FindName("ctxPrinterMenu")
    $ctxPowerMenu = $Window.FindName("ctxPowerMenu")
    $ctxRestartComputer = $Window.FindName("ctxRestartComputer")
    $ctxShutdownComputer = $Window.FindName("ctxShutdownComputer")
    $ctxActiveUsers = $Window.FindName("ctxActiveUsers")

    if ($lvData -and $lvData.ContextMenu) {
        $lvData.AddHandler([System.Windows.Controls.Control]::ContextMenuOpeningEvent, [System.Windows.Controls.ContextMenuEventHandler]{
            $sel = $lvData.SelectedItem
            $isComp = ($sel -and $sel.Type -eq "Computer")
            $vis = if ($isComp) { "Visible" } else { "Collapsed" }
            
            if ($ctxGPResult) { $ctxGPResult.Visibility = $vis }
            if ($ctxUptime) { $ctxUptime.Visibility = $vis }
            if ($ctxViewProcesses) { $ctxViewProcesses.Visibility = $vis }
            if ($ctxPrinterMenu) { $ctxPrinterMenu.Visibility = $vis }
            if ($ctxPowerMenu) { $ctxPowerMenu.Visibility = $vis }
            if ($ctxActiveUsers) { $ctxActiveUsers.Visibility = $vis }
            
            if ($isComp -and $ctxActiveUsers) {
                $ctxActiveUsers.Tag = $null
                $ctxActiveUsers.Items.Clear()
                $dummyItem = New-Object System.Windows.Controls.MenuItem
                $dummyItem.Header = "Loading..."
                $ctxActiveUsers.Items.Add($dummyItem) | Out-Null
            }
        }.GetNewClosure())
    }

    # --- LAZY LOADING TAB LOGIC ---
    $LoadActiveTab = {
        $comp = $State.CurrentTabComputer
        if (-not $comp -or -not $tabQuickAsset) { return }
        
        $idx = $tabQuickAsset.SelectedIndex
        
        if ($idx -eq 0 -and $txtTabSysInfo) { 
            if ($txtTabSysInfo.Text -match "Select a computer" -or $txtTabSysInfo.Text -match "Loading") {
                $txtTabSysInfo.Text = "Querying $comp..."
                $action = [System.Action]{
                    try {
                        $job = Start-Job -ScriptBlock {
                            param($c)
                            $os = Get-CimInstance Win32_OperatingSystem -ComputerName $c -ErrorAction Stop
                            $cs = Get-CimInstance Win32_ComputerSystem -ComputerName $c -ErrorAction Stop
                            return "Host: $($cs.Name)`nOS: $($os.Caption)`nVersion: $($os.Version)`nModel: $($cs.Model)`nRAM: $([math]::Round($cs.TotalPhysicalMemory / 1GB, 2)) GB"
                        } -ArgumentList $comp
                        
                        $timeout = 50
                        while ($job.State -eq 'Running' -and $timeout -gt 0) { Start-Sleep -Milliseconds 100; $timeout-- }
                        if ($job.State -eq 'Running') { Stop-Job $job -ErrorAction SilentlyContinue; throw "Connection timed out (5s)." }
                        
                        $res = Receive-Job $job -ErrorAction Stop
                        $txtTabSysInfo.Dispatcher.Invoke({ $txtTabSysInfo.Text = $res })
                    } catch { 
                        $err = $_.Exception.Message
                        $txtTabSysInfo.Dispatcher.Invoke({ $txtTabSysInfo.Text = "Offline or Access Denied:`n$err" }) 
                    } finally { if ($job) { Remove-Job $job -Force -ErrorAction SilentlyContinue } }
                }
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, $action) | Out-Null
            }
        } elseif ($idx -eq 1 -and $lvTabProcesses) {
            if (-not $lvTabProcesses.ItemsSource) {
                [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                $action = [System.Action]{
                    try { $procs = Get-RemoteProcesses -ComputerName $comp; $lvTabProcesses.ItemsSource = $procs } catch {}
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                }
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, $action) | Out-Null
            }
        } elseif ($idx -eq 2 -and $lvTabServices) {
            if (-not $lvTabServices.ItemsSource) {
                [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                $action = [System.Action]{
                    try { $svcs = Get-RemoteServices -ComputerName $comp; $lvTabServices.ItemsSource = $svcs } catch {}
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                }
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, $action) | Out-Null
            }
        }
    }.GetNewClosure()

    if ($tabQuickAsset) { $tabQuickAsset.Add_SelectionChanged({ param($sender, $e) if ($e.OriginalSource -eq $tabQuickAsset) { & $LoadActiveTab } }.GetNewClosure()) }

    if ($lvData) {
        $lvData.Add_SelectionChanged({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                if ($tabQuickAsset) { $tabQuickAsset.Visibility = "Visible" }
                $State.CurrentTabComputer = $lvData.SelectedItem.Name
                if ($lvTabProcesses) { $lvTabProcesses.ItemsSource = $null }
                if ($lvTabServices) { $lvTabServices.ItemsSource = $null }
                if ($txtTabSysInfo) { $txtTabSysInfo.Text = "Loading data for $($State.CurrentTabComputer)..." }
                & $LoadActiveTab
            } else {
                if ($tabQuickAsset) { $tabQuickAsset.Visibility = "Collapsed" }
            }
        }.GetNewClosure())
    }

    # --- GPResult ---
    if ($ctxGPResult) {
        $ctxGPResult.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                $comp = $lvData.SelectedItem.Name
                $colors = Get-FluentThemeColors $State
                $gpWin = Load-XamlWindow -XamlPath (Join-Path $AppRoot "UI\Dialogs\GPResultDialog.xaml") -ThemeColors $colors
                $gpWin.Owner = $Window
                
                $btnCancel = $gpWin.FindName("btnCancel"); if ($btnCancel) { $btnCancel.Add_Click({ $gpWin.Close() }.GetNewClosure()) }
                $btnOk = $gpWin.FindName("btnOk")
                if ($btnOk) {
                    $btnOk.Add_Click({
                        $txtUser = $gpWin.FindName("txtUser"); $user = if ($txtUser) { $txtUser.Text } else { "" }; $gpWin.Close()
                        $loadingXaml = @"
                        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Processing" Width="320" Height="120" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                            <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                                <Border.Effect><DropShadowEffect BlurRadius="15" ShadowDepth="4" Opacity="0.3"/></Border.Effect>
                                <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                                    <TextBlock Text="Generating GPResult Report..." FontSize="14" FontWeight="SemiBold" Foreground="{Theme_Fg}" HorizontalAlignment="Center" Margin="0,0,0,12"/>
                                    <ProgressBar IsIndeterminate="True" Width="240" Height="4" Foreground="{Theme_PrimaryBg}" Background="{Theme_BtnBg}" BorderThickness="0"/>
                                </StackPanel>
                            </Border>
                        </Window>
"@
                        $xamlText = $loadingXaml; foreach ($key in $colors.Keys) { $xamlText = $xamlText.Replace("{Theme_$key}", $colors[$key]) }
                        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlText))
                        $loadWin = [System.Windows.Markup.XamlReader]::Load($reader); $loadWin.Owner = $Window; $loadWin.Show() 
                        
                        $tempFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "GPResult_$(Get-Date -Format 'HHmmss')_$comp.html")
                        $job = Start-Job -ScriptBlock {
                            param($c, $u, $f)
                            $args = @("/S", $c); if (-not [string]::IsNullOrWhiteSpace($u)) { $args += "/USER"; $args += $u } else { $args += "/SCOPE"; $args += "COMPUTER" }
                            $args += "/H"; $args += $f; $args += "/F"
                            $p = Start-Process -FilePath "gpresult.exe" -ArgumentList $args -NoNewWindow -PassThru -Wait
                            if ($p.ExitCode -ne 0) { throw "Exit code $($p.ExitCode)" }
                        } -ArgumentList $comp, $user, $tempFile
                        
                        $timer = New-Object System.Windows.Threading.DispatcherTimer
                        $timer.Interval = [TimeSpan]::FromMilliseconds(500)
                        $timerTick = {
                            if ($job.State -ne 'Running') {
                                $timer.Stop(); $loadWin.Close()
                                if ($job.State -eq 'Completed' -and (Test-Path $tempFile)) { Start-Process "msedge.exe" -ArgumentList "--app=""$tempFile""" } 
                                else {
                                    $reason = if ($job.ChildJobs[0].JobStateInfo.Reason) { $job.ChildJobs[0].JobStateInfo.Reason.Message } else { "Unknown error" }
                                    Show-AppMessageBox -Message "GPResult failed:`n$reason" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors
                                }
                                Remove-Job $job -Force
                            }
                        }.GetNewClosure()
                        $timer.Add_Tick($timerTick); $timer.Start()
                    }.GetNewClosure())
                }
                $gpWin.Show()
            }
        }.GetNewClosure())
    }

    # --- Uptime ---
    if ($ctxUptime) {
        $ctxUptime.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                $computer = $lvData.SelectedItem.Name
                $colors = Get-FluentThemeColors $State
                
                $loadingXaml = @"
                <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Processing" Width="320" Height="120" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                    <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                        <Border.Effect><DropShadowEffect BlurRadius="15" ShadowDepth="4" Opacity="0.3"/></Border.Effect>
                        <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                            <TextBlock Text="Querying Uptime for $computer..." FontSize="14" FontWeight="SemiBold" Foreground="{Theme_Fg}" HorizontalAlignment="Center" Margin="0,0,0,12"/>
                            <ProgressBar IsIndeterminate="True" Width="240" Height="4" Foreground="{Theme_PrimaryBg}" Background="{Theme_BtnBg}" BorderThickness="0"/>
                        </StackPanel>
                    </Border>
                </Window>
"@
                $xamlText = $loadingXaml; foreach ($key in $colors.Keys) { $xamlText = $xamlText.Replace("{Theme_$key}", $colors[$key]) }
                $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlText))
                $loadWin = [System.Windows.Markup.XamlReader]::Load($reader); $loadWin.Owner = $Window; $loadWin.Show() 

                $job = Start-Job -ScriptBlock {
                    param($c)
                    try {
                        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $c -ErrorAction Stop
                        $lastBoot = $os.LastBootUpTime; $uptime = (Get-Date) - $lastBoot
                        return [PSCustomObject]@{ Success = $true; LastBootUpTime = $lastBoot; Days = $uptime.Days; Hours = $uptime.Hours; Minutes = $uptime.Minutes }
                    } catch { return [PSCustomObject]@{ Success = $false; ErrorMessage = $_.Exception.Message } }
                } -ArgumentList $computer

                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(500); $startTime = Get-Date
                
                $timerTick = {
                    if ($job.State -ne 'Running' -or ((Get-Date) - $startTime).TotalSeconds -ge 10) {
                        $timer.Stop(); $loadWin.Close()
                        if ($job.State -eq 'Completed') {
                            $up = Receive-Job $job -ErrorAction SilentlyContinue
                            if ($up -and $up.Success) { Show-AppMessageBox -Message "Computer: $computer`n`nStatus: ONLINE`nLast Boot: $($up.LastBootUpTime)`n`nUptime: $($up.Days) days, $($up.Hours) hours, $($up.Minutes) minutes" -Title "Uptime Check" -IconType "Information" -OwnerWindow $Window -ThemeColors $colors | Out-Null } 
                            elseif ($up -and -not $up.Success) { Show-AppMessageBox -Message "Failed to query uptime for $computer.`n`nError: $($up.ErrorMessage)" -Title "Connection Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null } 
                            else { Show-AppMessageBox -Message "Failed to retrieve uptime data." -Title "Connection Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null }
                        } else {
                            Stop-Job $job -ErrorAction SilentlyContinue
                            $reason = "Connection timed out after 10 seconds or the job crashed."
                            if ($job.ChildJobs[0].JobStateInfo.Reason) { $reason = $job.ChildJobs[0].JobStateInfo.Reason.Message }
                            Show-AppMessageBox -Message "Failed to query uptime for $computer.`n`nError: $reason" -Title "Connection Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors | Out-Null
                        }
                        Remove-Job $job -Force -ErrorAction SilentlyContinue
                    }
                }.GetNewClosure()
                $timer.Add_Tick($timerTick); $timer.Start()
            }
        }.GetNewClosure())
    }

    # --- System Manager ---
    if ($ctxViewProcesses) {
        $ctxViewProcesses.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                $comp = $lvData.SelectedItem.Name
                $colors = Get-FluentThemeColors $State
                
                $procWin = Load-XamlWindow -XamlPath (Join-Path $AppRoot "UI\Windows\ProcessManager.xaml") -ThemeColors $colors
                $procWin.Owner = $Window
                $procWin.Title = "System Manager - $comp"
                
                $lblHeaderTitle = $procWin.FindName("lblHeaderTitle"); if ($lblHeaderTitle) { $lblHeaderTitle.Text = "Managing $comp" }
                
                $tabControlMain = $procWin.FindName("tabControlMain")
                $lvSoftware = $procWin.FindName("lvSoftware")
                $lblDiskSpace = $procWin.FindName("lblDiskSpace")
                $lvProcesses = $procWin.FindName("lvProcesses")
                $lvServices = $procWin.FindName("lvServices")
                $lvProfiles = $procWin.FindName("lvProfiles")
                $lvDevices = $procWin.FindName("lvDevices")
                $lvEvents = $procWin.FindName("lvEvents")
                $lblProcStatus = $procWin.FindName("lblProcStatus")
                $lblUptime = $procWin.FindName("lblUptime")
                
                $btnStartProcess = $procWin.FindName("btnStartProcess")
                $btnRefreshProcs = $procWin.FindName("btnRefreshProcs")
                $btnCloseProcs = $procWin.FindName("btnCloseProcs")
                $chkAutoRefreshProcs = $procWin.FindName("chkAutoRefreshProcs")
                
                $ctxKillProcess = $procWin.FindName("ctxKillProcess")
                $ctxStartService = $procWin.FindName("ctxStartService")
                $ctxStopService = $procWin.FindName("ctxStopService")
                $ctxRestartService = $procWin.FindName("ctxRestartService")
                $ctxDeleteProfile = $procWin.FindName("ctxDeleteProfile")
                $ctxUninstallSoftware = $procWin.FindName("ctxUninstallSoftware")
                $ctxEnableDevice = $procWin.FindName("ctxEnableDevice")
                $ctxDisableDevice = $procWin.FindName("ctxDisableDevice")
                
                $State.SoftLastSortCol = $null; $State.SoftSortDesc = $false
                $State.ProcLastSortCol = $null; $State.ProcSortDesc = $false
                $State.SvcLastSortCol = $null; $State.SvcSortDesc = $false
                $State.ProfLastSortCol = $null; $State.ProfSortDesc = $false
                $State.DevLastSortCol = $null;  $State.DevSortDesc = $false
                $State.EvtLastSortCol = $null;  $State.EvtSortDesc = $false
                $State.IsProcRefreshing = $false; $State.LastTabIndex = 0

                $DoRefresh = {
                    if ($State.IsProcRefreshing) { return }
                    $State.IsProcRefreshing = $true
                    $idx = if ($tabControlMain) { $tabControlMain.SelectedIndex } else { 0 }
                    
                    if ($lblProcStatus) { 
                        if ($idx -eq 0) { $lblProcStatus.Text = "Refreshing software inventory..." }
                        elseif ($idx -eq 1) { $lblProcStatus.Text = "Refreshing processes..." }
                        elseif ($idx -eq 2) { $lblProcStatus.Text = "Refreshing services..." }
                        elseif ($idx -eq 3) { $lblProcStatus.Text = "Refreshing user profiles..." }
                        elseif ($idx -eq 4) { $lblProcStatus.Text = "Refreshing hardware devices..." }
                        elseif ($idx -eq 5) { $lblProcStatus.Text = "Querying recent warnings & errors..." }
                    }
                    
                    if ($btnRefreshProcs) { $btnRefreshProcs.IsEnabled = $false }
                    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                    
                    $frame = New-Object System.Windows.Threading.DispatcherFrame
                    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Action]{ $frame.Continue = $false }) | Out-Null
                    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
                    
                    try {
                        if ($idx -eq 0) {
                            $disk = Get-RemoteDiskSpace -ComputerName $comp
                            if ($lblDiskSpace -and $disk) {
                                $sizeGB = [math]::Round($disk.Size / 1GB, 1); $freeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
                                $lblDiskSpace.Text = "C:\ Drive: $freeGB GB Free out of $sizeGB GB ($([math]::Round(($freeGB / $sizeGB) * 100, 1))% Free)"
                            }
                            $rawSoft = Get-RemoteInstalledSoftware -ComputerName $comp; $resSoft = @()
                            if ($rawSoft) { foreach ($r in $rawSoft) { $resSoft += [PSCustomObject]@{ DisplayName = $r.DisplayName; DisplayVersion = $r.DisplayVersion; Publisher = $r.Publisher; InstallDate = $r.InstallDate; UninstallString = $r.UninstallString; QuietUninstallString = $r.QuietUninstallString } } }
                            if ($lvSoftware) { if ($State.SoftLastSortCol) { $resSoft = $resSoft | Sort-Object -Property $State.SoftLastSortCol -Descending:$State.SoftSortDesc }; $lvSoftware.ItemsSource = $resSoft }
                        } elseif ($idx -eq 1) {
                            $rawProcs = Get-RemoteProcesses -ComputerName $comp; $resProcs = @()
                            if ($rawProcs) { foreach ($r in $rawProcs) { $resProcs += [PSCustomObject]@{ Name = $r.Name; Id = $r.Id; CPU = $r.CPU; MemMB = $r.MemMB; Description = $r.Description } } }
                            if ($lvProcesses) { if ($State.ProcLastSortCol) { $resProcs = $resProcs | Sort-Object -Property $State.ProcLastSortCol -Descending:$State.ProcSortDesc }; $lvProcesses.ItemsSource = $resProcs }
                        } elseif ($idx -eq 2) {
                            $rawSvcs = Get-RemoteServices -ComputerName $comp; $resSvcs = @()
                            if ($rawSvcs) { foreach ($s in $rawSvcs) { $resSvcs += [PSCustomObject]@{ Name = $s.Name; DisplayName = $s.DisplayName; State = $s.State; StartMode = $s.StartMode } } }
                            if ($lvServices) { if ($State.SvcLastSortCol) { $resSvcs = $resSvcs | Sort-Object -Property $State.SvcLastSortCol -Descending:$State.SvcSortDesc } else { $resSvcs = $resSvcs | Sort-Object -Property Name }; $lvServices.ItemsSource = $resSvcs }
                        } elseif ($idx -eq 3) {
                            $rawProfs = Get-RemoteUserProfiles -ComputerName $comp; $resProfs = @()
                            if ($rawProfs) { foreach ($p in $rawProfs) { $resProfs += [PSCustomObject]@{ LocalPath = $p.LocalPath; LastUseTime = $p.LastUseTime; Loaded = $p.Loaded; SID = $p.SID } } }
                            if ($lvProfiles) { if ($State.ProfLastSortCol) { $resProfs = $resProfs | Sort-Object -Property $State.ProfLastSortCol -Descending:$State.ProfSortDesc }; $lvProfiles.ItemsSource = $resProfs }
                        } elseif ($idx -eq 4) {
                            $rawDevs = Get-RemoteDevices -ComputerName $comp; $resDevs = @()
                            if ($rawDevs) { foreach ($d in $rawDevs) { $resDevs += [PSCustomObject]@{ FriendlyName = $d.FriendlyName; Class = $d.Class; Status = $d.Status; Manufacturer = $d.Manufacturer; InstanceId = $d.InstanceId } } }
                            if ($lvDevices) { if ($State.DevLastSortCol) { $resDevs = $resDevs | Sort-Object -Property $State.DevLastSortCol -Descending:$State.DevSortDesc } else { $resDevs = $resDevs | Sort-Object -Property Class, FriendlyName }; $lvDevices.ItemsSource = $resDevs }
                        } elseif ($idx -eq 5) {
                            $rawEvts = Get-RemoteEventLogs -ComputerName $comp; $resEvts = @()
                            if ($rawEvts) { foreach ($e in $rawEvts) { $resEvts += [PSCustomObject]@{ TimeCreated = $e.TimeCreated; Level = $e.LevelDisplayName; Id = $e.Id; Source = $e.ProviderName; Message = if ($e.Message) { $e.Message -replace "`r", "" -replace "`n", "  " } else { "" } } } }
                            if ($lvEvents) { if ($State.EvtLastSortCol) { $resEvts = $resEvts | Sort-Object -Property $State.EvtLastSortCol -Descending:$State.EvtSortDesc }; $lvEvents.ItemsSource = $resEvts }
                        }

                        try {
                            $upJob = Start-Job -ScriptBlock {
                                param($c)
                                try { $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $c -ErrorAction Stop; $lastBoot = $os.LastBootUpTime; $uptime = (Get-Date) - $lastBoot; return [PSCustomObject]@{ Success=$true; Days=$uptime.Days; Hours=$uptime.Hours; Minutes=$uptime.Minutes } } 
                                catch { return [PSCustomObject]@{ Success=$false; ErrorMessage=$_.Exception.Message } }
                            } -ArgumentList $comp
                            
                            $tCount = 40; while ($upJob.State -eq 'Running' -and $tCount -gt 0) { Start-Sleep -Milliseconds 100; $tCount-- }
                            if ($upJob.State -eq 'Completed') {
                                $upData = Receive-Job $upJob -ErrorAction SilentlyContinue
                                if ($upData -and $upData.Success) { if ($lblUptime) { $lblUptime.Text = "Uptime: $($upData.Days) days, $($upData.Hours) hours, $($upData.Minutes) minutes" } } 
                                else { $msg = if ($upData -and $upData.ErrorMessage) { $upData.ErrorMessage } else { "Unknown Error" }; if ($lblUptime) { $lblUptime.Text = "Uptime: Unavailable ($msg)" } }
                            } else { Stop-Job $upJob -ErrorAction SilentlyContinue; if ($lblUptime) { $lblUptime.Text = "Uptime: Timeout (No Response)" } }
                            Remove-Job $upJob -Force -ErrorAction SilentlyContinue
                        } catch { if ($lblUptime) { $lblUptime.Text = "Uptime: Error - $($_.Exception.Message)" } }

                        if ($lblProcStatus) { $autoMode = if ($chkAutoRefreshProcs -and $chkAutoRefreshProcs.IsChecked) { " (Auto-refresh: 15s)" } else { "" }; $lblProcStatus.Text = "Last updated: $(Get-Date -Format 'HH:mm:ss')$autoMode" }
                    } catch { if ($lblProcStatus) { $lblProcStatus.Text = "Error: $($_.Exception.Message)" } }
                    
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    if ($btnRefreshProcs) { $btnRefreshProcs.IsEnabled = $true }
                    $State.IsProcRefreshing = $false
                }.GetNewClosure()
                
                if ($btnRefreshProcs) { $btnRefreshProcs.Add_Click($DoRefresh) }
                if ($btnCloseProcs) { $btnCloseProcs.Add_Click({ $procWin.Close() }.GetNewClosure()) }
                if ($tabControlMain) { $tabControlMain.Add_SelectionChanged({ param($sender, $e) if ($e.OriginalSource -eq $tabControlMain -and $State.LastTabIndex -ne $tabControlMain.SelectedIndex) { $State.LastTabIndex = $tabControlMain.SelectedIndex; & $DoRefresh } }.GetNewClosure()) }

                $ListSortAction = {
                    param($sender, $e)
                    $source = $e.OriginalSource
                    while ($source -and -not ($source -is [System.Windows.Controls.GridViewColumnHeader])) { if ($source -is [System.Windows.FrameworkElement]) { $source = $source.Parent } else { break } }
                    if ($source -and ($source -is [System.Windows.Controls.GridViewColumnHeader]) -and $source.Role -ne "Padding") {
                        $column = $source.Column
                        if ($column -and $column.DisplayMemberBinding) {
                            $sortBy = $column.DisplayMemberBinding.Path.Path
                            $isDesc = $false
                            if ($sender.Name -eq "lvSoftware") { if ($State.SoftLastSortCol -eq $sortBy) { $State.SoftSortDesc = -not $State.SoftSortDesc } else { $State.SoftSortDesc = $false; $State.SoftLastSortCol = $sortBy }; $isDesc = $State.SoftSortDesc }
                            elseif ($sender.Name -eq "lvProcesses") { if ($State.ProcLastSortCol -eq $sortBy) { $State.ProcSortDesc = -not $State.ProcSortDesc } else { $State.ProcSortDesc = $false; $State.ProcLastSortCol = $sortBy }; $isDesc = $State.ProcSortDesc }
                            elseif ($sender.Name -eq "lvServices") { if ($State.SvcLastSortCol -eq $sortBy) { $State.SvcSortDesc = -not $State.SvcSortDesc } else { $State.SvcSortDesc = $false; $State.SvcLastSortCol = $sortBy }; $isDesc = $State.SvcSortDesc }
                            elseif ($sender.Name -eq "lvProfiles") { if ($State.ProfLastSortCol -eq $sortBy) { $State.ProfSortDesc = -not $State.ProfSortDesc } else { $State.ProfSortDesc = $false; $State.ProfLastSortCol = $sortBy }; $isDesc = $State.ProfSortDesc }
                            elseif ($sender.Name -eq "lvDevices") { if ($State.DevLastSortCol -eq $sortBy) { $State.DevSortDesc = -not $State.DevSortDesc } else { $State.DevSortDesc = $false; $State.DevLastSortCol = $sortBy }; $isDesc = $State.DevSortDesc }
                            elseif ($sender.Name -eq "lvEvents") { if ($State.EvtLastSortCol -eq $sortBy) { $State.EvtSortDesc = -not $State.EvtSortDesc } else { $State.EvtSortDesc = $false; $State.EvtLastSortCol = $sortBy }; $isDesc = $State.EvtSortDesc }

                            if ($sender.ItemsSource) { $items = @($sender.ItemsSource); if ($items.Count -gt 0) { $sorted = $items | Sort-Object -Property $sortBy -Descending:$isDesc; $sender.ItemsSource = @($sorted) } }
                        }
                    }
                }.GetNewClosure()
                
                if ($lvSoftware) { $lvSoftware.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$ListSortAction) }
                if ($lvProcesses) { $lvProcesses.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$ListSortAction) }
                if ($lvServices) { $lvServices.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$ListSortAction) }
                if ($lvProfiles) { $lvProfiles.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$ListSortAction) }
                if ($lvDevices) { $lvDevices.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$ListSortAction) }
                if ($lvEvents) { $lvEvents.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$ListSortAction) }

                # Buttons inside System Manager
                if ($ctxUninstallSoftware) {
                    $ctxUninstallSoftware.Add_Click({
                        if ($lvSoftware -and $lvSoftware.SelectedItem) {
                            $app = $lvSoftware.SelectedItem; $dispName = $app.DisplayName
                            $conf = Show-AppMessageBox -Message "Are you sure you want to silently uninstall '$dispName' from $comp?" -Title "Confirm Uninstall" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $procWin -ThemeColors $colors
                            if ($conf -eq "Yes") {
                                try { Uninstall-RemoteSoftware -ComputerName $comp -QuietUninstallString $app.QuietUninstallString -UninstallString $app.UninstallString; Show-AppMessageBox -Message "Uninstall command triggered." -Title "Success" -ThemeColors $colors; & $DoRefresh }
                                catch { Show-AppMessageBox -Message "Uninstall failed:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -ThemeColors $colors }
                            }
                        }
                    }.GetNewClosure())
                }
                
                if ($btnStartProcess) {
                    $btnStartProcess.Add_Click({
                        $inputXaml = @"
                        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Run Task" Width="380" SizeToContent="Height" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                            <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                                <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3"/></Border.Effect>
                                <Grid>
                                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                                    <Border Grid.Row="0" Background="Transparent" Padding="16,16,16,8"><TextBlock Text="Run New Task on $($comp)" FontSize="15" FontWeight="SemiBold" Foreground="{Theme_Fg}"/></Border>
                                    <StackPanel Grid.Row="1" Margin="16,8,16,16"><TextBlock Text="Executable or Command Line:" FontSize="12" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><TextBox x:Name="txtCmd" Height="30" Background="{Theme_BtnBg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Padding="6,4" VerticalContentAlignment="Center"/></StackPanel>
                                    <Border Grid.Row="2" Background="{Theme_BtnBg}" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="{Theme_BtnBorder}"><StackPanel Orientation="Horizontal" HorizontalAlignment="Right"><Button x:Name="btnOk" Content="Run" Width="80" Height="28" Margin="0,0,8,0" Background="{Theme_PrimaryBg}" Foreground="{Theme_PrimaryFg}" BorderThickness="0" IsDefault="True"/><Button x:Name="btnCancel" Content="Cancel" Width="80" Height="28" Background="{Theme_Bg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" IsCancel="True"/></StackPanel></Border>
                                </Grid>
                            </Border>
                        </Window>
"@
                        foreach ($key in $colors.Keys) { $inputXaml = $inputXaml.Replace("{Theme_$key}", $colors[$key]) }
                        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($inputXaml)); $inpWin = [System.Windows.Markup.XamlReader]::Load($reader); $inpWin.Owner = $procWin
                        
                        $btnOk = $inpWin.FindName("btnOk"); $btnCancel = $inpWin.FindName("btnCancel"); $txtCmd = $inpWin.FindName("txtCmd")
                        if ($btnCancel) { $btnCancel.Add_Click({ $inpWin.Close() }.GetNewClosure()) }
                        if ($btnOk) {
                            $btnOk.Add_Click({
                                $cmd = $txtCmd.Text; $inpWin.Close()
                                if ([string]::IsNullOrWhiteSpace($cmd)) { return }
                                try { Start-RemoteProcess -ComputerName $comp -CommandLine $cmd; Add-AppLog -Event "Task Started" -Username "System" -Details "Executed '$cmd' on $comp." -Config $Config -State $State -Status "Success"; & $DoRefresh } 
                                catch { Show-AppMessageBox -Message "Failed to start task:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $procWin -ThemeColors $colors }
                            }.GetNewClosure())
                        }
                        $inpWin.Show()
                    }.GetNewClosure())
                }

                if ($ctxKillProcess) { $ctxKillProcess.Add_Click({ if ($lvProcesses -and $lvProcesses.SelectedItem) { $pidToKill = $lvProcesses.SelectedItem.Id; $procName = $lvProcesses.SelectedItem.Name; $conf = Show-AppMessageBox -Message "Kill process '$procName' (PID: $pidToKill) on $comp?" -Title "Confirm" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $procWin -ThemeColors $colors; if ($conf -eq "Yes") { try { Stop-RemoteProcess -ComputerName $comp -ProcessId $pidToKill; Show-AppMessageBox -Message "Process killed." -Title "Success" -ThemeColors $colors; & $DoRefresh } catch { Show-AppMessageBox -Message "Failed to kill process: $($_.Exception.Message)" -Title "Error" -IconType "Error" -ThemeColors $colors } } } }.GetNewClosure()) }
                if ($ctxDeleteProfile) { $ctxDeleteProfile.Add_Click({ if ($lvProfiles -and $lvProfiles.SelectedItem) { $pPath = $lvProfiles.SelectedItem.LocalPath; $pSID = $lvProfiles.SelectedItem.SID; $conf = Show-AppMessageBox -Message "Permanently delete user profile '$pPath' on $comp?`n`nThis cannot be undone." -Title "Confirm Delete" -ButtonType "YesNo" -IconType "Error" -OwnerWindow $procWin -ThemeColors $colors; if ($conf -eq "Yes") { try { Remove-RemoteUserProfile -ComputerName $comp -SID $pSID; Show-AppMessageBox -Message "Profile deleted." -Title "Success" -ThemeColors $colors; & $DoRefresh } catch { Show-AppMessageBox -Message "Failed to delete profile: $($_.Exception.Message)" -Title "Error" -IconType "Error" -ThemeColors $colors } } } }.GetNewClosure()) }

                $ExecuteServiceAction = { param($ActionName) if ($lvServices -and $lvServices.SelectedItem) { $svcName = $lvServices.SelectedItem.Name; $dispName = $lvServices.SelectedItem.DisplayName; $conf = Show-AppMessageBox -Message "Are you sure you want to $ActionName the service '$dispName' on $comp?" -Title "Confirm $ActionName" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $procWin -ThemeColors $colors; if ($conf -eq "Yes") { [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait; try { Invoke-RemoteServiceAction -ComputerName $comp -ServiceName $svcName -Action $ActionName; [System.Windows.Input.Mouse]::OverrideCursor = $null; Show-AppMessageBox -Message "Service command sent successfully." -Title "Success" -ThemeColors $colors; & $DoRefresh } catch { [System.Windows.Input.Mouse]::OverrideCursor = $null; Show-AppMessageBox -Message "Service error:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -ThemeColors $colors } } } }.GetNewClosure()
                if ($ctxStartService) { $ctxStartService.Add_Click({ & $ExecuteServiceAction "Start" }.GetNewClosure()) }
                if ($ctxStopService) { $ctxStopService.Add_Click({ & $ExecuteServiceAction "Stop" }.GetNewClosure()) }
                if ($ctxRestartService) { $ctxRestartService.Add_Click({ & $ExecuteServiceAction "Restart" }.GetNewClosure()) }
                
                $ExecuteDeviceAction = { param($ActionName) if ($lvDevices -and $lvDevices.SelectedItem) { $devName = $lvDevices.SelectedItem.FriendlyName; $devId = $lvDevices.SelectedItem.InstanceId; $conf = Show-AppMessageBox -Message "Are you sure you want to $ActionName the device '$devName' on $comp?" -Title "Confirm $ActionName" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $procWin -ThemeColors $colors; if ($conf -eq "Yes") { [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait; try { Set-RemoteDeviceState -ComputerName $comp -InstanceId $devId -Action $ActionName; [System.Windows.Input.Mouse]::OverrideCursor = $null; Show-AppMessageBox -Message "Device $ActionName command sent successfully." -Title "Success" -ThemeColors $colors; & $DoRefresh } catch { [System.Windows.Input.Mouse]::OverrideCursor = $null; Show-AppMessageBox -Message "Device error:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -ThemeColors $colors } } } }.GetNewClosure()
                if ($ctxEnableDevice) { $ctxEnableDevice.Add_Click({ & $ExecuteDeviceAction "Enable" }.GetNewClosure()) }
                if ($ctxDisableDevice) { $ctxDisableDevice.Add_Click({ & $ExecuteDeviceAction "Disable" }.GetNewClosure()) }

                & $DoRefresh
                $procAutoTimer = New-Object System.Windows.Threading.DispatcherTimer
                $procAutoTimer.Interval = [TimeSpan]::FromSeconds(15)
                $procAutoTimer.Add_Tick({ & $DoRefresh }.GetNewClosure())
                
                if ($chkAutoRefreshProcs) {
                    $chkAutoRefreshProcs.Add_Checked({ $procAutoTimer.Start(); & $DoRefresh }.GetNewClosure())
                    $chkAutoRefreshProcs.Add_Unchecked({ $procAutoTimer.Stop() }.GetNewClosure())
                    if ($chkAutoRefreshProcs.IsChecked -eq $true) { $procAutoTimer.Start() }
                }
                
                $procWin.Add_Closed({ if ($procAutoTimer) { $procAutoTimer.Stop() } }.GetNewClosure())
                $procWin.Show()
            }
        }.GetNewClosure())
    }

    # --- Other Actions ---
    if ($ctxRestartComputer) { $ctxRestartComputer.Add_Click({ $comp = $lvData.SelectedItem.Name; $conf = Show-AppMessageBox -Message "Restart $comp?" -Title "Confirm" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State); if ($conf -eq "Yes") { try { Invoke-RemotePowerAction -ComputerName $comp -Action "Restart" } catch { Show-AppMessageBox -Message "Failed: $($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) } } }.GetNewClosure()) }
    if ($ctxShutdownComputer) { $ctxShutdownComputer.Add_Click({ $comp = $lvData.SelectedItem.Name; $conf = Show-AppMessageBox -Message "Shutdown $comp?" -Title "Confirm" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State); if ($conf -eq "Yes") { try { Invoke-RemotePowerAction -ComputerName $comp -Action "Shutdown" } catch { Show-AppMessageBox -Message "Failed: $($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) } } }.GetNewClosure()) }
    if ($ctxPrinterMenu) {
        $ctxPrinterMenu.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                $targetPC = $lvData.SelectedItem.Name
                $pmScript = Join-Path $AppRoot "PrinterManager.ps1"
                if (-not (Test-Path $pmScript)) { Show-AppMessageBox -Message "Script not found at:`n$pmScript" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State); return }
                Add-AppLog -Event "Printer Management" -Username "System" -Details "Launching Printer Manager for $targetPC..." -Config $Config -State $State -Status "Info"
                try { Start-Process -FilePath "powershell.exe" -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$pmScript`" -ComputerName `"$targetPC`"" -WindowStyle Hidden } 
                catch { Show-AppMessageBox -Message "Launch Failed:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
            }
        }.GetNewClosure())
    }

    if ($ctxActiveUsers) {
        $ctxActiveUsers.Add_SubmenuOpened({
            param($sender, $e)
            if (-not $lvData.SelectedItem -or $lvData.SelectedItem.Type -ne "Computer") { return }
            $comp = $lvData.SelectedItem.Name
            if ($sender.Tag -eq $comp) { return } 
            $sender.Items.Clear(); $loading = New-Object System.Windows.Controls.MenuItem; $loading.Header = "Querying..."; $loading.IsEnabled = $false; $sender.Items.Add($loading) | Out-Null
            $frame = New-Object System.Windows.Threading.DispatcherFrame
            [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Action]{ $frame.Continue = $false }) | Out-Null
            [System.Windows.Threading.Dispatcher]::PushFrame($frame)
            $sessions = Get-RemoteActiveUsers -ComputerName $comp
            $sender.Items.Clear()
            if (-not $sessions) { $noUsers = New-Object System.Windows.Controls.MenuItem; $noUsers.Header = "No active users"; $noUsers.IsEnabled = $false; $sender.Items.Add($noUsers) | Out-Null } 
            else {
                foreach ($s in $sessions) {
                    $uItem = New-Object System.Windows.Controls.MenuItem; $uItem.Header = "$($s.Username) (ID: $($s.SessionId), $($s.State))"
                    $lItem = New-Object System.Windows.Controls.MenuItem; $lItem.Header = "Logoff $($s.Username)"; $lItem.Foreground = [System.Windows.Media.Brushes]::Red
                    $closure = { param($bU, $bI, $bC) $action = { $conf = Show-AppMessageBox -Message "Logoff $bU from $bC?" -Title "Confirm" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State); if ($conf -eq "Yes") { try { Stop-RemoteUserSession -ComputerName $bC -SessionId $bI; Show-AppMessageBox -Message "Logged off." -Title "Success" -ThemeColors (Get-FluentThemeColors $State) } catch { Show-AppMessageBox -Message "Failed: $_" -Title "Error" -IconType "Error" -ThemeColors (Get-FluentThemeColors $State) } } }.GetNewClosure(); return $action }
                    $lItem.Add_Click((& $closure $s.Username $s.SessionId $comp))
                    $uItem.Items.Add($lItem) | Out-Null
                    $sender.Items.Add($uItem) | Out-Null
                }
            }
            $sender.Tag = $comp
        }.GetNewClosure())
    }
}
Export-ModuleMember -Function *