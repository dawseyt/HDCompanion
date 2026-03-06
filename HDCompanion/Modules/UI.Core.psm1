# ============================================================================
# UI.Core.psm1 - Base Application UI, Theming, Search, and Refresh
# ============================================================================

function Show-AppMessageBox {
    param(
        [string]$Message,
        [string]$Title = "Information",
        [string]$ButtonType = "OK", # OK, YesNo, OKCancel
        [string]$IconType = "Information", # Information, Warning, Error, Question
        $OwnerWindow,
        $ThemeColors
    )
    
    $xamlPath = Join-Path $PSScriptRoot "..\UI\Dialogs\MessageBox.xaml"
    if (-not (Test-Path $xamlPath)) {
        [System.Windows.MessageBox]::Show($Message, $Title)
        return "OK"
    }

    $msgWin = Load-XamlWindow -XamlPath $xamlPath -ThemeColors $ThemeColors
    
    $lblTitle = $msgWin.FindName("lblTitle")
    if ($lblTitle) { $lblTitle.Text = $Title }
    
    $txtMessageBody = $msgWin.FindName("txtMessageBody")
    if ($txtMessageBody) { $txtMessageBody.Text = $Message }
    
    $pathIcon = $msgWin.FindName("pathIcon")
    if ($pathIcon) {
        switch ($IconType) {
            "Error" { $pathIcon.Data = [System.Windows.Media.Geometry]::Parse("M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"); $pathIcon.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(209, 52, 56)) }
            "Warning" { $pathIcon.Data = [System.Windows.Media.Geometry]::Parse("M12 2L1 21h22L12 2zm0 3.83L19.53 19H4.47L12 5.83zM11 10h2v5h-2v-5zm0 6h2v2h-2v-2z"); $pathIcon.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(255, 185, 0)) }
            "Question" { $pathIcon.Data = [System.Windows.Media.Geometry]::Parse("M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 17h-2v-2h2v2zm2.07-7.75l-.9.92C13.45 12.9 13 13.5 13 15h-2v-.5c0-1.1.45-2.1 1.17-2.83l1.24-1.26c.37-.36.59-.86.59-1.41 0-1.1-.9-2-2-2s-2 .9-2 2H8c0-2.21 1.79-4 4-4s4 1.79 4 4c0 .88-.36 1.68-.93 2.25z"); $pathIcon.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0, 120, 212)) }
            Default { $pathIcon.Data = [System.Windows.Media.Geometry]::Parse("M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h-2v6h2zm0-8h-2V7h2v2z"); $pathIcon.Fill = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0, 120, 212)) }
        }
    }
    
    $msgWin.FindName("btnYes").Visibility = if ($ButtonType -eq "YesNo") { "Visible" } else { "Collapsed" }
    $msgWin.FindName("btnNo").Visibility = if ($ButtonType -eq "YesNo") { "Visible" } else { "Collapsed" }
    $msgWin.FindName("btnOk").Visibility = if ($ButtonType -in @("OK", "OKCancel")) { "Visible" } else { "Collapsed" }
    $msgWin.FindName("btnCancel").Visibility = if ($ButtonType -eq "OKCancel") { "Visible" } else { "Collapsed" }
    
    $script:msgRes = "Cancel"
    $msgWin.FindName("btnYes").Add_Click({ $script:msgRes = "Yes"; $msgWin.Close() })
    $msgWin.FindName("btnNo").Add_Click({ $script:msgRes = "No"; $msgWin.Close() })
    $msgWin.FindName("btnOk").Add_Click({ $script:msgRes = "OK"; $msgWin.Close() })
    $msgWin.FindName("btnCancel").Add_Click({ $script:msgRes = "Cancel"; $msgWin.Close() })
    $msgWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $msgWin.DragMove() })
    
    if ($OwnerWindow -and $OwnerWindow.IsVisible) { $msgWin.Owner = $OwnerWindow; $msgWin.WindowStartupLocation = "CenterOwner" }
    $msgWin.ShowDialog() | Out-Null
    return $script:msgRes
}

function Register-CoreUIEvents {
    param($Window, $Config, $State)

    $AppRoot = Split-Path -Path $PSScriptRoot -Parent

    $State.UIControls.txtLog = $Window.FindName("txtLog")
    $lvData = $Window.FindName("lvData")
    $gvData = $Window.FindName("gvData")
    $txtSearch = $Window.FindName("txtSearch")
    $btnSearch = $Window.FindName("btnSearch")
    $btnRefresh = $Window.FindName("btnRefresh")
    $btnViewLog = $Window.FindName("btnViewLog")
    $lblStatus = $Window.FindName("lblStatus")
    $btnThemeToggle = $Window.FindName("btnThemeToggle")
    $iconTheme = $Window.FindName("iconTheme")
    $cbAutoRefresh = $Window.FindName("cbAutoRefresh")
    $chkEnableEmail = $Window.FindName("chkEnableEmail")
    
    $btnDashboard = $Window.FindName("btnDashboard")
    $btnModifyConfig = $Window.FindName("btnModifyConfig")
    $btnHelp = $Window.FindName("btnHelp")
    $btnTechDocs = $Window.FindName("btnTechDocs")
    
    $ctxDetails = $Window.FindName("ctxDetails")
    $overlayDetails = $Window.FindName("overlayDetails")
    $borderDetails = $Window.FindName("borderDetails")
    $txtDetailsContent = $Window.FindName("txtDetailsContent")
    $btnCloseDetails = $Window.FindName("btnCloseDetails")
    $thumbResizeDetails = $Window.FindName("thumbResizeDetails")

    if ($txtDetailsContent) {
        $ctxCopy = New-Object System.Windows.Controls.ContextMenu
        $miCopy = New-Object System.Windows.Controls.MenuItem
        $miCopy.Header = "Copy All Details to Clipboard"
        $miCopy.Add_Click({
            if (-not [string]::IsNullOrWhiteSpace($txtDetailsContent.Text)) {
                [System.Windows.Clipboard]::SetText($txtDetailsContent.Text)
                Show-AppMessageBox -Message "Details copied to clipboard." -Title "Copied" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) | Out-Null
            }
        })
        $ctxCopy.Items.Add($miCopy) | Out-Null
        $txtDetailsContent.ContextMenu = $ctxCopy
    }

    $ApplyTheme = {
        param($TargetTheme)
        $res = $Window.Resources
        $themeData = if ($TargetTheme -eq "Light") { $Config.LightModeColors } else { $Config.DarkModeColors }
        
        function Get-ColorFromConfig ($rgbArray) {
            if ($rgbArray -and $rgbArray.Count -eq 3) { return [System.Windows.Media.Color]::FromRgb($rgbArray[0], $rgbArray[1], $rgbArray[2]) }
            return [System.Windows.Media.Colors]::Transparent 
        }

        $res["WindowBackground"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.Background))
        $res["CardBackground"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.Card))
        $res["AccentFill"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.Primary))
        $res["TextPrimary"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.Text))
        $res["TextSecondary"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.TextSecondary))
        $res["ControlStroke"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.Secondary))
        
        if ($themeData.Hover) { $res["HoverFill"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.Hover)) } 
        else { if ($TargetTheme -eq "Light") { $res["HoverFill"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(229,229,229)) } else { $res["HoverFill"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(80,80,85)) } }
        
        if ($themeData.AltRow) { $res["AltRowBg"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.AltRow)) } 
        else { if ($TargetTheme -eq "Light") { $res["AltRowBg"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(249,249,249)) } else { $res["AltRowBg"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(55,55,60)) } }

        if ($themeData.OnlineText) { $res["OnlineTextBrush"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.OnlineText)) }
        else { if ($TargetTheme -eq "Dark") { $res["OnlineTextBrush"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(150, 255, 150)) } else { $res["OnlineTextBrush"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0, 100, 0)) } }

        if ($themeData.OfflineText) { $res["OfflineTextBrush"] = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.OfflineText)) }
        else { if ($TargetTheme -eq "Dark") { $res["OfflineTextBrush"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(255, 150, 150)) } else { $res["OfflineTextBrush"] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(178, 34, 34)) } }
        
        if ($State.UIControls.txtLog) { $State.UIControls.txtLog.Foreground = [System.Windows.Media.SolidColorBrush]::new((Get-ColorFromConfig $themeData.Text)) }
        
        if ($iconTheme) {
            if ($TargetTheme -eq "Light") { $iconTheme.Data = [System.Windows.Media.Geometry]::Parse("M12 3c-4.97 0-9 4.03-9 9s4.03 9 9 9 9-4.03 9-9c0-.46-.04-.92-.1-1.36-.98 1.37-2.58 2.26-4.4 2.26-3.03 0-5.5-2.47-5.5-5.5 0-1.82.89-3.42 2.26-4.4-.44-.06-.9-.1-1.36-.1z") } 
            else { $iconTheme.Data = [System.Windows.Media.Geometry]::Parse("M6.76 4.84l-1.8-1.79-1.41 1.41 1.79 1.79 1.42-1.41zM4 10.5H1v2h3v-2zm9-9.95h-2V3.5h2V.55zm7.45 3.91l-1.41-1.41-1.79 1.79 1.41 1.41 1.79-1.79zm-3.21 13.7l1.79 1.8 1.41-1.41-1.8-1.79-1.4 1.4zM20 10.5v2h3v-2h-3zm-8-5c-3.31 0-6 2.69-6 6s2.69 6 6 6 6-2.69 6-6-2.69-6-6-6zm-1 16.95h2V19.5h-2v2.95zm-7.45-3.91l1.41 1.41 1.79-1.8-1.41-1.41-1.79 1.8z") }
        }
        $State.CurrentTheme = $TargetTheme
    }.GetNewClosure()

    $UpdateGridColumns = {
        if (-not $gvData) { return }
        $gvData.Columns.Clear()
        $cols = @(
            @{ Header="Name"; Binding="Name"; Width=160; HasStatusIndicator=$true },
            @{ Header="Type"; Binding="Type"; Width=80 },
            @{ Header="AD Description"; Binding="Description"; Width=200 },
            @{ Header="Locked / Sys Desc"; Binding="LockedOut"; Width=150 },
            @{ Header="Last Logon"; Binding="LastLogonDate"; Width=140; Format="{0:MM/dd/yyyy}" }
        )
        foreach ($c in $cols) {
            $col = New-Object System.Windows.Controls.GridViewColumn
            $col.Width = $c.Width
            if ($c.HasStatusIndicator) {
                $cellTemplate = @"
                <DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
                    <StackPanel Orientation="Horizontal">
                        <Ellipse Width="10" Height="10" Margin="0,0,6,0" VerticalAlignment="Center">
                            <Ellipse.Style>
                                <Style TargetType="Ellipse">
                                    <Setter Property="Visibility" Value="Collapsed"/>
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding IsOnline}" Value="True"><Setter Property="Fill" Value="#4CAF50"/><Setter Property="Visibility" Value="Visible"/></DataTrigger>
                                        <DataTrigger Binding="{Binding IsOnline}" Value="False"><Setter Property="Fill" Value="#E53935"/><Setter Property="Visibility" Value="Visible"/></DataTrigger>
                                    </Style.Triggers>
                                </Style>
                            </Ellipse.Style>
                        </Ellipse>
                        <TextBlock Text="{Binding $($c.Binding)}" VerticalAlignment="Center"/>
                    </StackPanel>
                </DataTemplate>
"@
                $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($cellTemplate))
                $col.CellTemplate = [System.Windows.Markup.XamlReader]::Load($reader)
            } else {
                $binding = New-Object System.Windows.Data.Binding($c.Binding)
                if ($c.Format) { $binding.StringFormat = $c.Format }
                $col.DisplayMemberBinding = $binding
            }
            $headerTemplate = "<DataTemplate xmlns=`"http://schemas.microsoft.com/winfx/2006/xaml/presentation`"><TextBlock Text=`"$($c.Header)`" HorizontalAlignment=`"Left`" Margin=`"5,0,0,0`"/></DataTemplate>"
            $headerReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($headerTemplate))
            $col.HeaderTemplate = [System.Windows.Markup.XamlReader]::Load($headerReader)
            $gvData.Columns.Add($col)
        }
    }.GetNewClosure()

    $SortListAction = {
        param($sender, $e)
        $source = $e.OriginalSource
        while ($source -and -not ($source -is [System.Windows.Controls.GridViewColumnHeader])) {
            if ($source -is [System.Windows.FrameworkElement]) { $source = $source.Parent } else { break }
        }
        if ($source -and ($source -is [System.Windows.Controls.GridViewColumnHeader]) -and $source.Role -ne "Padding") {
            $column = $source.Column
            $sortBy = if ($column.DisplayMemberBinding) { $column.DisplayMemberBinding.Path.Path } else { "Name" }
            if ($sortBy) {
                if ($State.LastSortCol -eq $sortBy) { $State.SortDescending = -not $State.SortDescending } 
                else { $State.SortDescending = $false; $State.LastSortCol = $sortBy }
                if ($lvData -and $lvData.ItemsSource) {
                    $items = @($lvData.ItemsSource) 
                    if ($items.Count -gt 0) {
                         $sorted = $items | Sort-Object -Property $sortBy -Descending:$State.SortDescending
                         $lvData.ItemsSource = @($sorted)
                         if ($lblStatus) { $lblStatus.Text = "Sorted by $sortBy ($if ($State.SortDescending) { 'Descending' } else { 'Ascending' })" }
                    }
                }
            }
        }
    }.GetNewClosure()

    $PerformSearch = {
        if (-not $txtSearch) { return }
        $term = $txtSearch.Text
        if ([string]::IsNullOrWhiteSpace($term)) {
            Add-AppLog -Event "Search" -Username "System" -Details "Please enter a search term." -Config $Config -State $State -Color "Orange" -Status "Warning"
            return
        }
        
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        Add-AppLog -Event "Search" -Username "System" -Details "Searching Directory for '$term'..." -Config $Config -State $State -Status "Info"
        
        $userResults = Search-ADUsers -SearchTerm $term -Config $Config
        $compResults = Search-ADComputers -SearchTerm $term -Config $Config
        
        if ($compResults -and $compResults.Count -gt 0) {
            $compNames = @($compResults | Select-Object -ExpandProperty Name | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($compNames.Count -gt 0) {
                $job = Invoke-Command -ComputerName $compNames -ScriptBlock { 
                    try { (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Description } catch { "WMI Error" }
                } -AsJob -ErrorAction SilentlyContinue
                
                if ($job) {
                    $timeoutCount = 40
                    while ($job.State -eq 'Running' -and $timeoutCount -gt 0) { Start-Sleep -Milliseconds 100; $timeoutCount-- }
                    $wmiResults = Receive-Job $job -ErrorAction SilentlyContinue
                    Stop-Job $job -ErrorAction SilentlyContinue
                    Remove-Job $job -Force -ErrorAction SilentlyContinue
                    
                    foreach ($comp in $compResults) {
                        if (-not $comp.psobject.Properties['LockedOut']) { $comp | Add-Member -MemberType NoteProperty -Name "LockedOut" -Value "" -Force }
                        $match = @($wmiResults | Where-Object { $_.PSComputerName -eq $comp.Name })
                        if ($match.Count -gt 0) {
                            $descVal = $match[0]
                            if ([string]::IsNullOrWhiteSpace($descVal)) { $comp.LockedOut = "<Blank>" } else { $comp.LockedOut = [string]$descVal }
                        } else { $comp.LockedOut = "Unreachable" }
                    }
                }
            }
        }
        
        $allResults = @($userResults) + @($compResults)
        if ($lvData) { $lvData.ItemsSource = $allResults }
        if ($lblStatus) { $lblStatus.Text = "Found $($allResults.Count) objects matching '$term'. (Auto-refresh paused 10m)" }
        
        $State.IsSearchPaused = $true
        $State.RefreshTargetTime = (Get-Date).AddMinutes(10)
        if ($State.Timer -and $State.Timer.Interval.TotalSeconds -ne 1) {
            $State.Timer.Interval = [TimeSpan]::FromSeconds(1)
            if (-not $State.Timer.IsEnabled) { $State.Timer.Start() }
        }
        [System.Windows.Input.Mouse]::OverrideCursor = $null
    }.GetNewClosure()

    $RefreshAction = {
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        if ($State.IsSearchPaused) {
            $State.IsSearchPaused = $false
            if ($State.Timer -and -not $State.Timer.IsEnabled) { $State.Timer.Start() }
        }
        $State.RefreshTargetTime = (Get-Date).AddSeconds($State.RefreshIntervalSeconds)
        if ($btnRefresh) { $btnRefresh.Content = "Refresh List ($($State.RefreshIntervalSeconds)s)" }

        $locked = Get-LockedADUsers -Config $Config
        $safeLocked = @($locked)
        
        $isFirstRun = ($null -eq $State.LoggedLockoutTimes)
        if ($isFirstRun) { $State.LoggedLockoutTimes = @{} }
        
        $logDir = $Config.GeneralSettings.LogDirectoryUNC
        $today = Get-Date -Format "yyyyMMdd"
        $logFile = Join-Path $logDir "UnlockLog_$today.csv"
        $recentLogs = $null
        $logsLoaded = $false
        
        $currentLockedUsers = @()
        if ($safeLocked.Count -gt 0) {
            foreach ($u in $safeLocked) {
                $currentLockedUsers += $u.Name
                $newLockoutTime = if ($u.LockoutTime) { "$($u.LockoutTime)" } else { "Unknown" }
                
                if (-not $State.LoggedLockoutTimes.ContainsKey($u.Name) -or $State.LoggedLockoutTimes[$u.Name] -ne $newLockoutTime) {
                    if (-not [string]::IsNullOrWhiteSpace($newLockoutTime) -and $newLockoutTime -ne "Unknown") {
                        $isDuplicate = $false
                        if (-not $logsLoaded) {
                            if (Test-Path -LiteralPath $logFile) { try { $recentLogs = @(Import-Csv -LiteralPath $logFile -ErrorAction Stop | Select-Object -Last 100) } catch { $recentLogs = @() } } else { $recentLogs = @() }
                            $logsLoaded = $true
                        }
                        if ($recentLogs) {
                            $dup = $recentLogs | Where-Object { $_.Event -eq "Lockout Detected" -and $_.Username -eq $u.Name }
                            if ($dup) {
                                foreach ($d in $dup) {
                                    if ($d.Details -match "Time:\s*(.*)\)") {
                                        try {
                                            $lTime = [datetime]$matches[1]; $nTime = [datetime]$newLockoutTime
                                            if ([math]::Abs(($nTime - $lTime).TotalMinutes) -lt 5) { $isDuplicate = $true; break }
                                        } catch { if ($d.Details -match [regex]::Escape($newLockoutTime)) { $isDuplicate = $true; break } }
                                    } elseif ($d.Details -match [regex]::Escape($newLockoutTime)) { $isDuplicate = $true; break }
                                }
                            }
                        }
                        $State.LoggedLockoutTimes[$u.Name] = $newLockoutTime
                        if (-not $isDuplicate) {
                            $detailsMsg = "Account lockout detected. (AD Time: $newLockoutTime)"
                            Add-AppLog -Event "Lockout Detected" -Username $u.Name -Details $detailsMsg -Config $Config -State $State -Status "Warning" -Color "Orange"
                            if ($null -ne $recentLogs) { $recentLogs += [PSCustomObject]@{ Event = "Lockout Detected"; Username = $u.Name; Details = $detailsMsg; Timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss") } }
                        }
                    }
                }
            }
        }

        if (-not $isFirstRun) {
            $clearedUsers = @()
            foreach ($trackedUser in $State.LoggedLockoutTimes.Keys) {
                if ($trackedUser -notin $currentLockedUsers) { $clearedUsers += $trackedUser }
            }
            foreach ($cleared in $clearedUsers) {
                $isDuplicate = $false
                if (-not $logsLoaded) {
                    if (Test-Path -LiteralPath $logFile) { try { $recentLogs = @(Import-Csv -LiteralPath $logFile -ErrorAction Stop | Select-Object -Last 100) } catch { $recentLogs = @() } } else { $recentLogs = @() }
                    $logsLoaded = $true
                }
                if ($recentLogs) {
                    $dup = $recentLogs | Where-Object { $_.Event -eq "Lockout Cleared" -and $_.Username -eq $cleared }
                    foreach ($d in $dup) { try { if ([math]::Abs(((Get-Date) - [datetime]$d.Timestamp).TotalMinutes) -lt 5) { $isDuplicate = $true; break } } catch {} }
                }
                if (-not $isDuplicate) {
                    Add-AppLog -Event "Lockout Cleared" -Username $cleared -Details "Account is no longer locked." -Config $Config -State $State -Status "Info" -Color "Green"
                    if ($null -ne $recentLogs) { $recentLogs += [PSCustomObject]@{ Event = "Lockout Cleared"; Username = $cleared; Timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss") } }
                }
                $State.LoggedLockoutTimes.Remove($cleared)
            }
        }

        if ($lvData) { $lvData.ItemsSource = $safeLocked }
        if ($lblStatus) { $lblStatus.Text = "Found $($safeLocked.Count) locked accounts. (Last Update: $(Get-Date -Format 'HH:mm:ss'))" }
        [System.Windows.Input.Mouse]::OverrideCursor = $null
    }.GetNewClosure()

    # Expose actions globally to the other UI modules
    $State.Actions.RefreshData = $RefreshAction

    & $UpdateGridColumns
    if ($lvData) {
        $lvData.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$SortListAction)
        $lvData.AddHandler([System.Windows.Controls.Control]::ContextMenuOpeningEvent, [System.Windows.Controls.ContextMenuEventHandler]{
            $sel = $lvData.SelectedItem
            if ($ctxDetails) { $ctxDetails.Visibility = if ($sel) { "Visible" } else { "Collapsed" } }
        }.GetNewClosure())
    }
    
    if ($btnThemeToggle) { $btnThemeToggle.Add_Click({ $newTheme = if ($State.CurrentTheme -eq "Light") { "Dark" } else { "Light" }; & $ApplyTheme -TargetTheme $newTheme }.GetNewClosure()) }
    if ($btnSearch) { $btnSearch.Add_Click($PerformSearch) }
    if ($txtSearch) { $txtSearch.Add_KeyDown({ param($sender, $e) if ($e.Key -eq 'Enter' -and -not [string]::IsNullOrWhiteSpace($txtSearch.Text)) { & $PerformSearch } }.GetNewClosure()) }
    if ($btnRefresh) { $btnRefresh.Add_Click($RefreshAction) }

    if ($chkEnableEmail) {
        $chkEnableEmail.Add_Checked({ $Config.EmailSettings.EnableEmailNotifications = $true }.GetNewClosure())
        $chkEnableEmail.Add_Unchecked({ $Config.EmailSettings.EnableEmailNotifications = $false }.GetNewClosure())
    }

    if ($btnViewLog) {
        $btnViewLog.Add_Click({
            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
            $logs = Get-AppLogFiles -Config $Config
            $colors = Get-FluentThemeColors $State
            $logWin = Load-XamlWindow -XamlPath (Join-Path $AppRoot "UI\Windows\LogViewer.xaml") -ThemeColors $colors
            $logWin.Owner = $Window
            
            $lvLogs = $logWin.FindName("lvLogs")
            $txtOp = $logWin.FindName("txtFilterOperator")
            $txtUsr = $logWin.FindName("txtFilterUser")
            $dpStart = $logWin.FindName("dpStartDate")
            $dpEnd = $logWin.FindName("dpEndDate")
            
            $ApplyFilter = {
                $filtered = $logs | Where-Object {
                    $pass = $true
                    if ($txtOp.Text) { $pass = $pass -and ($_.Operator -match $txtOp.Text) }
                    if ($txtUsr.Text) { $pass = $pass -and ($_.Username -match $txtUsr.Text) }
                    if ($_.Timestamp -and $dpStart.SelectedDate) { try { if ([DateTime]$_.Timestamp -lt $dpStart.SelectedDate) { $pass = $false } } catch {} }
                    if ($_.Timestamp -and $dpEnd.SelectedDate) { try { if ([DateTime]$_.Timestamp -gt $dpEnd.SelectedDate.AddDays(1)) { $pass = $false } } catch {} }
                    return $pass
                }
                if ($lvLogs) { $lvLogs.ItemsSource = @($filtered | Sort-Object Timestamp -Descending) }
            }.GetNewClosure()
            
            & $ApplyFilter
            
            $btnFilter = $logWin.FindName("btnFilter"); if ($btnFilter) { $btnFilter.Add_Click($ApplyFilter) }
            $btnCloseLog = $logWin.FindName("btnCloseLog"); if ($btnCloseLog) { $btnCloseLog.Add_Click({ $logWin.Close() }.GetNewClosure()) }
            $btnExport = $logWin.FindName("btnExport")
            if ($btnExport) {
                $btnExport.Add_Click({
                    $sfd = New-Object Microsoft.Win32.SaveFileDialog; $sfd.Filter = "CSV (*.csv)|*.csv"; $sfd.FileName = "Export_Logs.csv"
                    if ($sfd.ShowDialog() -eq $true -and $lvLogs) { 
                        $lvLogs.ItemsSource | Export-Csv -Path $sfd.FileName -NoTypeInformation
                        Show-AppMessageBox -Message "Exported." -Title "Success" -ThemeColors $colors
                    }
                }.GetNewClosure())
            }
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            $logWin.Show()
        }.GetNewClosure())
    }

    $ShowDetailsAction = {
        if ($lvData.SelectedItem) {
            $det = Get-UserDetails -Identity $lvData.SelectedItem.Name -Type $lvData.SelectedItem.Type -Config $Config
            if ($det) {
                if ($txtDetailsContent) { 
                    $displayObj = [ordered]@{}
                    $dateProps = @("accountexpires", "badpasswordtime", "lastlogon", "lastlogontimestamp", "lockouttime", "pwdlastset")
                    $visibleProps = $det.PSObject.Properties | Where-Object { $_.Name -notmatch "^(PropertyNames|AddedProperties|RemovedProperties|ModifiedProperties|ClearProperties|SessionInfo)$" -and $_.Name -notmatch "Certificate" } | Sort-Object Name
                    foreach ($p in $visibleProps) {
                        $val = $p.Value
                        if ($null -ne $val -and $p.Name.ToLower() -in $dateProps) {
                            $valInt = $val -as [Int64]
                            if ($null -ne $valInt) {
                                if ($valInt -eq 0 -or $valInt -ge 9223372036854770000) { $val = "Never" } 
                                else { try { $val = [datetime]::FromFileTime($valInt).ToString("MM/dd/yyyy h:mm:ss tt") } catch {} }
                            }
                        }
                        $displayObj[$p.Name] = $val
                    }
                    $txtDetailsContent.Text = ([PSCustomObject]$displayObj | Format-List | Out-String)
                }
                if ($overlayDetails) { $overlayDetails.Visibility = "Visible" }
            }
        }
    }.GetNewClosure()

    if ($ctxDetails) { $ctxDetails.Add_Click($ShowDetailsAction) }
    if ($lvData) { $lvData.Add_MouseDoubleClick($ShowDetailsAction) }
    if ($btnCloseDetails) { $btnCloseDetails.Add_Click({ if ($overlayDetails) { $overlayDetails.Visibility = "Collapsed" } }.GetNewClosure()) }
    if ($thumbResizeDetails) {
        $thumbResizeDetails.Add_DragDelta({ param($sender, $e)
            $newWidth = $borderDetails.Width + $e.HorizontalChange; $newHeight = $borderDetails.Height + $e.VerticalChange
            if ($newWidth -gt 300) { $borderDetails.Width = $newWidth }
            if ($newHeight -gt 200) { $borderDetails.Height = $newHeight }
        }.GetNewClosure())
    }

    if ($btnDashboard) { $btnDashboard.Add_Click({ Start-Process "msedge.exe" -ArgumentList "--app=""https://vm-simplify/simplifyit/custom/fileuploads/acctdashboard.html""" }.GetNewClosure()) }
    if ($btnHelp) { $btnHelp.Add_Click({ 
        $helpPath = Join-Path $AppRoot "AccountMonitoringDocumentation.html"
        if (Test-Path $helpPath) { Start-Process "msedge.exe" -ArgumentList "--app=""$helpPath""" } else { Show-AppMessageBox -Message "Help document not found." -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) | Out-Null }
    }.GetNewClosure()) }
    if ($btnTechDocs) { $btnTechDocs.Add_Click({ 
        $techPath = Join-Path $AppRoot "AccountMonitorTechDoc.html"
        if (Test-Path $techPath) { Start-Process "msedge.exe" -ArgumentList "--app=""$techPath""" } else { Show-AppMessageBox -Message "Tech document not found." -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) | Out-Null }
    }.GetNewClosure()) }
    if ($btnModifyConfig) {
        $btnModifyConfig.Add_Click({
            $editorPath = Join-Path $AppRoot "ConfigEditor.html"
            if (Test-Path $editorPath) {
                $rawHtml = Get-Content -Path $editorPath -Raw
                $jsonStr = $Config | ConvertTo-Json -Depth 10 -Compress
                $rawHtml = $rawHtml -replace 'window\.INJECTED_CONFIG\s*=\s*null;', "window.INJECTED_CONFIG = $jsonStr;"
                $tempPath = Join-Path $env:TEMP "HDCompanion_ConfigEditor.html"
                Set-Content -Path $tempPath -Value $rawHtml -Force
                Start-Process "msedge.exe" -ArgumentList "--app=""$tempPath"""
                $res = Show-AppMessageBox -Message "Configuration Editor launched in Edge.`n`nPlease save your changes to 'AcctMonitorCfg.json' in the script directory.`n`nClick OK to reload settings now." -Title "Edit Configuration" -ButtonType "OKCancel" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
                if ($res -eq "OK") {
                    $newConfig = Get-AppConfig
                    $Config.PSObject.Properties | ForEach-Object { $Config.($_.Name) = $newConfig.($_.Name) }
                    & $ApplyTheme -TargetTheme $State.CurrentTheme
                    if ($Config.ControlProperties) {
                        $lblTitle = $Window.FindName("lblTitle"); if ($lblTitle) { $lblTitle.Text = $Config.ControlProperties.TitleLabel.Text }
                        $lblSubtitle = $Window.FindName("lblSubtitle"); if ($lblSubtitle) { $lblSubtitle.Text = $Config.ControlProperties.SubtitleLabel.Text }
                        if ($btnRefresh) { $btnRefresh.Content = $Config.ControlProperties.RefreshButton.Text }
                        $btnUnlock = $Window.FindName("btnUnlock"); if ($btnUnlock) { $btnUnlock.Content = $Config.ControlProperties.UnlockButton.Text }
                        $btnUnlockAll = $Window.FindName("btnUnlockAll"); if ($btnUnlockAll) { $btnUnlockAll.Content = $Config.ControlProperties.UnlockAllButton.Text }
                        if ($btnSearch) { $btnSearch.Content = $Config.ControlProperties.SearchButton.Text }
                        if ($btnViewLog) { $btnViewLog.Content = $Config.ControlProperties.ViewLogButton.Text }
                    }
                    if ($chkEnableEmail) { $chkEnableEmail.IsChecked = $Config.EmailSettings.EnableEmailNotifications }
                    $cfgPath = if ($newConfig.LoadedConfigPath) { $newConfig.LoadedConfigPath } else { "Unknown" }
                    Add-AppLog -Event "System" -Username "System" -Details "Configuration reloaded from: $cfgPath" -Config $Config -State $State -Status "Info" -Color "Blue"
                }
            } else { Show-AppMessageBox -Message "ConfigEditor.html not found." -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
        }.GetNewClosure())
    }

    $State.Timer = [System.Windows.Threading.DispatcherTimer]::new()
    $State.Timer.Add_Tick({ 
        $overlayReset = $Window.FindName("overlayReset")
        if ($overlayReset -and $overlayReset.Visibility -eq "Visible" -or ($overlayDetails -and $overlayDetails.Visibility -eq "Visible")) { return }
        $diff = $State.RefreshTargetTime - (Get-Date)
        if ($diff.TotalSeconds -le 0) { & $RefreshAction } 
        elseif ($btnRefresh) { $btnRefresh.Content = "Refresh List ({0}s)" -f [math]::Ceiling($diff.TotalSeconds) }
    }.GetNewClosure())
    
    if ($cbAutoRefresh) {
        $cbAutoRefresh.Add_SelectionChanged({
            $sec = 180; switch ($cbAutoRefresh.SelectedItem) { "30 seconds"{$sec=30}; "1 minute"{$sec=60}; "2 minutes"{$sec=120}; "5 minutes"{$sec=300}; "10 minutes"{$sec=600}; "15 minutes"{$sec=900}; "30 minutes"{$sec=1800} }
            $State.RefreshIntervalSeconds = $sec; $State.RefreshTargetTime = (Get-Date).AddSeconds($sec)
        }.GetNewClosure())
    }

    $Window.Add_Loaded({
        & $ApplyTheme -TargetTheme "Light"
        if ($btnRefresh) { $btnRefresh.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }
        $State.Timer.Interval = [TimeSpan]::FromSeconds(1); $State.Timer.Start()
        
        $cfgPath = if ($Config.LoadedConfigPath) { $Config.LoadedConfigPath } else { "Unknown" }
        Add-AppLog -Event "Config" -Username "System" -Details "Using configuration from: $cfgPath" -Config $Config -State $State -Status "Info" -Color "Blue"
    }.GetNewClosure())
}

Export-ModuleMember -Function *