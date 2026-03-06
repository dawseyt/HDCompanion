# ============================================================================
# UI.Freshservice.psm1 - Freshservice Integration Interface Logic
# ============================================================================

function Submit-FSQuickTicket {
    param($RequesterEmail, $Subject, $Description, $Status, $Config, $Category, $SubCategory, $ItemCategory, $CustomFields, $AssigneeEmail)
    $statusMap = @{ "Open"=2; "Pending"=3; "Resolved"=4; "Closed"=5 }
    $statusId = if ($statusMap.ContainsKey($Status)) { $statusMap[$Status] } else { 4 }
    $token = $null
    if ($null -ne $Config.FreshserviceSettings -and $null -ne $Config.FreshserviceSettings.Token) { $token = $Config.FreshserviceSettings.Token }
    elseif ($null -ne $Config.FreshserviceSettings -and $null -ne $Config.FreshserviceSettings.ApiKey) { $token = $Config.FreshserviceSettings.ApiKey }
    elseif ($null -ne $Config.ApiKeys -and $null -ne $Config.ApiKeys.FreshserviceToken) { $token = $Config.ApiKeys.FreshserviceToken }
    elseif ($null -ne $Config.ApiKeys -and $null -ne $Config.ApiKeys.FreshserviceApiKey) { $token = $Config.ApiKeys.FreshserviceApiKey }
    elseif ($null -ne $Config.GeneralSettings -and $null -ne $Config.GeneralSettings.FreshserviceApiKey) { $token = $Config.GeneralSettings.FreshserviceApiKey }
    elseif ($null -ne $Config.GeneralSettings -and $null -ne $Config.GeneralSettings.FreshserviceToken) { $token = $Config.GeneralSettings.FreshserviceToken }
    $url = $null
    if ($null -ne $Config.FreshserviceSettings -and $null -ne $Config.FreshserviceSettings.Url) { $url = $Config.FreshserviceSettings.Url }
    elseif ($null -ne $Config.ApiUrls -and $null -ne $Config.ApiUrls.FreshserviceUrl) { $url = $Config.ApiUrls.FreshserviceUrl }
    elseif ($null -ne $Config.GeneralSettings -and $null -ne $Config.GeneralSettings.FreshserviceDomain) { $url = $Config.GeneralSettings.FreshserviceDomain }
    elseif ($null -ne $Config.GeneralSettings -and $null -ne $Config.GeneralSettings.FreshserviceUrl) { $url = $Config.GeneralSettings.FreshserviceUrl }
    if (-not $token -or -not $url) { throw "Freshservice API Token or URL is missing from configuration." }
    $url = $url.TrimEnd('/')
    $encodedCreds = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($token):X"))
    $headers = @{ "Authorization" = "Basic $encodedCreds"; "Content-Type" = "application/json" }
    
    $responderId = $null
    if (-not [string]::IsNullOrWhiteSpace($AssigneeEmail) -and $AssigneeEmail -ne "Unassigned") {
        try {
            $agentRes = Invoke-RestMethod -Uri "$url/api/v2/agents?email=$AssigneeEmail" -Method Get -Headers $headers -ErrorAction Stop
            if ($agentRes.agents -and $agentRes.agents.Count -gt 0) { $responderId = $agentRes.agents[0].id }
        } catch { Write-Warning "Could not resolve Assignee email to Agent ID." }
    }
    $bodyHash = @{ description = $Description; subject = $Subject; status = $statusId; priority = 1; source = 2; email = $RequesterEmail; category = $Category; sub_category = $SubCategory; item_category = $ItemCategory; custom_fields = $CustomFields }
    if ($responderId) { $bodyHash.Add("responder_id", [long]$responderId) }
    $body = $bodyHash | ConvertTo-Json -Depth 5
    
    try {
        $res = Invoke-RestMethod -Uri "$url/api/v2/tickets" -Method Post -Headers $headers -Body $body -ErrorAction Stop
        return [string]$res.ticket.id
    } catch {
        $errorMsg = $_.Exception.Message
        if ($_.Exception.Response) {
            try {
                $reader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                $rawResponse = $reader.ReadToEnd(); $json = $rawResponse | ConvertFrom-Json
                if ($json.description) { $errorMsg += "`n`nFS Reason: $($json.description)" }
                if ($json.errors) { foreach ($err in $json.errors) { $errorMsg += "`n- $($err.field): $($err.message)" } }
            } catch { }
        }
        throw $errorMsg
    }
}

function Register-FreshserviceUIEvents {
    param($Window, $Config, $State)

    $lvData = $Window.FindName("lvData")
    $ctxFSInventory = $Window.FindName("ctxFSInventory")
    $ctxFindComputer = $Window.FindName("ctxFindComputer")
    $ctxOpenFSRecord = $Window.FindName("ctxOpenFSRecord")
    
    # Inject Dynamic Context Menu Item
    $ctxQuickTicket = New-Object System.Windows.Controls.MenuItem
    $ctxQuickTicket.Header = "Log Quick Ticket"
    $ctxQuickTicket.Visibility = "Collapsed"
    if ($lvData -and $lvData.ContextMenu) {
        $lvData.ContextMenu.Items.Insert(3, $ctxQuickTicket)
    }

    if ($lvData -and $lvData.ContextMenu) {
        $lvData.AddHandler([System.Windows.Controls.Control]::ContextMenuOpeningEvent, [System.Windows.Controls.ContextMenuEventHandler]{
            $sel = $lvData.SelectedItem
            $isUser = ($sel -and $sel.Type -eq "User")
            $isComp = ($sel -and $sel.Type -eq "Computer")
            
            if ($ctxQuickTicket) { $ctxQuickTicket.Visibility = if ($isUser) { "Visible" } else { "Collapsed" } }
            if ($ctxOpenFSRecord) { $ctxOpenFSRecord.Visibility = if ($isUser) { "Visible" } else { "Collapsed" } }
            if ($ctxFindComputer) { $ctxFindComputer.Visibility = if ($isUser) { "Visible" } else { "Collapsed" } }
            if ($ctxFSInventory) { $ctxFSInventory.Visibility = if ($isComp) { "Visible" } else { "Collapsed" } }
        }.GetNewClosure())
    }

    if ($ctxQuickTicket) {
        $ctxQuickTicket.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "User") {
                $userObj = $lvData.SelectedItem; $colors = Get-FluentThemeColors $State
                $reqEmail = if ($userObj.Email) { $userObj.Email } else { "$($userObj.Name)@pelicancu.com" }
                
                $tktXaml = @"
                <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Create Quick Ticket" Width="550" Height="700" WindowStartupLocation="CenterOwner" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                    <Window.Resources>
                        <Style TargetType="TextBox"><Setter Property="Background" Value="{Theme_BtnBg}"/><Setter Property="Foreground" Value="{Theme_Fg}"/><Setter Property="BorderBrush" Value="{Theme_BtnBorder}"/><Setter Property="BorderThickness" Value="1"/><Setter Property="Padding" Value="6,4"/></Style>
                        <Style TargetType="ComboBox"><Setter Property="Background" Value="{Theme_BtnBg}"/><Setter Property="Foreground" Value="{Theme_Fg}"/><Setter Property="BorderBrush" Value="{Theme_BtnBorder}"/><Setter Property="BorderThickness" Value="1"/></Style>
                        <Style TargetType="Button">
                            <Setter Property="Template">
                                <Setter.Value><ControlTemplate TargetType="Button"><Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value>
                            </Setter>
                        </Style>
                    </Window.Resources>
                    <Border Background="{Theme_Bg}" CornerRadius="8" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" Margin="15">
                        <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3"/></Border.Effect>
                        <Grid>
                            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                            <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="16,16,16,8" Cursor="Hand"><TextBlock Text="Log Quick Ticket" FontSize="16" FontWeight="SemiBold" Foreground="{Theme_Fg}"/></Border>
                            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                                <StackPanel Margin="16,8,16,16">
                                    <TextBlock Text="Requester Email:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/>
                                    <TextBox Text="$reqEmail" IsReadOnly="True" Margin="0,0,0,12" Background="Transparent" BorderThickness="0" FontWeight="Bold"/>
                                    <TextBlock Text="Subject:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/>
                                    <TextBox x:Name="txtSubject" Text="Unlock Account: $reqEmail" Margin="0,0,0,12"/>
                                    <TextBlock Text="Notes / Description:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/>
                                    <TextBox x:Name="txtDesc" Text="The Microsoft account for $reqEmail has been locked due to too many bad password attempts." AcceptsReturn="True" TextWrapping="Wrap" Height="60" Margin="0,0,0,12"/>
                                    <TextBlock Text="Ticket Status:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/>
                                    <ComboBox x:Name="cbStatus" Height="30" SelectedIndex="0" Margin="0,0,0,16"><ComboBoxItem Content="Open"/><ComboBoxItem Content="Pending"/><ComboBoxItem Content="Resolved"/><ComboBoxItem Content="Closed"/></ComboBox>
                                    
                                    <TextBlock Text="Required Ticket Properties:" FontWeight="Bold" Foreground="{Theme_Fg}" Margin="0,0,0,8"/>
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="15"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                                        
                                        <TextBlock Grid.Row="0" Grid.Column="0" Text="Category:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbCategory" Grid.Row="1" Grid.Column="0" Height="28" Margin="0,0,0,12"/>
                                        <TextBlock Grid.Row="0" Grid.Column="2" Text="Sub-Category:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><TextBox x:Name="txtSubCategory" Grid.Row="1" Grid.Column="2" Height="28" Text="Windows" Margin="0,0,0,12"/>
                                        <TextBlock Grid.Row="2" Grid.Column="0" Text="Item:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><TextBox x:Name="txtItemCategory" Grid.Row="3" Grid.Column="0" Height="28" Text="Account Unlock" Margin="0,0,0,12"/>
                                        <TextBlock Grid.Row="2" Grid.Column="2" Text="Resolved Remotely:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbResolvedRemotely" Grid.Row="3" Grid.Column="2" Height="28" Margin="0,0,0,12"><ComboBoxItem Content="Yes" IsSelected="True"/><ComboBoxItem Content="No"/></ComboBox>
                                        <TextBlock Grid.Row="4" Grid.Column="0" Text="Member Impacting:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbMemberImpacting" Grid.Row="5" Grid.Column="0" Height="28" Margin="0,0,0,12"><ComboBoxItem Content="Yes"/><ComboBoxItem Content="No" IsSelected="True"/></ComboBox>
                                        <TextBlock Grid.Row="4" Grid.Column="2" Text="Who is Affected:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbWhoAffected" Grid.Row="5" Grid.Column="2" Height="28" Margin="0,0,0,12"><ComboBoxItem Content="You Only" IsSelected="True"/><ComboBoxItem Content="Entire Department/Branch"/><ComboBoxItem Content="Company Wide"/></ComboBox>
                                        <TextBlock Grid.Row="6" Grid.Column="0" Text="Prevents Crit. Operations:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbPreventsOps" Grid.Row="7" Grid.Column="0" Height="28" Margin="0,0,0,4"><ComboBoxItem Content="Yes"/><ComboBoxItem Content="No" IsSelected="True"/></ComboBox>
                                        <TextBlock Grid.Row="6" Grid.Column="2" Text="Is there a Workaround:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbWorkaround" Grid.Row="7" Grid.Column="2" Height="28" Margin="0,0,0,4"><ComboBoxItem Content="Yes" IsSelected="True"/><ComboBoxItem Content="No"/></ComboBox>
                                        <TextBlock Grid.Row="8" Grid.Column="0" Text="Assignee:" FontSize="11" Foreground="{Theme_SecFg}" Margin="0,0,0,4"/><ComboBox x:Name="cbAssignee" Grid.Row="9" Grid.Column="0" Height="28" Margin="0,0,0,4" IsEditable="True"/>
                                    </Grid>
                                </StackPanel>
                            </ScrollViewer>
                            <Border Grid.Row="2" Background="{Theme_BtnBg}" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="{Theme_BtnBorder}">
                                <Grid>
                                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                        <Button x:Name="btnSubTkt" Content="Submit Ticket" Width="100" Height="28" Margin="0,0,8,0" Background="{Theme_PrimaryBg}" Foreground="{Theme_PrimaryFg}" BorderThickness="0" IsDefault="True"/>
                                        <Button x:Name="btnCancelTkt" Content="Cancel" Width="80" Height="28" Background="{Theme_Bg}" Foreground="{Theme_Fg}" BorderBrush="{Theme_BtnBorder}" BorderThickness="1" IsCancel="True"/>
                                    </StackPanel>
                                    <Thumb x:Name="thumbResizeTkt" HorizontalAlignment="Right" VerticalAlignment="Bottom" Width="12" Height="12" Cursor="SizeNWSE" Margin="0,0,-8,-8" Background="Transparent" ToolTip="Resize Window"/>
                                </Grid>
                            </Border>
                        </Grid>
                    </Border>
                </Window>
"@
                $xamlText = $tktXaml; foreach ($key in $colors.Keys) { $xamlText = $xamlText.Replace("{Theme_$key}", $colors[$key]) }
                $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlText))
                $tktWin = [System.Windows.Markup.XamlReader]::Load($reader)
                $tktWin.Owner = $Window
                
                $tktWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $tktWin.DragMove() }.GetNewClosure())
                $thumbResizeTkt = $tktWin.FindName("thumbResizeTkt")
                if ($thumbResizeTkt) {
                    $thumbResizeTkt.Add_DragDelta({ param($sender, $e)
                        $newWidth = $tktWin.Width + $e.HorizontalChange; $newHeight = $tktWin.Height + $e.VerticalChange
                        if ($newWidth -gt 400) { $tktWin.Width = $newWidth }
                        if ($newHeight -gt 350) { $tktWin.Height = $newHeight }
                    }.GetNewClosure())
                }

                $txtSub = $tktWin.FindName("txtSubject"); $txtDesc = $tktWin.FindName("txtDesc"); $cbStatus = $tktWin.FindName("cbStatus")
                $cbCategory = $tktWin.FindName("cbCategory")
                $cats = @('Access & Security','Accounting','Account Servicing','Alerts','Auditing','Building and Grounds Maintenance','Change Request','Card Services','Cyber Security Incident','Development & Reporting','Digital Banking','Documents','Email','Facilities','File & Folder','Genesys Cloud Change Management','Hardware','Human Resources','Microsoft Authenticator','Mobile Device Management','Morning Tasks','Network','Office Furniture','Purchasing','Quality Assurance (QA)','Relocations','Scheduled Maintenance','Software','Time Tracking','Training','User Accounts','VDI','Video Playback','WFH Application','Freshservice','Member Experience','Project Management','ITM/ATM')
                if ($cbCategory) { foreach ($c in $cats) { $cbCategory.Items.Add($c) | Out-Null }; $cbCategory.SelectedItem = 'User Accounts' }

                $cbAssignee = $tktWin.FindName("cbAssignee")
                if ($cbAssignee) {
                    $cbAssignee.Items.Add("Unassigned") | Out-Null
                    $emails = @()
                    if ($null -ne $Config.EmailSettings) {
                        foreach ($p in $Config.EmailSettings.PSObject.Properties) {
                            if ($p.Value -is [array]) { $emails += $p.Value }
                            elseif ($p.Value -is [string] -and $p.Value -match "@") { $emails += $p.Value -split ',' }
                        }
                    }
                    $emails = $emails | Select-Object -Unique | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "@" }
                    foreach ($em in $emails) { $cbAssignee.Items.Add($em) | Out-Null }
                    $cbAssignee.SelectedIndex = 0
                }

                $txtSubCategory = $tktWin.FindName("txtSubCategory"); $txtItemCategory = $tktWin.FindName("txtItemCategory")
                $cbResolvedRemotely = $tktWin.FindName("cbResolvedRemotely"); $cbMemberImpacting = $tktWin.FindName("cbMemberImpacting")
                $cbWhoAffected = $tktWin.FindName("cbWhoAffected"); $cbPreventsOps = $tktWin.FindName("cbPreventsOps"); $cbWorkaround = $tktWin.FindName("cbWorkaround")

                $btnCancel = $tktWin.FindName("btnCancelTkt"); if ($btnCancel) { $btnCancel.Add_Click({ $tktWin.Close() }.GetNewClosure()) }
                $btnSubmit = $tktWin.FindName("btnSubTkt")
                if ($btnSubmit) {
                    $btnSubmit.Add_Click({
                        $s = $txtSub.Text; $d = $txtDesc.Text; $st = $cbStatus.Text
                        $cat = if ($cbCategory) { $cbCategory.Text } else { "" }; $subCat = if ($txtSubCategory) { $txtSubCategory.Text } else { "" }; $itemCat = if ($txtItemCategory) { $txtItemCategory.Text } else { "" }; $assignee = if ($cbAssignee) { $cbAssignee.Text } else { "Unassigned" }
                        $cFields = @{ "resolved_remotely" = $cbResolvedRemotely.Text; "is_this_directly_member_impacting" = $cbMemberImpacting.Text; "who_or_what_is_affected_by_this_impact" = $cbWhoAffected.Text; "does_this_prevent_critical_business_operations_from_continuing" = $cbPreventsOps.Text; "is_there_a_workaround_for_this_issue" = $cbWorkaround.Text }
                        $tktWin.Close()
                        
                        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                        Add-AppLog -Event "Freshservice" -Username $userObj.Name -Details "Submitting quick ticket..." -Config $Config -State $State -Status "Info"
                        try {
                            $id = Submit-FSQuickTicket -RequesterEmail $reqEmail -Subject $s -Description $d -Status $st -Config $Config -Category $cat -SubCategory $subCat -ItemCategory $itemCat -CustomFields $cFields -AssigneeEmail $assignee
                            [System.Windows.Input.Mouse]::OverrideCursor = $null
                            Show-AppMessageBox -Message "Freshservice ticket successfully created!`n`nTicket Number: $id" -Title "Ticket Created" -IconType "Information" -OwnerWindow $Window -ThemeColors $colors
                        } catch {
                            [System.Windows.Input.Mouse]::OverrideCursor = $null
                            Show-AppMessageBox -Message "Failed to create ticket:`n`n$($_.Exception.Message)" -Title "API Error" -IconType "Error" -OwnerWindow $Window -ThemeColors $colors
                        }
                    }.GetNewClosure())
                }
                $tktWin.Show()
            }
        }.GetNewClosure())
    }

    if ($ctxOpenFSRecord) {
        $ctxOpenFSRecord.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "User") {
                $userObj = $lvData.SelectedItem
                [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                Add-AppLog -Event "Freshservice" -Username $userObj.Name -Details "Opening Freshservice record..." -Config $Config -State $State -Status "Info"
                try {
                    $record = Get-FSUserRecord -User $userObj -Config $Config
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    if ($record) { try { Start-Process "msedge.exe" -ArgumentList "--app=""$($record.Url)""" } catch { Show-AppMessageBox -Message "Failed to open Microsoft Edge." -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) } } 
                    else { Show-AppMessageBox -Message "No Requester or Agent record found for $($userObj.Name) in Freshservice." -Title "Not Found" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
                } catch {
                     [System.Windows.Input.Mouse]::OverrideCursor = $null
                     Show-AppMessageBox -Message "Error contacting Freshservice:`n`n$($_.Exception.Message)" -Title "API Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
                }
            }
        }.GetNewClosure())
    }

    if ($ctxFindComputer) {
        $ctxFindComputer.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "User") {
                $userObj = $lvData.SelectedItem
                [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                Add-AppLog -Event "Freshservice" -Username $userObj.Name -Details "Querying assigned computers..." -Config $Config -State $State -Status "Info"
                try {
                    $assets = Get-FSUserAsset -User $userObj -Config $Config
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    if ($assets -and $assets.Count -gt 0) {
                        $compList = ""; foreach ($a in $assets) { $compList += "- $($a.name)`n" }
                        Show-AppMessageBox -Message "Found $($assets.Count) computer(s) assigned to $($userObj.Name) in Freshservice:`n`n$compList" -Title "Assigned Computers" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
                    } else { Show-AppMessageBox -Message "No assets currently assigned to $($userObj.Name)." -Title "Not Found" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
                } catch { 
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    Show-AppMessageBox -Message "Error querying Freshservice:`n`n$($_.Exception.Message)" -Title "API Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) 
                }
            }
        }.GetNewClosure())
    }

    if ($ctxFSInventory) {
        $ctxFSInventory.Add_Click({
            if ($lvData.SelectedItem -and $lvData.SelectedItem.Type -eq "Computer") {
                $computerName = $lvData.SelectedItem.Name
                [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                Add-AppLog -Event "Freshservice" -Username "System" -Details "Fetching inventory for $computerName..." -Config $Config -State $State -Status "Info"
                try {
                    $asset = Get-FSAssetDetails -AssetName $computerName -Config $Config
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    if ($asset) {
                        if ($asset.display_id) {
                            $base = "https://helpdesk.pelicanstatecu.com/cmdb/items/"
                            $targetUri = [Uri]("$base$($asset.display_id)")
                            try { Start-Process "msedge.exe" -ArgumentList "--app=""$($targetUri.AbsoluteUri)""" } catch { Show-AppMessageBox -Message "Failed to open link." -Title "Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
                        } else { Show-AppMessageBox -Message "Asset found but missing Display ID." -Title "Warning" -IconType "Warning" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
                    } else { Show-AppMessageBox -Message "Computer '$computerName' not found in Freshservice inventory." -Title "Not Found" -IconType "Information" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State) }
                } catch {
                     [System.Windows.Input.Mouse]::OverrideCursor = $null
                     Show-AppMessageBox -Message "Error contacting Freshservice: $($_.Exception.Message)" -Title "API Error" -IconType "Error" -OwnerWindow $Window -ThemeColors (Get-FluentThemeColors $State)
                }
            }
        }.GetNewClosure())
    }
}
Export-ModuleMember -Function *