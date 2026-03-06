# ============================================================================
# PrinterManager.ps1 - Standalone Printer Management Thread
# ============================================================================
param(
    [Parameter(Mandatory=$false)]
    [string]$ComputerName = $env:COMPUTERNAME
)

# 1. Environment Setup
try {
    $global:PSScriptRoot = $PSScriptRoot
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
} catch { exit }

# Ensure 64-bit execution (Get-PrinterDriver often fails silently in 32-bit processes on 64-bit OS)
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $psExe = "$env:windir\sysnative\WindowsPowerShell\v1.0\powershell.exe"
    Start-Process $psExe -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`" -ComputerName `"$ComputerName`""
    exit
}

# 2. Import Required Backend Modules
Import-Module "$PSScriptRoot\Modules\CoreLogic.psm1" -Force -DisableNameChecking
Import-Module "$PSScriptRoot\Modules\RemoteManagement.psm1" -Force -DisableNameChecking

# 3. Setup Config and Inline Theme Colors
$Config = Get-AppConfig
$isDark = ($Config.GeneralSettings.DefaultTheme -eq "Dark")

# Define exact color map for inline XAML string replacement
$c = @{
    Bg         = if ($isDark) { "#2D2D30" } else { "#FFFFFF" }
    Fg         = if ($isDark) { "#FFFFFF" } else { "#202020" }
    SecFg      = if ($isDark) { "#CCCCCC" } else { "#5D5D5D" }
    BtnBg      = if ($isDark) { "#3E3E42" } else { "#F3F3F3" }
    BtnBorder  = if ($isDark) { "#555555" } else { "#CCCCCC" }
    PrimaryBg  = "#0078D4"
    PrimaryFg  = "#FFFFFF"
    GridBorder = if ($isDark) { "#444444" } else { "#E0E0E0" }
    HoverBg    = if ($isDark) { "#505055" } else { "#E5E5E5" }
    AltRowBg   = if ($isDark) { "#37373C" } else { "#F9F9F9" }
    Danger     = if ($isDark) { "#FF6666" } else { "#D13438" }
}

# 4. Universal Message Box (Embedded XAML)
function Show-LocalMessageBox {
    param([string]$Message, [string]$Title = "Information", [string]$ButtonType = "OK", [string]$IconType = "Information", $OwnerWindow)
    
    $iconData = ""
    $iconColor = $c.Fg
    switch ($IconType) {
        "Error" { $iconData = "M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"; $iconColor = "#D13438" }
        "Warning" { $iconData = "M12 2L1 21h22L12 2zm0 3.83L19.53 19H4.47L12 5.83zM11 10h2v5h-2v-5zm0 6h2v2h-2v-2z"; $iconColor = "#FFB900" }
        "Question" { $iconData = "M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 17h-2v-2h2v2zm2.07-7.75l-.9.92C13.45 12.9 13 13.5 13 15h-2v-.5c0-1.1.45-2.1 1.17-2.83l1.24-1.26c.37-.36.59-.86.59-1.41 0-1.1-.9-2-2-2s-2 .9-2 2H8c0-2.21 1.79-4 4-4s4 1.79 4 4c0 .88-.36 1.68-.93 2.25z"; $iconColor = "#0078D4" }
        Default { $iconData = "M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h-2v6h2zm0-8h-2V7h2v2z"; $iconColor = "#0078D4" }
    }

    $btnOkVis = if ($ButtonType -in @("OK", "OKCancel")) { "Visible" } else { "Collapsed" }
    $btnCancelVis = if ($ButtonType -eq "OKCancel") { "Visible" } else { "Collapsed" }
    $btnYesVis = if ($ButtonType -eq "YesNo") { "Visible" } else { "Collapsed" }
    $btnNoVis = if ($ButtonType -eq "YesNo") { "Visible" } else { "Collapsed" }

    $msgXaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="$Title" Width="450" SizeToContent="Height" MinHeight="180" 
            WindowStartupLocation="CenterOwner" ResizeMode="NoResize" WindowStyle="None" AllowsTransparency="True" Background="Transparent"
            FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
        <Window.Resources>
            <Style TargetType="Button">
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button">
                            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </Window.Resources>
        <Border Background="$($c.Bg)" CornerRadius="8" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Margin="15">
            <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3" Color="Black"/></Border.Effect>
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid x:Name="TitleBar" Grid.Row="0" Margin="20,20,20,0" Background="Transparent" Cursor="Hand">
                    <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                    <Path Grid.Column="0" Data="$iconData" Fill="$iconColor" Width="20" Height="20" Stretch="Uniform" VerticalAlignment="Center"/>
                    <TextBlock Grid.Column="1" Text="$Title" FontSize="16" FontWeight="SemiBold" Foreground="$($c.Fg)" VerticalAlignment="Center" Margin="12,0,0,0"/>
                </Grid>
                <TextBox x:Name="txtMessageBody" Grid.Row="1" Margin="52,12,20,20" IsReadOnly="True" Background="Transparent" BorderThickness="0" TextWrapping="Wrap" Foreground="$($c.SecFg)" FontSize="13" VerticalScrollBarVisibility="Auto" MaxHeight="300"/>
                <Border Grid.Row="2" Background="$($c.BtnBg)" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="$($c.BtnBorder)">
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="btnYes" Content="Yes" Width="80" Height="28" Margin="0,0,8,0" Visibility="$btnYesVis" Background="$($c.PrimaryBg)" Foreground="$($c.PrimaryFg)" BorderThickness="0"/>
                        <Button x:Name="btnNo" Content="No" Width="80" Height="28" Margin="0,0,0,0" Visibility="$btnNoVis" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                        <Button x:Name="btnOk" Content="OK" Width="80" Height="28" Margin="0,0,8,0" Visibility="$btnOkVis" Background="$($c.PrimaryBg)" Foreground="$($c.PrimaryFg)" BorderThickness="0"/>
                        <Button x:Name="btnCancel" Content="Cancel" Width="80" Height="28" Margin="0,0,0,0" Visibility="$btnCancelVis" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>
    </Window>
"@
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($msgXaml))
    $msgWin = [System.Windows.Markup.XamlReader]::Load($reader)
    
    $msgWin.FindName("txtMessageBody").Text = $Message
    $msgWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $msgWin.DragMove() })
    
    $script:msgRes = "Cancel"
    $msgWin.FindName("btnYes").Add_Click({ $script:msgRes = "Yes"; $msgWin.Close() })
    $msgWin.FindName("btnNo").Add_Click({ $script:msgRes = "No"; $msgWin.Close() })
    $msgWin.FindName("btnOk").Add_Click({ $script:msgRes = "OK"; $msgWin.Close() })
    $msgWin.FindName("btnCancel").Add_Click({ $script:msgRes = "Cancel"; $msgWin.Close() })
    
    if ($OwnerWindow -and $OwnerWindow.IsVisible) { $msgWin.Owner = $OwnerWindow; $msgWin.WindowStartupLocation = "CenterOwner" }
    else { $msgWin.WindowStartupLocation = "CenterScreen" }
    
    $msgWin.ShowDialog() | Out-Null
    return $script:msgRes
}

# 5. Load the Unified Printer Manager XAML (Embedded)
$pmXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="winPrinterManager" Title="Printer Management" Height="600" Width="950" WindowStartupLocation="CenterScreen" ResizeMode="CanResize" WindowStyle="SingleBorderWindow" Background="$($c.Bg)"
        FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Bd" Property="Opacity" Value="0.5"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background" Value="$($c.Bg)"/>
            <Setter Property="Foreground" Value="$($c.SecFg)"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="BorderBrush" Value="$($c.GridBorder)"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <Border Grid.Row="0" Background="$($c.Bg)" Padding="16" BorderBrush="$($c.GridBorder)" BorderThickness="0,0,0,1">
            <Grid>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock x:Name="lblHeaderTitle" Text="Printer Management: $ComputerName" FontWeight="SemiBold" FontSize="18" Foreground="$($c.Fg)"/>
                    <TextBlock x:Name="lblDescription" Text="Asset: Querying..." FontSize="13" Foreground="$($c.SecFg)" Margin="0,2,0,0"/>
                    <TextBlock x:Name="lblStatus" Text="Ready." FontSize="12" Foreground="$($c.SecFg)" Margin="0,4,0,0"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="btnRestartSpooler" Content="Restart Spooler" Margin="0,0,10,0" Width="120" Height="30" Background="$($c.BtnBg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                    <Button x:Name="btnInstallPrinter" Content="Install New Printer" Width="140" Height="30" Background="$($c.PrimaryBg)" Foreground="$($c.PrimaryFg)" BorderThickness="0"/>
                </StackPanel>
            </Grid>
        </Border>

        <ListView x:Name="lvPrinters" Grid.Row="1" Margin="8" BorderThickness="0" Background="Transparent" Foreground="$($c.Fg)" AlternationCount="2">
            <ListView.ItemContainerStyle>
                <Style TargetType="ListViewItem">
                    <Setter Property="Height" Value="28"/>
                    <Setter Property="Background" Value="Transparent"/>
                    <Setter Property="Foreground" Value="$($c.Fg)"/>
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="ListViewItem">
                                <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="Transparent" BorderThickness="0" CornerRadius="4" Padding="4,0" Margin="0,1">
                                    <GridViewRowPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsSelected" Value="true">
                                        <Setter TargetName="Bd" Property="Background" Value="$($c.PrimaryBg)"/>
                                        <Setter Property="Foreground" Value="$($c.PrimaryFg)"/>
                                    </Trigger>
                                    <Trigger Property="IsMouseOver" Value="true">
                                        <Setter TargetName="Bd" Property="Background" Value="$($c.HoverBg)"/>
                                    </Trigger>
                                    <MultiTrigger>
                                        <MultiTrigger.Conditions>
                                            <Condition Property="IsSelected" Value="true"/>
                                            <Condition Property="IsMouseOver" Value="true"/>
                                        </MultiTrigger.Conditions>
                                        <Setter TargetName="Bd" Property="Background" Value="$($c.PrimaryBg)"/>
                                        <Setter Property="Foreground" Value="$($c.PrimaryFg)"/>
                                    </MultiTrigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                    <Style.Triggers>
                        <Trigger Property="ItemsControl.AlternationIndex" Value="1">
                            <Setter Property="Background" Value="$($c.AltRowBg)"/>
                        </Trigger>
                    </Style.Triggers>
                </Style>
            </ListView.ItemContainerStyle>
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" Width="250"/>
                    <GridViewColumn Header="Driver" DisplayMemberBinding="{Binding DriverName}" Width="250"/>
                    <GridViewColumn Header="Port" DisplayMemberBinding="{Binding PortName}" Width="150"/>
                    <GridViewColumn Header="Shared" DisplayMemberBinding="{Binding Shared}" Width="60"/>
                    <GridViewColumn Header="Status" DisplayMemberBinding="{Binding PrinterStatus}" Width="100"/>
                </GridView>
            </ListView.View>
            <ListView.ContextMenu>
                <ContextMenu>
                    <MenuItem Header="View Details" x:Name="ctxPrnDetails"/>
                    <MenuItem Header="Rename Printer" x:Name="ctxPrnRename"/>
                    <MenuItem Header="Remove Printer" x:Name="ctxPrnRemove"/>
                </ContextMenu>
            </ListView.ContextMenu>
        </ListView>

        <Border Grid.Row="2" Background="$($c.BtnBg)" BorderBrush="$($c.GridBorder)" BorderThickness="0,1,0,0" Padding="12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                 <Button x:Name="btnRefresh" Content="Refresh List" Margin="0,0,10,0" Width="100" Height="28" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                 <Button x:Name="btnPrnDetails" Content="View Details" Margin="0,0,10,0" Width="100" Height="28" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                 <Button x:Name="btnPrnRemove" Content="Remove Printer" Margin="0,0,10,0" Width="110" Height="28" Background="$($c.Bg)" Foreground="$($c.Danger)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                 <Button x:Name="btnClose" Content="Close" Width="80" Height="28" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

$pmReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($pmXaml))
$prnWin = [System.Windows.Markup.XamlReader]::Load($pmReader)

# Bind Controls
$lvPrinters = $prnWin.FindName("lvPrinters")
$lblHeaderTitle = $prnWin.FindName("lblHeaderTitle")
$lblDescription = $prnWin.FindName("lblDescription")
$lblStatus = $prnWin.FindName("lblStatus")
$btnRestartSpooler = $prnWin.FindName("btnRestartSpooler")
$btnInstallPrinter = $prnWin.FindName("btnInstallPrinter")
$btnRefresh = $prnWin.FindName("btnRefresh")
$btnPrnDetails = $prnWin.FindName("btnPrnDetails")
$btnPrnRemove = $prnWin.FindName("btnPrnRemove")
$btnClose = $prnWin.FindName("btnClose")

$ctxPrnDetails = $prnWin.FindName("ctxPrnDetails")
$ctxPrnRename = $prnWin.FindName("ctxPrnRename")
$ctxPrnRemove = $prnWin.FindName("ctxPrnRemove")

# 6. Primary Action Logic (Closures Removed for Direct Binding Access)

$RefreshPrinters = {
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    $lblStatus.Text = "Querying printers from $ComputerName..."
    $lblStatus.Foreground = [System.Windows.Media.Brushes]::Orange
    
    $frame = New-Object System.Windows.Threading.DispatcherFrame
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Action]{ $frame.Continue = $false }) | Out-Null
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
    
    # Query Local System Properties Description Quietly
    if ($lblDescription -and $lblDescription.Text -eq "Asset: Querying...") {
        try {
            $desc = Invoke-Command -ComputerName $ComputerName -ScriptBlock { (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Description } -ErrorAction SilentlyContinue
            if ([string]::IsNullOrWhiteSpace($desc)) { $desc = "Not Set" }
            $lblDescription.Text = "Asset: $desc"
        } catch {
            $lblDescription.Text = "Asset: Unavailable"
        }
    }
    
    try {
        $printers = Get-RemotePrinters -ComputerName $ComputerName
        if ($lvPrinters) { $lvPrinters.ItemsSource = @($printers | Sort-Object Name) }
        $lblStatus.Text = "Found $($printers.Count) printers. (Updated: $(Get-Date -Format 'HH:mm:ss'))"
        $lblStatus.Foreground = [System.Windows.Media.Brushes]::Green
    } catch {
        $lblStatus.Text = "Error communicating with $ComputerName."
        $lblStatus.Foreground = [System.Windows.Media.Brushes]::Red
        Show-LocalMessageBox -Message "Failed to fetch printers:`n$($_.Exception.Message)" -Title "Connection Error" -IconType "Error" -OwnerWindow $prnWin | Out-Null
    }
    [System.Windows.Input.Mouse]::OverrideCursor = $null
}

$ShowPrinterDetails = {
    if ($lvPrinters -and $lvPrinters.SelectedItem) {
        $selP = $lvPrinters.SelectedItem
        
        $props = $selP.PSObject.Properties | Where-Object { $_.MemberType -match "Property" } | Sort-Object Name
        $text = ""
        foreach ($p in $props) { 
            $val = if ($null -ne $p.Value) { $p.Value.ToString() } else { "" }
            $text += "{0,-25} : {1}`r`n" -f $p.Name, $val
        }
        
        # Embedded Details XAML
        $detXaml = @"
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                Title="Printer Details" Height="400" Width="600" WindowStartupLocation="CenterOwner" ResizeMode="CanResize" WindowStyle="ToolWindow" Background="$($c.Bg)"
                FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
            <Window.Resources>
                <Style TargetType="Button">
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="Button">
                                <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>
            </Window.Resources>
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Border Grid.Row="0" Background="$($c.BtnBg)" Padding="12" BorderBrush="$($c.GridBorder)" BorderThickness="0,0,0,1">
                    <TextBlock Text="Property Listing for $($selP.Name)" FontWeight="SemiBold" Foreground="$($c.Fg)"/>
                </Border>
                <TextBox x:Name="txtPrnDetails" Grid.Row="1" IsReadOnly="True" FontFamily="Consolas" FontSize="12" VerticalScrollBarVisibility="Auto" BorderThickness="0" Margin="10" Background="Transparent" Foreground="$($c.Fg)"/>
                <Border Grid.Row="2" Background="$($c.BtnBg)" BorderBrush="$($c.GridBorder)" BorderThickness="0,1,0,0" Padding="12">
                    <Button x:Name="btnDetClose" Content="Close" HorizontalAlignment="Right" Width="80" Height="28" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                </Border>
            </Grid>
        </Window>
"@
        $dReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($detXaml))
        $detWin = [System.Windows.Markup.XamlReader]::Load($dReader)
        
        $detWin.Owner = $prnWin
        $detWin.FindName("txtPrnDetails").Text = $text
        $detWin.FindName("btnDetClose").Add_Click({ $detWin.Close() })
        
        $detWin.ShowDialog() | Out-Null
    }
}

$RemovePrinterAction = {
    if ($lvPrinters -and $lvPrinters.SelectedItem) {
        $pName = $lvPrinters.SelectedItem.Name
        $conf = Show-LocalMessageBox -Message "Are you sure you want to permanently remove the printer '$pName' from $ComputerName?" -Title "Confirm Removal" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $prnWin
        
        if ($conf -eq "Yes") {
            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
            try {
                Remove-RemotePrinter -ComputerName $ComputerName -PrinterName $pName
                Add-AppLog -Event "Printer Remove" -Username "System" -Details "Removed printer '$pName' from $ComputerName." -Config $Config -State $State -Status "Success"
                Show-LocalMessageBox -Message "Printer '$pName' removed successfully." -Title "Success" -OwnerWindow $prnWin | Out-Null
                & $RefreshPrinters
            } catch { 
                Show-LocalMessageBox -Message "Failed to remove printer:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $prnWin | Out-Null
                Add-AppLog -Event "Printer Remove" -Username "System" -Details "Failed to remove '$pName' on $($ComputerName): $($_.Exception.Message)" -Config $Config -State $State -Status "Error"
            }
            [System.Windows.Input.Mouse]::OverrideCursor = $null
        }
    }
}

$RenamePrinterAction = {
    if ($lvPrinters -and $lvPrinters.SelectedItem) {
        $pName = $lvPrinters.SelectedItem.Name
        
        # Embedded Rename Dialog XAML
        $renXaml = @"
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                Title="Rename Printer" Width="380" SizeToContent="Height" WindowStartupLocation="CenterOwner" 
                WindowStyle="None" AllowsTransparency="True" Background="Transparent"
                FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
            <Window.Resources>
                <Style TargetType="Button">
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="Button">
                                <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>
            </Window.Resources>
            <Border Background="$($c.Bg)" CornerRadius="8" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Margin="15">
                <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3"/></Border.Effect>
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="16,16,16,8" Cursor="Hand">
                        <TextBlock Text="Rename Printer" FontSize="16" FontWeight="SemiBold" Foreground="$($c.Fg)"/>
                    </Border>
                    
                    <StackPanel Grid.Row="1" Margin="16,8,16,16">
                        <TextBlock Text="Enter a new name for '$pName':" FontSize="12" Foreground="$($c.SecFg)" Margin="0,0,0,4" TextWrapping="Wrap"/>
                        <TextBox x:Name="txtNewName" Height="30" Background="$($c.BtnBg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Padding="6,4" VerticalContentAlignment="Center" Text="$pName"/>
                    </StackPanel>
                    
                    <Border Grid.Row="2" Background="$($c.BtnBg)" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="$($c.BtnBorder)">
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button x:Name="btnRenOk" Content="Rename" Width="80" Height="28" Margin="0,0,8,0" Background="$($c.PrimaryBg)" Foreground="$($c.PrimaryFg)" BorderThickness="0" IsDefault="True"/>
                            <Button x:Name="btnRenCancel" Content="Cancel" Width="80" Height="28" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1" IsCancel="True"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </Border>
        </Window>
"@
        $rReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($renXaml))
        $renWin = [System.Windows.Markup.XamlReader]::Load($rReader)
        $renWin.Owner = $prnWin
        
        $renWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $renWin.DragMove() })
        
        $txtNewName = $renWin.FindName("txtNewName")
        $btnRenOk = $renWin.FindName("btnRenOk")
        $btnRenCancel = $renWin.FindName("btnRenCancel")
        
        $script:newPrnName = ""
        
        if ($btnRenCancel) { $btnRenCancel.Add_Click({ $renWin.Close() }) }
        if ($btnRenOk) {
            $btnRenOk.Add_Click({
                if (-not [string]::IsNullOrWhiteSpace($txtNewName.Text)) {
                    $script:newPrnName = $txtNewName.Text
                    $renWin.Close()
                } else {
                    Show-LocalMessageBox -Message "Printer name cannot be blank." -Title "Validation" -IconType "Warning" -OwnerWindow $renWin | Out-Null
                }
            })
        }
        
        $txtNewName.SelectAll()
        $txtNewName.Focus() | Out-Null
        $renWin.ShowDialog() | Out-Null
        
        if ($script:newPrnName -and $script:newPrnName -ne $pName) {
            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
            try {
                Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                    param($old, $new)
                    Rename-Printer -Name $old -NewName $new -ErrorAction Stop
                } -ArgumentList $pName, $script:newPrnName
                
                Add-AppLog -Event "Printer Rename" -Username "System" -Details "Renamed printer '$pName' to '$($script:newPrnName)' on $ComputerName." -Config $Config -State $State -Status "Success"
                [System.Windows.Input.Mouse]::OverrideCursor = $null
                Show-LocalMessageBox -Message "Printer renamed to '$($script:newPrnName)' successfully." -Title "Success" -IconType "Information" -OwnerWindow $prnWin | Out-Null
                & $RefreshPrinters
            } catch {
                [System.Windows.Input.Mouse]::OverrideCursor = $null
                Show-LocalMessageBox -Message "Failed to rename printer:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $prnWin | Out-Null
                Add-AppLog -Event "Printer Rename" -Username "System" -Details "Failed to rename '$pName' on $($ComputerName): $($_.Exception.Message)" -Config $Config -State $State -Status "Error"
            }
        }
    }
}

$InstallPrinterAction = {
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    $lblStatus.Text = "Fetching remote driver manifest (15s timeout)..."
    $lblStatus.Foreground = [System.Windows.Media.Brushes]::Orange
    
    $frame = New-Object System.Windows.Threading.DispatcherFrame
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Action]{ $frame.Continue = $false }) | Out-Null
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
    
    $remoteDrivers = @()
    $job = $null
    try {
        $job = Invoke-Command -ComputerName $ComputerName -ScriptBlock { Get-PrinterDriver | Select-Object -ExpandProperty Name } -AsJob
        
        if ($job -and (Wait-Job $job -Timeout 15)) {
            if ($job.State -eq 'Failed') { throw $job.ChildJobs[0].JobStateInfo.Reason }
            $remoteDrivers = @(Receive-Job $job -ErrorAction Stop | Sort-Object)
        } else {
            if ($job) { Stop-Job $job }
            throw "Connection timed out (15s)."
        }
    } catch {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
        $lblStatus.Text = "Driver fetch failed."
        Show-LocalMessageBox -Message "Unable to retrieve drivers from ${ComputerName}.`n`nError: $($_.Exception.Message)" -Title "Fetch Failed" -IconType "Warning" -OwnerWindow $prnWin | Out-Null
        return
    } finally {
        if ($job) { Remove-Job $job -ErrorAction SilentlyContinue }
    }
    [System.Windows.Input.Mouse]::OverrideCursor = $null
    $lblStatus.Text = "Ready."
    $lblStatus.Foreground = [System.Windows.Media.Brushes]::Green
    
    # Embedded Install Dialog XAML
    $instXaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="Remote Printer Install" Width="420" SizeToContent="Height" WindowStartupLocation="CenterOwner" 
            WindowStyle="None" AllowsTransparency="True" Background="Transparent"
            FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
        <Window.Resources>
            <Style TargetType="Button">
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
            <Style TargetType="TextBox">
                <Setter Property="Background" Value="$($c.BtnBg)"/>
                <Setter Property="Foreground" Value="$($c.Fg)"/>
                <Setter Property="BorderBrush" Value="$($c.BtnBorder)"/>
                <Setter Property="BorderThickness" Value="1"/>
                <Setter Property="Padding" Value="6,4"/>
                <Setter Property="VerticalContentAlignment" Value="Center"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="TextBox">
                            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                <ScrollViewer x:Name="PART_ContentHost"/>
                            </Border>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </Window.Resources>
        <Border Background="$($c.Bg)" CornerRadius="8" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Margin="15">
            <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3"/></Border.Effect>
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="16,16,16,8" Cursor="Hand">
                    <TextBlock Text="Install Printer on $ComputerName" FontSize="16" FontWeight="SemiBold" Foreground="$($c.Fg)"/>
                </Border>
                
                <StackPanel Grid.Row="1" Margin="16,8,16,16">
                    <TextBlock Text="Printer Name (Friendly Name):" FontSize="11" Foreground="$($c.SecFg)" Margin="0,0,0,4"/>
                    <TextBox x:Name="txtName" Height="30" Margin="0,0,0,12"/>
                    
                    <TextBlock Text="Driver Name (Select or Type):" FontSize="11" Foreground="$($c.SecFg)" Margin="0,0,0,4"/>
                    <Grid Margin="0,0,0,12">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="32"/>
                            <ColumnDefinition Width="45"/>
                        </Grid.ColumnDefinitions>
                        <ComboBox x:Name="cbDriver" Grid.Column="0" Height="30" IsEditable="True" Margin="0,0,4,0"/>
                        <Button x:Name="btnUploadDriver" Grid.Column="1" Content="File" FontWeight="Bold" ToolTip="Upload and Stage Driver from local .INF file" Background="$($c.BtnBg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Margin="0,0,2,0"/>
                        <Button x:Name="btnLocalDriver" Grid.Column="2" Content="Local" FontSize="10" FontWeight="Bold" ToolTip="Copy installed driver from this PC" Background="$($c.BtnBg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                    </Grid>
                    
                    <TextBlock Text="IP Address:" FontSize="11" Foreground="$($c.SecFg)" Margin="0,0,0,4"/>
                    <TextBox x:Name="txtIP" Height="30"/>
                </StackPanel>
                
                <Border Grid.Row="2" Background="$($c.BtnBg)" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="$($c.BtnBorder)">
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="btnInstall" Content="Install" Width="80" Height="28" Margin="0,0,8,0" Background="$($c.PrimaryBg)" Foreground="$($c.PrimaryFg)" BorderThickness="0" IsDefault="True"/>
                        <Button x:Name="btnCancel" Content="Cancel" Width="80" Height="28" IsCancel="True" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>
    </Window>
"@
    $iReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($instXaml))
    $instWin = [System.Windows.Markup.XamlReader]::Load($iReader)
    
    $instWin.Owner = $prnWin
    $instWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $instWin.DragMove() })
    
    $cbDriver = $instWin.FindName("cbDriver")
    if ($cbDriver -and $remoteDrivers) { 
        $cbDriver.Items.Clear()
        foreach($d in $remoteDrivers) { $cbDriver.Items.Add($d) | Out-Null }
        if ($cbDriver.Items.Count -gt 0) { $cbDriver.SelectedIndex = 0 }
    }
    
    $txtName = $instWin.FindName("txtName")
    $txtIP = $instWin.FindName("txtIP")
    $btnUploadDriver = $instWin.FindName("btnUploadDriver")
    $btnLocalDriver = $instWin.FindName("btnLocalDriver")
    $btnInstall = $instWin.FindName("btnInstall")
    $btnCancel = $instWin.FindName("btnCancel")

    $script:prnInput = $null

    if ($btnUploadDriver) {
        $btnUploadDriver.Add_Click({
            $ofd = New-Object System.Windows.Forms.OpenFileDialog
            $ofd.Filter = "Driver Information File (*.inf)|*.inf"
            $ofd.Title = "Select Driver INF File"
            
            if ($ofd.ShowDialog() -eq "OK") {
                try {
                    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                    $infPath = $ofd.FileName
                    
                    # NATIVE FILE COPY AND STAGE
                    $drvDir = Split-Path $infPath
                    $folderName = Split-Path $drvDir -Leaf
                    $remoteTemp = "\\$ComputerName\c$\Temp\HDC_Drivers"
                    if (-not (Test-Path $remoteTemp)) { New-Item -ItemType Directory -Path $remoteTemp -Force | Out-Null }
                    
                    Copy-Item -Path $drvDir -Destination $remoteTemp -Recurse -Force
                    $remoteLocalPath = "C:\Temp\HDC_Drivers\$folderName"
                    
                    $updatedRemoteDrivers = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                        param($path)
                        # Stage using pnputil
                        $res = pnputil.exe /add-driver "$path\*.inf" /subdirs /install
                        
                        # Return fresh list
                        $drvList = @(Get-PrinterDriver -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name | Sort-Object)
                        return [PSCustomObject]@{ Pnp = ($res | Out-String); Drivers = $drvList }
                    } -ArgumentList $remoteLocalPath
                    
                    # Refresh the combo box with remote drivers
                    if ($cbDriver -and $updatedRemoteDrivers.Drivers) {
                        $currentText = $cbDriver.Text
                        $cbDriver.Items.Clear()
                        foreach($d in $updatedRemoteDrivers.Drivers) { $cbDriver.Items.Add($d) | Out-Null }
                        $cbDriver.Text = $currentText
                    }
                    
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    Show-LocalMessageBox -Message "Driver staged successfully.`n`nOutput:`n$($updatedRemoteDrivers.Pnp)`n`nPlease select the new driver from the dropdown or type its exact Model Name." -Title "Driver Staged" -IconType "Information" -OwnerWindow $instWin | Out-Null
                } catch { 
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    Show-LocalMessageBox -Message "Failed to deploy driver:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $instWin | Out-Null 
                }
            }
        })
    }
    
    if ($btnLocalDriver) {
        $btnLocalDriver.Add_Click({
            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
            Import-Module PrintManagement -ErrorAction SilentlyContinue
            
            # Use EXACT AccountMonitorOld.ps1 native logic without closures
            $localDrivers = @(Get-PrinterDriver -ErrorAction SilentlyContinue | Select-Object Name, InfPath | Sort-Object Name)
            
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            
            if (-not $localDrivers -or $localDrivers.Count -eq 0) {
                Show-LocalMessageBox -Message "No local drivers were found on this computer. Ensure you run as Administrator." -Title "No Drivers" -IconType "Warning" -OwnerWindow $instWin | Out-Null
                return
            }

            # Embedded Select Driver XAML
            $selXaml = @"
            <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                    Title="Select Local Driver" Width="400" Height="500" WindowStartupLocation="CenterOwner" 
                    WindowStyle="None" AllowsTransparency="True" Background="Transparent"
                    FontFamily="Segoe UI Variable Display, Segoe UI, sans-serif">
                <Window.Resources>
                    <Style TargetType="Button">
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="Button">
                                    <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.8"/></Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </Window.Resources>
                <Border Background="$($c.Bg)" CornerRadius="8" BorderBrush="$($c.BtnBorder)" BorderThickness="1" Margin="15">
                    <Border.Effect><DropShadowEffect BlurRadius="20" ShadowDepth="5" Opacity="0.3"/></Border.Effect>
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <Border x:Name="TitleBar" Grid.Row="0" Background="Transparent" Padding="16,16,16,8" Cursor="Hand">
                            <TextBlock Text="Select Local Driver" FontSize="16" FontWeight="SemiBold" Foreground="$($c.Fg)"/>
                        </Border>
                        
                        <ListBox x:Name="lbDrivers" Grid.Row="1" Margin="16,8,16,16" DisplayMemberPath="Name" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1"/>
                        
                        <Border Grid.Row="2" Background="$($c.BtnBg)" CornerRadius="0,0,8,8" Padding="16,12" BorderThickness="0,1,0,0" BorderBrush="$($c.BtnBorder)">
                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                                <Button x:Name="btnSelOk" Content="Use Selected" Width="100" Height="28" Margin="0,0,8,0" Background="$($c.PrimaryBg)" Foreground="$($c.PrimaryFg)" BorderThickness="0" IsDefault="True"/>
                                <Button x:Name="btnSelCancel" Content="Cancel" Width="80" Height="28" Background="$($c.Bg)" Foreground="$($c.Fg)" BorderBrush="$($c.BtnBorder)" BorderThickness="1" IsCancel="True"/>
                            </StackPanel>
                        </Border>
                    </Grid>
                </Border>
            </Window>
"@
            $sReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($selXaml))
            $locWin = [System.Windows.Markup.XamlReader]::Load($sReader)
            
            $locWin.FindName("TitleBar").Add_MouseLeftButtonDown({ $locWin.DragMove() })
            
            $lbLoc = $locWin.FindName("lbDrivers")
            $btnSelOk = $locWin.FindName("btnSelOk")
            $btnSelCancel = $locWin.FindName("btnSelCancel")
            
            $lbLoc.ItemsSource = $localDrivers
            $locWin.Owner = $instWin 
            
            $script:selectedLocalDriver = $null
            
            if ($btnSelCancel) { $btnSelCancel.Add_Click({ $locWin.Close() }) }
            if ($btnSelOk) {
                $btnSelOk.Add_Click({
                    if ($lbLoc.SelectedItem) {
                        $script:selectedLocalDriver = $lbLoc.SelectedItem
                        $locWin.Close()
                    }
                })
            }
            
            $locWin.ShowDialog() | Out-Null

            if ($script:selectedLocalDriver) {
                try {
                    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                    
                    $drvName = $script:selectedLocalDriver.Name
                    $infPath = $script:selectedLocalDriver.InfPath
                    
                    if (-not $infPath -or -not (Test-Path $infPath)) {
                        throw "Could not locate the physical INF file directory for driver: $drvName"
                    }
                    
                    # NATIVE FILE COPY AND STAGE (Bypass custom module)
                    $drvDir = Split-Path $infPath
                    $folderName = Split-Path $drvDir -Leaf
                    $remoteTemp = "\\$ComputerName\c$\Temp\HDC_Drivers"
                    if (-not (Test-Path $remoteTemp)) { New-Item -ItemType Directory -Path $remoteTemp -Force | Out-Null }
                    
                    # Copy the driver folder from Local C: to Remote C:
                    Copy-Item -Path $drvDir -Destination $remoteTemp -Recurse -Force
                    $remoteLocalPath = "C:\Temp\HDC_Drivers\$folderName"
                    
                    $updatedRemoteDrivers = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                        param($path, $name)
                        
                        # 1. Stage the driver natively using PnPUtil
                        pnputil.exe /add-driver "$path\*.inf" /subdirs /install | Out-Null
                        
                        # 2. Add the driver to the Print Spooler natively
                        try { Add-PrinterDriver -Name $name -ErrorAction Stop } catch { Write-Warning $_.Exception.Message }
                        
                        # 3. Return fresh list from Spooler
                        return @(Get-PrinterDriver -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name | Sort-Object)
                    } -ArgumentList $remoteLocalPath, $drvName
                    
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    
                    # Guarantee it appears in the Combobox
                    if ($cbDriver -and $updatedRemoteDrivers) {
                        $cbDriver.Items.Clear()
                        foreach($d in $updatedRemoteDrivers) { $cbDriver.Items.Add($d) | Out-Null }
                        
                        if ($cbDriver.Items.Contains($drvName)) {
                            $cbDriver.SelectedItem = $drvName
                        } else {
                            # Safety fallback if exact match failed but staged successfully
                            $cbDriver.Items.Add($drvName) | Out-Null
                            $cbDriver.SelectedItem = $drvName
                        }
                        $cbDriver.Text = $drvName 
                        $cbDriver.Focus()
                    }
                    Show-LocalMessageBox -Message "Driver '$drvName' uploaded, staged, and natively loaded into the remote spooler successfully." -Title "Success" -IconType "Information" -OwnerWindow $instWin | Out-Null
                } catch { 
                    [System.Windows.Input.Mouse]::OverrideCursor = $null
                    Show-LocalMessageBox -Message "Deploy failed:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $instWin | Out-Null 
                }
            }
        })
    }
    
    if ($btnCancel) { $btnCancel.Add_Click({ $instWin.Close() }) }
    
    if ($btnInstall) {
        $btnInstall.Add_Click({
            $pName = if ($txtName) { $txtName.Text } else { "" }
            $pDrv = if ($cbDriver) { $cbDriver.Text } else { "" }
            $pIp = if ($txtIP) { $txtIP.Text } else { "" }
            
            if ([string]::IsNullOrWhiteSpace($pName) -or [string]::IsNullOrWhiteSpace($pDrv) -or [string]::IsNullOrWhiteSpace($pIp)) {
                Show-LocalMessageBox -Message "All fields are required." -Title "Validation Error" -IconType "Warning" -OwnerWindow $instWin | Out-Null
                return
            }
            $script:prnInput = @{ Name = $pName; Driver = $pDrv; IP = $pIp }
            $instWin.Close()
        })
    }
    
    $instWin.ShowDialog() | Out-Null
    
    if ($script:prnInput) {
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        Add-AppLog -Event "Printer Install" -Username "System" -Details "Installing '$($script:prnInput.Name)' on $ComputerName..." -Config $Config -State $State -Status "Info"
        
        try {
            Install-RemotePrinter -ComputerName $ComputerName -PrinterName $script:prnInput.Name -DriverName $script:prnInput.Driver -IPAddress $script:prnInput.IP
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            Show-LocalMessageBox -Message "Printer installed successfully on ${ComputerName}." -Title "Success" -IconType "Information" -OwnerWindow $prnWin | Out-Null
            Add-AppLog -Event "Printer Install" -Username "System" -Details "Successfully installed printer on ${ComputerName}." -Config $Config -State $State -Status "Success" -Color "Green"
            & $RefreshPrinters
        } catch {
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            Show-LocalMessageBox -Message "Installation failed:`n$($_.Exception.Message)" -Title "Remote Error" -IconType "Error" -OwnerWindow $prnWin | Out-Null
            Add-AppLog -Event "Printer Install" -Username "System" -Details "Failed on ${ComputerName}: $($_.Exception.Message)" -Config $Config -State $State -Status "Error" -Color "Red"
        }
    }
}

$RestartSpoolerAction = {
    $conf = Show-LocalMessageBox -Message "Are you sure you want to restart the Print Spooler on $ComputerName?" -Title "Confirm Restart" -ButtonType "YesNo" -IconType "Warning" -OwnerWindow $prnWin
    if ($conf -eq "Yes") {
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        try {
            Restart-RemoteSpooler -ComputerName $ComputerName
            Add-AppLog -Event "Service" -Username "System" -Details "Spooler restarted on $ComputerName" -Config $Config -State $State
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            Show-LocalMessageBox -Message "Print Spooler restarted successfully." -Title "Success" -OwnerWindow $prnWin | Out-Null
            & $RefreshPrinters
        } catch {
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            Show-LocalMessageBox -Message "Failed to restart spooler:`n$($_.Exception.Message)" -Title "Error" -IconType "Error" -OwnerWindow $prnWin | Out-Null
        }
    }
}

# 7. Wire Up Bindings
if ($btnRefresh) { $btnRefresh.Add_Click($RefreshPrinters) }
if ($btnClose) { $btnClose.Add_Click({ $prnWin.Close() }) }

if ($btnPrnDetails) { $btnPrnDetails.Add_Click($ShowPrinterDetails) }
if ($ctxPrnDetails) { $ctxPrnDetails.Add_Click($ShowPrinterDetails) }
if ($ctxPrnRename) { $ctxPrnRename.Add_Click($RenamePrinterAction) }
if ($lvPrinters) { $lvPrinters.Add_MouseDoubleClick($ShowPrinterDetails) }

if ($btnPrnRemove) { $btnPrnRemove.Add_Click($RemovePrinterAction) }
if ($ctxPrnRemove) { $ctxPrnRemove.Add_Click($RemovePrinterAction) }

if ($btnInstallPrinter) { $btnInstallPrinter.Add_Click($InstallPrinterAction) }
if ($btnRestartSpooler) { $btnRestartSpooler.Add_Click($RestartSpoolerAction) }

# 8. Start
$prnWin.Add_Loaded({ & $RefreshPrinters })
$prnWin.ShowDialog() | Out-Null