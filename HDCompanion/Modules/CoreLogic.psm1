# ============================================================================
# CoreLogic.psm1 - Configuration, Logging, and App Setup
# ============================================================================

function Get-AppConfig {
    $centralConfigPath = "\\vm-isserver\toolkit\[it toolkit]\hdcompanion\hdcompanioncfg.json"
    
    # IMPROVEMENT: Prevent app hang if network share is unreachable
    $isNetworkAvailable = $false
    if ($centralConfigPath -match "^\\\\([^\\]+)") {
        $serverName = $matches[1]
        try {
            $ping = New-Object System.Net.NetworkInformation.Ping
            if ($ping.Send($serverName, 500).Status -eq "Success") { $isNetworkAvailable = $true }
        } catch {}
    }

    if ($isNetworkAvailable -and (Test-Path -LiteralPath $centralConfigPath)) {
        try {
            $jsonConfig = Get-Content -LiteralPath $centralConfigPath -Raw | ConvertFrom-Json
            $jsonConfig | Add-Member -MemberType NoteProperty -Name "LoadedConfigPath" -Value $centralConfigPath -Force
            $Script:CurrentConfigPath = $centralConfigPath
            return $jsonConfig
        } catch {}
    }

    $localConfigPath = Join-Path $PSScriptRoot "hdcompanioncfg.json"
    if (Test-Path -LiteralPath $localConfigPath) {
        try {
            $jsonConfig = Get-Content -LiteralPath $localConfigPath -Raw | ConvertFrom-Json
            $jsonConfig | Add-Member -MemberType NoteProperty -Name "LoadedConfigPath" -Value $localConfigPath -Force
            $Script:CurrentConfigPath = $localConfigPath
            return $jsonConfig
        } catch {}
    } 

    $Script:CurrentConfigPath = "Embedded Defaults (File not found)"
    $defaultConfig = [PSCustomObject]@{
        LoadedConfigPath = $Script:CurrentConfigPath
        GeneralSettings = @{
            DomainName         = "pscu.local"
            LogDirectoryUNC    = "\\vm-isserver\toolkit\[it toolkit]\hdcompanion\Logs"
            LogRetentionDays   = 90
            FilteredUsers      = @("guest", "support", "admbrian")
            DefaultTheme       = "Light"
            SplashtopAPIToken  = "" 
            FreshserviceDomain = "https://pelicanstatecreditunion.freshservice.com"
            FreshserviceAPIKey = ""
        }
        AutoSettings = @{
            AutoRefreshOptions = @("30 seconds", "1 minute", "2 minutes", "5 minutes", "10 minutes", "15 minutes", "30 minutes")
            AutoUnlockIntervalSeconds = 300
        }
        EmailSettings = @{
            EnableEmailNotifications = $false 
            SmtpServer   = "10.104.100.165"
            SmtpPort     = 25
            EnableSsl    = $true
            FromAddress  = "AccountMonitor@pelicancu.com"
            ToAddress    = @("tdawsey@pelicancu.com", "bhigginbotham@pelicancu.com", "adaigle@pelicancu.com")
            SmtpUsername = ""
            SmtpPassword = ""
        }
        LightModeColors = @{
            Text = @(32, 32, 32)
            Primary = @(0, 90, 158)
            Danger = @(200, 30, 30)
            TextSecondary = @(96, 96, 96)
            Background = @(243, 243, 243)
            Success = @(16, 124, 16)
            Card = @(255, 255, 255)
            Secondary = @(200, 200, 200) 
            Hover = @(229, 229, 229)
            AltRow = @(249, 249, 249)
            OnlineText = @(0, 100, 0)
            OfflineText = @(178, 34, 34)
        }
        DarkModeColors = @{
            Text = @(255, 255, 255)
            Primary = @(0, 120, 212) 
            Danger = @(255, 100, 100)
            TextSecondary = @(200, 200, 200)
            Background = @(32, 32, 32)
            Success = @(108, 203, 95)
            Card = @(45, 45, 48)
            Secondary = @(100, 100, 100)
            Hover = @(80, 80, 85)
            AltRow = @(55, 55, 60)
            OnlineText = @(150, 255, 150)
            OfflineText = @(255, 150, 150)
        }
        ControlProperties = @{
            TitleLabel = @{ Text = "Helpdesk Companion" }
            SubtitleLabel = @{ Text = "Manage locked accounts, remote systems, and AD" }
            UnlockButton = @{ Text = "Unlock Selected" }
            UnlockAllButton = @{ Text = "Unlock All" }
            RefreshButton = @{ Text = "Refresh" }
            SearchButton = @{ Text = "Search" }
            ViewLogButton = @{ Text = "View Logs" }
        }
    }
    return $defaultConfig
}

function Initialize-LogDirectory {
    param($Config)
    $path = $Config.GeneralSettings.LogDirectoryUNC
    
    if (-not (Test-Path -LiteralPath $path)) {
        try {
            $escapedPath = $path.Replace("[", "`[").Replace("]", "`]")
            New-Item -Path $escapedPath -ItemType Directory -Force | Out-Null
        } catch {
            $fallback = Join-Path $env:TEMP "ADUnlockLogs"
            $Config.GeneralSettings.LogDirectoryUNC = $fallback
            if (-not (Test-Path -LiteralPath $fallback)) {
                New-Item -Path $fallback -ItemType Directory -Force | Out-Null
            }
        }
    }
    
    if (-not (Test-Path "B:")) {
        try { subst B: $Config.GeneralSettings.LogDirectoryUNC | Out-Null } catch {}
    }
}

function Add-AppLog {
    param($Event, $Username, $Details, $Config, $State, $Status="Success", $Color="Black")
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $fileTimestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
    $msg = "[$timestamp] $Event - ${Username}: $Details"
    
    if ($State.UIControls.txtLog) {
        $State.UIControls.txtLog.Dispatcher.Invoke({
            $paragraph = New-Object System.Windows.Documents.Paragraph
            $paragraph.Margin = "0"
            $run = New-Object System.Windows.Documents.Run($msg)
            
            if ($State.CurrentTheme -eq "Dark") {
                if ($Color -eq "Black") {  $run.Foreground = [System.Windows.Media.Brushes]::White  } 
                else {
                    switch ($Color) {
                        "Blue"   { $run.Foreground = [System.Windows.Media.Brushes]::DeepSkyBlue }
                        "Red"    { $run.Foreground = [System.Windows.Media.Brushes]::LightCoral }
                        "Green"  { $run.Foreground = [System.Windows.Media.Brushes]::LightGreen }
                        "Orange" { $run.Foreground = [System.Windows.Media.Brushes]::Orange }
                        Default { $run.Foreground = [System.Windows.Media.Brushes]::White }
                    }
                }
            } else {
                if ($Color -eq "Black") {  $run.Foreground = [System.Windows.Media.Brushes]::Black  } 
                else {
                    try {
                        $brushConverter = New-Object System.Windows.Media.BrushConverter
                        $run.Foreground = $brushConverter.ConvertFromString($Color)
                        if ($Color -eq "Green" -or $Color -eq "Red" -or $Color -eq "Blue") { $run.FontWeight = "Bold" }
                    } catch { $run.Foreground = [System.Windows.Media.Brushes]::Black }
                }
            }

            $paragraph.Inlines.Add($run)
            $State.UIControls.txtLog.Document.Blocks.Add($paragraph)
            $State.UIControls.txtLog.ScrollToEnd()
        })
    }

    try {
        $logDir = $Config.GeneralSettings.LogDirectoryUNC
        $today = Get-Date -Format "yyyyMMdd"
        $logFile = Join-Path $logDir "UnlockLog_$today.csv"
        $operator = $env:USERNAME
        $machine = $env:COMPUTERNAME

        $obj = [PSCustomObject]@{ Timestamp = $fileTimestamp; Event = $Event; Username = $Username; Details = $Details; Status = $Status; Operator = $operator; MachineName = $machine }
        $obj | Export-Csv -LiteralPath $logFile -NoTypeInformation -Append -Encoding UTF8
    } catch {}
}

function Get-AppLogFiles {
    param($Config)
    $logDir = $Config.GeneralSettings.LogDirectoryUNC
    $allLogs = @()
    if (Test-Path -LiteralPath $logDir) {
        $files = Get-ChildItem -LiteralPath $logDir -Filter "UnlockLog_*.csv"
        foreach ($f in $files) {
            $content = Import-Csv -LiteralPath $f.FullName
            $allLogs += $content
        }
    }
    return $allLogs
}

function Get-FSAssetDetails {
    param($AssetName, $Config)
    $domain = $Config.GeneralSettings.FreshserviceDomain
    if (-not $domain) { $domain = "https://pelicanstatecreditunion.freshservice.com" }
    $domain = $domain.TrimEnd('/')
    if (-not $domain.StartsWith("http")) { $domain = "https://$domain" }
    
    $apiKey = $Config.GeneralSettings.FreshserviceAPIKey
    if (-not $apiKey) { throw "Freshservice API Key is missing. Please add 'FreshserviceAPIKey' to your config." }
    
    $encoded = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($apiKey):X"))
    $headers = @{ "Authorization" = "Basic $encoded" }
    $AssetQS = "name:'$AssetName'"
    $AssetEncoded = [Uri]::EscapeDataString("`"$AssetQS`"")
    $AssetURL = "$domain/api/v2/assets?query=$AssetEncoded"
    
    try {
        $AssetResp = Invoke-RestMethod -Uri $AssetURL -Headers $headers -Method Get -ErrorAction Stop
        if ($AssetResp.assets.Count -gt 0) { return ($AssetResp.assets | Sort-Object { [DateTime]$_.updated_at } -Descending | Select-Object -First 1) }
        return $null
    } catch {
        if ($_.Exception.Response.StatusCode -eq 'NotFound' -or $_.Exception.Message -match "404") { throw "HTTP 404 Not Found. Ensure your domain is correct in settings." } 
        else { throw $_.Exception.Message }
    }
}

function Get-FSUserAsset {
    param($User, $Config)
    $domain = $Config.GeneralSettings.FreshserviceDomain
    if (-not $domain) { $domain = "https://pelicanstatecreditunion.freshservice.com" }
    $domain = $domain.TrimEnd('/')
    if (-not $domain.StartsWith("http")) { $domain = "https://$domain" }
    
    $apiKey = $Config.GeneralSettings.FreshserviceAPIKey
    if (-not $apiKey) { throw "Freshservice API Key is missing. Please add 'FreshserviceAPIKey' to your config." }
    
    $encoded = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($apiKey):X"))
    $headers = @{ "Authorization" = "Basic $encoded" }

    try {
        if ([string]::IsNullOrWhiteSpace($User.EmailAddress)) {
            throw "User does not have an Email Address in Active Directory. Freshservice lookup requires an email."
        }

        # 1. Look up Requester by Email Address using the correct V2 query format
        $reqId = $null
        $email = $User.EmailAddress.Trim()
        $reqQuery = "primary_email:'$email'"
        $reqEncoded = [Uri]::EscapeDataString("`"$reqQuery`"")
        $ReqURL = "$domain/api/v2/requesters?query=$reqEncoded"

        $ReqResp = Invoke-RestMethod -Uri $ReqURL -Headers $headers -Method Get -ErrorAction Stop

        if ($ReqResp.requesters -and $ReqResp.requesters.Count -gt 0) { 
            $reqId = $ReqResp.requesters[0].id 
        } else {
            throw "User email '$email' was not found in the Freshservice Requesters database."
        }

        # 2. Query assets assigned to this specific user ID
        if ($reqId) {
            $AssetQS = "user_id:$reqId"
            $AssetEncoded = [Uri]::EscapeDataString("`"$AssetQS`"")
            $AssetURL = "$domain/api/v2/assets?query=$AssetEncoded"
            $AssetResp = Invoke-RestMethod -Uri $AssetURL -Headers $headers -Method Get -ErrorAction Stop
            
            if ($AssetResp.assets -and $AssetResp.assets.Count -gt 0) { 
                return $AssetResp.assets 
            }
        }
        
        return @()
    } catch {
        $errMsg = $_.Exception.Message
        
        if ($_.Exception.Response) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $respBody = $reader.ReadToEnd()
                $errMsg += "`n`nAPI Response:`n$respBody"
            } catch {}
        }

        if ($_.Exception.Response.StatusCode -eq 'NotFound' -or $errMsg -match "404") { 
            throw "HTTP 404 Not Found. Ensure your domain is correct in settings." 
        } else { 
            throw $errMsg 
        }
    }
}

function Get-FSUserRecord {
    param($User, $Config)
    
    $domain = $Config.GeneralSettings.FreshserviceDomain
    if (-not $domain) { $domain = "https://pelicanstatecreditunion.freshservice.com" }
    $domain = $domain.TrimEnd('/')
    if (-not $domain.StartsWith("http")) { $domain = "https://$domain" }
    
    $apiKey = $Config.GeneralSettings.FreshserviceAPIKey
    if (-not $apiKey) { throw "Freshservice API Key is missing. Please add 'FreshserviceAPIKey' to your config." }
    
    if ([string]::IsNullOrWhiteSpace($User.EmailAddress)) {
        throw "User does not have an Email Address in Active Directory. Freshservice lookup requires an email."
    }

    $encoded = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($apiKey):X"))
    $headers = @{ "Authorization" = "Basic $encoded" }
    $email = $User.EmailAddress.Trim()

    # 1. Try finding them as a Requester
    try {
        $reqQuery = "primary_email:'$email'"
        $reqEncoded = [Uri]::EscapeDataString("`"$reqQuery`"")
        $ReqURL = "$domain/api/v2/requesters?query=$reqEncoded"
        
        $ReqResp = Invoke-RestMethod -Uri $ReqURL -Headers $headers -Method Get -ErrorAction Stop
        
        if ($ReqResp.requesters -and $ReqResp.requesters.Count -gt 0) {
            $reqId = $ReqResp.requesters[0].id
            return [PSCustomObject]@{
                Type = "Requester"
                Id = $reqId
                Url = "$domain/itil/requesters/$reqId"
            }
        }
    } catch {
        # Suppress errors on the first pass, we will try the Agent database next
    }

    # 2. Try finding them as an Agent (If they aren't a standard requester)
    try {
        $agentEncoded = [Uri]::EscapeDataString($email)
        $AgentURL = "$domain/api/v2/agents?email=$agentEncoded"
        
        $AgentResp = Invoke-RestMethod -Uri $AgentURL -Headers $headers -Method Get -ErrorAction Stop
        
        if ($AgentResp.agents -and $AgentResp.agents.Count -gt 0) {
            $agentId = $AgentResp.agents[0].id
            return [PSCustomObject]@{
                Type = "Agent"
                Id = $agentId
                Url = "$domain/admin/agents/$agentId/edit"
            }
        }
    } catch {
        throw "Error querying the Agents database: $($_.Exception.Message)"
    }

    return $null
}

function Get-FluentThemeColors {
    param($State)
    $isDark = ($null -ne $State -and $State.CurrentTheme -eq "Dark")
    return @{
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
        Danger     = if ($isDark) { "#FF5252" } else { "#D13438" }
    }
}

function Load-XamlWindow {
    param( [Parameter(Mandatory=$true)][string]$XamlPath, [Parameter(Mandatory=$true)][hashtable]$ThemeColors )
    $xamlText = Get-Content -Path $XamlPath -Raw
    foreach ($key in $ThemeColors.Keys) {
        $placeholder = "{Theme_$key}"
        $xamlText = $xamlText.Replace($placeholder, $ThemeColors[$key])
    }
    
    $xmlSettings = New-Object System.Xml.XmlReaderSettings
    $xmlSettings.DtdProcessing = [System.Xml.DtdProcessing]::Parse
    
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlText), $xmlSettings)
    return [System.Windows.Markup.XamlReader]::Load($reader)
}

Export-ModuleMember -Function *