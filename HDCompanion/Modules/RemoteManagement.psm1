# ============================================================================
# RemoteManagement.psm1 - Remote System and Session Management
# ============================================================================

function Get-RemoteProcesses {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [switch]$AsJob
    )
    
    $cmdArgs = @{
        ComputerName = $ComputerName
        ScriptBlock = {
            Get-Process | Select-Object Name, Id, 
                @{Name='CPU';Expression={if($_.CPU){[math]::Round($_.CPU, 2)}else{0}}}, 
                @{Name='MemMB';Expression={if($_.WorkingSet64){[math]::Round($_.WorkingSet64 / 1MB, 2)}else{0}}}, 
                Description -ErrorAction SilentlyContinue
        }
        ErrorAction = 'Stop'
    }
    
    if ($AsJob) {
        $cmdArgs.Add('AsJob', $true)
    }
    
    Invoke-Command @cmdArgs
}

function Stop-RemoteProcess {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [int]$ProcessId
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock { 
        param($pidToKill) 
        Stop-Process -Id $pidToKill -Force -ErrorAction Stop
    } -ArgumentList $ProcessId -ErrorAction Stop
}

function Get-RemoteActiveUsers {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    
    $quserOutput = quser /server:$ComputerName 2>&1
    
    if ($LASTEXITCODE -ne 0 -or $quserOutput -match "No User exists" -or $quserOutput -match "Error") {
        return @()
    }
    
    $sessions = @()
    for ($i = 1; $i -lt $quserOutput.Count; $i++) {
        $line = $quserOutput[$i] -replace '^>', ' ' # Remove active session indicator
        
        $uName = $null
        $sId = $null
        $sState = $null
        
        if ($line.Length -ge 65) {
            # Fixed-width parsing for reliability
            $uName = $line.Substring(0, 22).Trim()
            $sId = $line.Substring(41, 5).Trim()
            $sState = $line.Substring(46, 8).Trim()
        } else {
            # Fallback regex/split parsing
            $tok = $line -split '\s+' | Where-Object { $_ }
            if ($tok.Count -ge 5) {
                $uName = $tok[0]
                $sId = ($tok | Where-Object { $_ -match '^\d+$' } | Select-Object -First 1)
                $sState = if ($line -match 'Active') { 'Active' } else { 'Disc' }
            }
        }
        
        if ($uName -and $sId) {
            $sessions += [PSCustomObject]@{
                Username  = $uName
                SessionId = $sId
                State     = $sState
            }
        }
    }
    return $sessions
}

function Stop-RemoteUserSession {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$SessionId
    )
    
    $logoffRes = logoff $SessionId /server:$ComputerName 2>&1
    
    if ($LASTEXITCODE -ne 0 -or (-not [string]::IsNullOrWhiteSpace($logoffRes) -and $logoffRes -match "Error")) {
        throw "Failed to logoff session $SessionId on ${ComputerName}. Details: $logoffRes"
    }
    return $true
}

function Invoke-RemoteGPResult {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputFile,
        
        [string]$TargetUser = $null
    )
    
    $gpArgs = @("/S", $ComputerName)
    if (-not [string]::IsNullOrWhiteSpace($TargetUser)) {
        $gpArgs += "/USER", $TargetUser
    } else {
        $gpArgs += "/SCOPE", "COMPUTER"
    }
    $gpArgs += "/H", $OutputFile, "/F"
    
    $p = Start-Process -FilePath "gpresult.exe" -ArgumentList $gpArgs -NoNewWindow -PassThru -Wait
    
    if ($p.ExitCode -ne 0 -or -not (Test-Path $OutputFile)) {
        throw "GPResult generation failed. Process returned exit code $($p.ExitCode)."
    }
    
    return [PSCustomObject]@{
        Success    = $true
        OutputFile = $OutputFile
        ExitCode   = $p.ExitCode
    }
}

function Get-RemotePrinterDrivers {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock { 
        Get-PrinterDriver | Select-Object -ExpandProperty Name 
    } -ErrorAction Stop
}

function Get-RemotePrinters {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock { 
        Get-Printer | Select-Object Name, DriverName, PortName, Shared, Published, DeviceType, PrinterStatus, Location, Comment
    } -ErrorAction Stop
}

function Remove-RemotePrinter {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$PrinterName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        param($pName)
        Remove-Printer -Name $pName -ErrorAction Stop
    } -ArgumentList $PrinterName -ErrorAction Stop
}

function Install-RemotePrinter {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$PrinterName,
        
        [Parameter(Mandatory=$true)]
        [string]$DriverName,
        
        [Parameter(Mandatory=$true)]
        [string]$IPAddress
    )
    
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        param($PName, $PDriver, $PIP)
        
        $portName = "IP_$PIP"
        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
            Add-PrinterPort -Name $portName -PrinterHostAddress $PIP -ErrorAction Stop
        }
        
        if (-not (Get-PrinterDriver -Name $PDriver -ErrorAction SilentlyContinue)) {
            Add-PrinterDriver -Name $PDriver -ErrorAction Stop
        }

        Add-Printer -Name $PName -DriverName $PDriver -PortName $portName -ErrorAction Stop
        return $true
    } -ArgumentList $PrinterName, $DriverName, $IPAddress -ErrorAction Stop
}

function Deploy-DriverToRemote {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourceInfPath,
        
        [Parameter(Mandatory=$true)]
        [string]$RemoteComputer
    )
    
    $parentDir = [System.IO.Path]::GetDirectoryName($SourceInfPath)
    $folderName = [System.IO.Path]::GetFileName($parentDir)
    $infName = [System.IO.Path]::GetFileName($SourceInfPath)
    
    $destPath = "\\$RemoteComputer\C`$\Temp\Drivers\Upload_$(Get-Date -Format 'yyyyMMddHHmmss')_$folderName"
    
    if (-not (Test-Path $destPath)) { 
        New-Item -Path $destPath -ItemType Directory -Force | Out-Null 
    }
    
    # Copy all files from the driver directory to remote temp location
    Copy-Item -Path "$parentDir\*" -Destination $destPath -Recurse -Force -ErrorAction Stop
    
    $localDest = $destPath.Replace("\\$RemoteComputer\C`$", "C:")
    
    # Execute PnPUtil on remote computer to stage the driver
    $pnpOutput = Invoke-Command -ComputerName $RemoteComputer -ScriptBlock {
        param($Path, $InfFile)
        $driverPath = Join-Path $Path $InfFile
        if (Test-Path $driverPath) {
            return (pnputil.exe /add-driver "$driverPath" /install)
        } else {
            throw "Driver file not found on remote machine at $driverPath"
        }
    } -ArgumentList $localDest, $infName -ErrorAction Stop
    
    return $pnpOutput
}

function Get-RemoteUptime {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    # FIX: Job wrapped with a manual sleep loop to bypass parameter errors and prevent hangs
    $job = Start-Job -ScriptBlock {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $args[0] -ErrorAction Stop
        $lastBoot = $os.LastBootUpTime
        $uptime = (Get-Date) - $lastBoot
        return [PSCustomObject]@{ ComputerName = $args[0]; LastBootUpTime = $lastBoot; Days = $uptime.Days; Hours = $uptime.Hours; Minutes = $uptime.Minutes }
    } -ArgumentList $ComputerName
    
    $tCount = 100
    while ($job.State -eq 'Running' -and $tCount -gt 0) { Start-Sleep -Milliseconds 100; $tCount-- }
    
    if ($job.State -eq 'Running') {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        throw "Connection timed out after 10 seconds."
    }
    
    $res = Receive-Job $job -ErrorAction Stop
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    return $res
}

function Restart-RemoteSpooler {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Restart-Service "Spooler" -Force -ErrorAction Stop
    } -ErrorAction Stop
}

function Invoke-RemotePowerAction {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("Restart", "Shutdown")]
        [string]$Action
    )
    if ($Action -eq "Restart") {
        Restart-Computer -ComputerName $ComputerName -Force -ErrorAction Stop
    } else {
        Stop-Computer -ComputerName $ComputerName -Force -ErrorAction Stop
    }
}

function Start-RemoteProcess {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$CommandLine
    )
    # FIX: Job wrapped with a manual sleep loop to bypass parameter errors and prevent hangs
    $job = Start-Job -ScriptBlock {
        param($c, $cmd)
        $proc = Invoke-CimMethod -ClassName Win32_Process -ComputerName $c -MethodName Create -Arguments @{ CommandLine = $cmd } -ErrorAction Stop
        if ($proc.ReturnValue -ne 0) {
            throw "WMI Return Code: $($proc.ReturnValue). (2 = Access Denied, 3 = Insufficient Privilege, 9 = Path Not Found)"
        }
    } -ArgumentList $ComputerName, $CommandLine
    
    $tCount = 100
    while ($job.State -eq 'Running' -and $tCount -gt 0) { Start-Sleep -Milliseconds 100; $tCount-- }
    
    if ($job.State -eq 'Running') {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        throw "Connection timed out (10s). The computer may be offline or firewalled."
    }
    
    $res = Receive-Job $job -ErrorAction Stop
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    return $res
}

function Get-RemoteServices {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    # FIX: Job wrapped with a manual sleep loop to bypass parameter errors and prevent hangs
    $job = Start-Job -ScriptBlock {
        Get-CimInstance -ClassName Win32_Service -ComputerName $args[0] -ErrorAction Stop | Select-Object Name, DisplayName, State, StartMode
    } -ArgumentList $ComputerName
    
    $tCount = 100
    while ($job.State -eq 'Running' -and $tCount -gt 0) { Start-Sleep -Milliseconds 100; $tCount-- }
    
    if ($job.State -eq 'Running') {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        throw "Connection timed out (10s). The computer may be offline or firewalled."
    }
    
    $res = Receive-Job $job -ErrorAction Stop
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    return $res
}

function Invoke-RemoteServiceAction {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$ServiceName,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("Start", "Stop", "Restart")]
        [string]$Action
    )
    
    # FIX: Job wrapped with a manual sleep loop to bypass parameter errors and prevent hangs
    $job = Start-Job -ScriptBlock {
        param($c, $s, $a)
        $svc = Get-CimInstance -ClassName Win32_Service -ComputerName $c -Filter "Name='$s'" -ErrorAction Stop
        if (-not $svc) { throw "Service '$s' not found on target." }

        if ($a -eq "Start") {
            Invoke-CimMethod -InputObject $svc -MethodName StartService -ErrorAction Stop | Out-Null
        } elseif ($a -eq "Stop") {
            Invoke-CimMethod -InputObject $svc -MethodName StopService -ErrorAction Stop | Out-Null
        } elseif ($a -eq "Restart") {
            Invoke-CimMethod -InputObject $svc -MethodName StopService -ErrorAction SilentlyContinue | Out-Null
            Start-Sleep -Seconds 2
            Invoke-CimMethod -InputObject $svc -MethodName StartService -ErrorAction Stop | Out-Null
        }
    } -ArgumentList $ComputerName, $ServiceName, $Action
    
    $tCount = 150
    while ($job.State -eq 'Running' -and $tCount -gt 0) { Start-Sleep -Milliseconds 100; $tCount-- }
    
    if ($job.State -eq 'Running') {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        throw "Command timed out (15s)."
    }
    
    $res = Receive-Job $job -ErrorAction Stop
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    return $res
}

function Get-RemoteDiskSpace {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop | Select-Object Size, FreeSpace
    } -ErrorAction Stop
}

function Get-RemoteUserProfiles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Get-CimInstance -ClassName Win32_UserProfile -Filter "Special=False" -ErrorAction Stop | Select-Object LocalPath, LastUseTime, Loaded, SID
    } -ErrorAction Stop
}

function Remove-RemoteUserProfile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$SID
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        param($s)
        $prof = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$s'" -ErrorAction Stop
        Invoke-CimMethod -InputObject $prof -MethodName Delete -ErrorAction Stop
    } -ArgumentList $SID -ErrorAction Stop
}

function Get-RemoteInstalledSoftware {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        $paths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
        Get-ItemProperty $paths -ErrorAction SilentlyContinue | 
            Where-Object { $_.DisplayName -and $_.SystemComponent -ne 1 -and $_.ParentKeyName -eq $null } | 
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, UninstallString, QuietUninstallString | 
            Sort-Object DisplayName -Unique
    } -ErrorAction SilentlyContinue
}

function Uninstall-RemoteSoftware {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [string]$QuietUninstallString,
        
        [string]$UninstallString
    )
    $cmd = $QuietUninstallString
    if ([string]::IsNullOrWhiteSpace($cmd)) {
        $cmd = $UninstallString
        if ([string]::IsNullOrWhiteSpace($cmd)) { throw "No uninstall string found for this application." }
        if ($cmd -match '(?i)msiexec') {
            $cmd = $cmd -replace '(?i)/I', '/X'
            if ($cmd -notmatch '(?i)/q') { $cmd += ' /qn /norestart' }
        } elseif ($cmd -match '(?i)unins\d{3}\.exe') {
            $cmd += ' /VERYSILENT /SUPPRESSMSGBOXES /NORESTART'
        } elseif ($cmd -match '(?i)uninstall\.exe') {
            $cmd += ' /S'
        }
    }
    Start-RemoteProcess -ComputerName $ComputerName -CommandLine $cmd
}

function Get-RemoteDevices {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Get-PnpDevice -ErrorAction SilentlyContinue | Select-Object FriendlyName, Class, Status, Manufacturer, InstanceId | Where-Object { -not [string]::IsNullOrWhiteSpace($_.FriendlyName) }
    } -ErrorAction Stop
}

function Set-RemoteDeviceState {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory=$true)]
        [string]$InstanceId,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("Enable", "Disable")]
        [string]$Action
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        param($id, $act)
        if ($act -eq 'Enable') {
            Enable-PnpDevice -InstanceId $id -Confirm:$false -ErrorAction Stop
        } else {
            Disable-PnpDevice -InstanceId $id -Confirm:$false -ErrorAction Stop
        }
    } -ArgumentList $InstanceId, $Action -ErrorAction Stop
}

function Get-RemoteEventLogs {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        try {
            Get-WinEvent -FilterHashtable @{LogName='System','Application'; Level=1,2,3} -MaxEvents 150 -ErrorAction Stop |
                Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, Message
        } catch { @() }
    } -ErrorAction Stop
}

Export-ModuleMember -Function *