<#
.SYNOPSIS
    Discover and connect to Beckhoff devices via ADS.
.DESCRIPTION
    This script discovers Beckhoff devices on the network and provides connection options in a modular, testable structure.
.PARAMETER TimeoutSeconds
    Timeout for user input in seconds.
.PARAMETER WinSCPPath
    Path to WinSCP executable.
.PARAMETER CerHostPath
    Path to CERHost executable.
.PARAMETER AdminUserName
    Administrator username for SSH and RDP connections.
.PARAMETER AdminPassword
    Administrator password for SSH and RDP connections.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [int]$TimeoutSeconds = 10,
    [Parameter(Mandatory=$false)]
    [string]$WinSCPPath    = "C:\Program Files (x86)\WinSCP\WinSCP.exe",
    [Parameter(Mandatory=$false)]
    [string]$CerHostPath   = "$PSScriptRoot\CERHOST.exe",
    [Parameter(Mandatory=$false)]
    [string]$AdminUserName = "Administrator",
    [Parameter(Mandatory=$false)]
    [string]$AdminPassword = "1"
)

Write-Verbose "PowerShell Version: $($PSVersionTable.PSVersion)"

function Read-InputWithTimeout {
    [CmdletBinding()]
    param(
        [int]$TimeoutSeconds = 10,
        [switch]$AllowRefresh
    )
    try {
        $endTime     = (Get-Date).AddSeconds($TimeoutSeconds)
        $inputString = ""
        while ((Get-Date) -lt $endTime) {
            if ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($false)
                if ($key.Key -eq 'Enter') {
                    if ($AllowRefresh) { return 'refresh' }
                    break
                } else {
                    $inputString += $key.KeyChar
                }
            }
            Start-Sleep -Milliseconds 100
        }
        return $inputString.Trim()
    } catch {
        throw "Error in Read-InputWithTimeout: $_"
    }
}

function Test-CERHostAvailability {
    [CmdletBinding()]
    param(
        [string]$IPAddress,
        [int]$Port               = 987,
        [int]$TimeoutMilliseconds = 1000
    )
    try {
        $tcpClient   = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $tcpClient.BeginConnect($IPAddress, $Port, $null, $null)
        if ($asyncResult.AsyncWaitHandle.WaitOne($TimeoutMilliseconds, $false)) {
            $tcpClient.EndConnect($asyncResult)
            return $true
        }
        return $false
    } catch {
        return $false
    } finally {
        if ($tcpClient) { $tcpClient.Close() }
    }
}

function Test-TcXaeMgmtModule {
    [CmdletBinding()]
    param()
    try {
        if ($PSVersionTable.PSEdition -eq 'Desktop') {
            $versionRange = '3.2.*'
        } elseif ($PSVersionTable.PSEdition -eq 'Core') {
            $versionRange = '6.*'
        } else {
            throw "Unknown PowerShell edition: $($PSVersionTable.PSEdition)"
        }
        Write-Verbose "Checking for TcXaeMgmt module version $versionRange"
        $module = Get-Module -ListAvailable -Name TcXaeMgmt |
                  Where-Object { $_.Version -like $versionRange } |
                  Sort-Object Version -Descending |
                  Select-Object -First 1
        if (-not $module) {
            Write-Information "Installing TcXaeMgmt module version $versionRange"
            $specificVersion = if ($versionRange -eq '3.2.*') { '3.2.34' } else { '6.0.294' }
            Install-Module -Name TcXaeMgmt -RequiredVersion $specificVersion -Scope CurrentUser -Force -AcceptLicense -SkipPublisherCheck
            $module = Get-Module -ListAvailable -Name TcXaeMgmt |
                      Where-Object { $_.Version -like $versionRange } |
                      Sort-Object Version -Descending |
                      Select-Object -First 1
            if (-not $module) {
                throw "TcXaeMgmt module version $versionRange not found after installation."
            }
        }
        Import-Module TcXaeMgmt -RequiredVersion $module.Version -Force
        Write-Verbose "Loaded TcXaeMgmt version $($module.Version)"
    } catch {
        throw "Error in Test-TcXaeMgmtModule: $_"
    }
}

function Show-NoTargetsMessage {
    [CmdletBinding()]
    param(
        [int]$TimeoutSeconds
    )
    Clear-Host
    Write-Host "No target devices found." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Possible reasons:" -ForegroundColor Cyan
    Write-Host "1. No Beckhoff devices are currently connected or powered on" -ForegroundColor Gray
    Write-Host "2. Network connectivity issues"                              -ForegroundColor Gray
    Write-Host "3. ADS route discovery is not functioning"                 -ForegroundColor Gray
    Write-Host ""
    Write-Host "Automatically retrying in $TimeoutSeconds seconds..."         -ForegroundColor Green
    Write-Host "Press Enter to manually refresh now"                        -ForegroundColor Green
}

function Show-TableAndPrompt {
    [CmdletBinding()]
    param(
        [array]$RemoteRoutes
    )
    try {
        Clear-Host
        $table = for ($i = 0; $i -lt $RemoteRoutes.Count; $i++) {
            $route     = $RemoteRoutes[$i]
            $isUnknown = -not (
                $route.RTSystem -like "Win*"    -or
                $route.RTSystem -like "TcBSD*"  -or
                $route.RTSystem -like "TcRTOS*" -or
                $route.RTSystem -match "Linux"  -or
                $route.RTSystem -match "CE"
            )
            [PSCustomObject]@{
                Number    = $i + 1
                Name      = $route.Name
                IP        = $route.Address
                AMSNetID  = $route.NetId
                OS        = if ($isUnknown) { "$($route.RTSystem) (Unknown)" } else { $route.RTSystem }
                IsUnknown = $isUnknown
            }
        }
        foreach ($row in $table) {
            if ($row.IsUnknown) {
                Write-Host ("{0,2} {1,-20} {2,-15} {3,-20} {4}" -f $row.Number, $row.Name, $row.IP, $row.AMSNetID, $row.OS) -ForegroundColor DarkGray
            } else {
                Write-Host ("{0,2} {1,-20} {2,-15} {3,-20} {4}" -f $row.Number, $row.Name, $row.IP, $row.AMSNetID, $row.OS)
            }
        }
        Write-Host ""
        Write-Host "Select a target by entering its number (or type 'exit' to quit):" -ForegroundColor Cyan
    } catch {
        throw "Error in Show-TableAndPrompt: $_"
    }
}

function Get-DeviceManagerUrl {
    [CmdletBinding()]
    param(
        [psobject]$Route
    )
    switch ($Route.RTSystem) {
        {$_ -like "Win*"}    { return "https://$($Route.Address)/config" }
        {$_ -like "TcBSD*"}  { return "https://$($Route.Address)" }
        {$_ -like "TcRTOS*"} { return "http://$($Route.Address)/config" }
        {$_ -match "Linux"}  { return "https://$($Route.Address)" }
        {$_ -match "CE"}     { return "https://$($Route.Address)/config" }
        default                { throw "Unsupported RTSystem type: $($Route.RTSystem)" }
    }
}

function Show-ConnectionMenu {
    [CmdletBinding()]
    param(
        [psobject]$Route,
        [string]$DeviceManagerUrl,
        [string]$WinSCPPath,
        [string]$CerHostPath,
        [string]$AdminUserName,
        [string]$AdminPassword
    )
    try {
        Write-Host "Connection options for target '$($Route.Name)':" -ForegroundColor Cyan
        switch ($true) {
            ($Route.RTSystem -like "TcBSD*" -or $Route.RTSystem -match "Linux") {
                Write-Host "   1) Open Beckhoff Device Manager webpage ($DeviceManagerUrl)"
                Write-Host "   2) Start SSH session"
                Write-Host "   3) Open WinSCP connection"
                Write-Host "   4) Open both SSH session and WinSCP"
                break
            }
            ($Route.RTSystem -like "TcRTOS*") {
                Write-Host "   1) Open Beckhoff Device Manager webpage ($DeviceManagerUrl)"
                break
            }
            ($Route.RTSystem -like "Win*") {
                Write-Host "   1) Open Beckhoff Device Manager webpage ($DeviceManagerUrl)"
                Write-Host "   2) Start Remote Desktop session"
                break
            }
            ($Route.RTSystem -match "CE") {
                $isAvailable = Test-CERHostAvailability -IPAddress $Route.Address
                Write-Host "   1) Open Beckhoff Device Manager webpage ($DeviceManagerUrl)"
                if ($isAvailable) {
                    Write-Host "   2) Start CERHost Remote Desktop session" -ForegroundColor Green
                } else {
                    Write-Host "   2) Start CERHost Remote Desktop session" -ForegroundColor Red
                    Write-Host "      Note: CERHost port (987) is not open. Enable CERHost on the host PC." -ForegroundColor Yellow
                }
                break
            }
            default {
                throw "Unsupported RTSystem type: $($Route.RTSystem)"
            }
        }
        return Read-Host "Enter your choice"
    } catch {
        throw "Error in Show-ConnectionMenu: $_"
    }
}

function Execute-ConnectionChoice {
    [CmdletBinding()]
    param(
        [psobject]$Route,
        [string]$Choice,
        [string]$DeviceManagerUrl,
        [string]$WinSCPPath,
        [string]$CerHostPath,
        [string]$AdminUserName,
        [string]$AdminPassword
    )
    try {
        switch ($Choice) {
            '1' {
                Start-Process $DeviceManagerUrl
            }
            '2' {
                if ($Route.RTSystem -like "Win*") {
                    $cmdkeyCommand = "cmdkey /generic:TERMSRV/$($Route.Address) /user:$AdminUserName /pass:$AdminPassword"
                    cmd /c $cmdkeyCommand | Out-Null
                    $rdpFile = Join-Path $env:TEMP ("$($Route.Name -replace '[\\\/:*?"<>|]', '_').rdp")
                    @"
screen mode id:i:2
full address:s:$($Route.Address)
desktopwidth:i:1280
desktopheight:i:720
session bpp:i:32
smart sizing:i:1
"@ | Set-Content $rdpFile -Encoding ASCII
                    Start-Process mstsc.exe $rdpFile
                } elseif ($Route.RTSystem -like "TcBSD*" -or $Route.RTSystem -match "Linux") {
                    $sshCommand = if ($Route.RTSystem -match "Linux") {
                        "ssh -m hmac-sha2-512-etm@openssh.com $AdminUserName@$($Route.Address)"
                    } else {
                        "ssh $AdminUserName@$($Route.Address)"
                    }
                    Start-Process powershell.exe -ArgumentList '-NoExit','-Command',$sshCommand
                } elseif ($Route.RTSystem -match "CE") {
                    if (Test-Path $CerHostPath) {
                        Start-Process -FilePath $CerHostPath -ArgumentList $Route.Address
                    } else {
                        Write-Warning "CERHOST.exe not found at $CerHostPath"
                    }
                }
            }
            '3' {
                if ($Route.RTSystem -like "TcBSD*" -or $Route.RTSystem -match "Linux") {
                    try {
                        if ($Route.RTSystem -like "TcBSD*") {
                            & $WinSCPPath "sftp://${AdminUserName}:$AdminPassword@$($Route.Address)/" "/rawsettings" "SftpServer=doas /usr/libexec/sftp-server"
                        } else {
                            & $WinSCPPath "sftp://$($Route.Address)"
                        }
                    } catch {
                        Start-Process "https://winscp.net/eng/download.php"
                    }
                }
            }
            '4' {
                if ($Route.RTSystem -like "TcBSD*" -or $Route.RTSystem -match "Linux") {
                    Execute-ConnectionChoice -Route $Route -Choice '2' -DeviceManagerUrl $DeviceManagerUrl -WinSCPPath $WinSCPPath -CerHostPath $CerHostPath -AdminUserName $AdminUserName -AdminPassword $AdminPassword
                    Execute-ConnectionChoice -Route $Route -Choice '3' -DeviceManagerUrl $DeviceManagerUrl -WinSCPPath $WinSCPPath -CerHostPath $CerHostPath -AdminUserName $AdminUserName -AdminPassword $AdminPassword
                }
            }
            default { 
                Write-Warning "Invalid choice: $Choice. Valid options are 1, 2, 3, or 4."
                return
            }
        }
    } 
    catch {
        throw "Error in Execute-ConnectionChoice: $_"
    }
}

function Start-ADSDiscovery {
    [CmdletBinding()]
    param(
        [int]$TimeoutSeconds,
        [string]$WinSCPPath,
        [string]$CerHostPath,
        [string]$AdminUserName,
        [string]$AdminPassword
    )
    $prevTargetListJSON = ''
    do {
        try {
            $adsRoutes    = Get-AdsRoute -All
            $remoteRoutes = $adsRoutes | Where-Object { -not $_.IsLocal } | Sort-Object Name
            if ($remoteRoutes.Count -eq 0) {
                Show-NoTargetsMessage -TimeoutSeconds $TimeoutSeconds
                $selection = Read-InputWithTimeout -TimeoutSeconds $TimeoutSeconds -AllowRefresh
                if ($selection -eq 'refresh') { continue }
                if ($selection -eq 'exit')    { break }
                continue
            }
            $currentJSON = $remoteRoutes | ConvertTo-Json -Compress -Depth 5
            if ($prevTargetListJSON -ne $currentJSON) {
                Show-TableAndPrompt -RemoteRoutes $remoteRoutes
                $prevTargetListJSON = $currentJSON
            }
            $selection = Read-InputWithTimeout -TimeoutSeconds $TimeoutSeconds
            if ($selection -eq 'exit') { break }
            if (-not $selection) { continue }
            if ($selection -notmatch '^[0-9]+$' -or [int]$selection -lt 1 -or [int]$selection -gt $remoteRoutes.Count) {
                Write-Warning "Invalid selection. Continuing..."
                Start-Sleep -Seconds 1
                # Force screen refresh by clearing the previous JSON
                $prevTargetListJSON = ''
                continue
            }
            $selectedRoute = $remoteRoutes[[int]$selection - 1]
            if (-not (
                $selectedRoute.RTSystem -like "Win*"    -or
                $selectedRoute.RTSystem -like "TcBSD*"  -or
                $selectedRoute.RTSystem -like "TcRTOS*" -or
                $selectedRoute.RTSystem -match "Linux"  -or
                $selectedRoute.RTSystem -match "CE"
            )) {
                Write-Warning "Unsupported device type: $($selectedRoute.RTSystem)"
                Start-Sleep -Seconds 2
                # Force screen refresh by clearing the previous JSON
                $prevTargetListJSON = ''
                continue
            }
            
            Write-Information "Selected target: $($selectedRoute.Name) [$($selectedRoute.Address)] (AMS $($selectedRoute.NetId))"
            $deviceManagerUrl = Get-DeviceManagerUrl -Route $selectedRoute
            $choice           = Show-ConnectionMenu -Route $selectedRoute -DeviceManagerUrl $deviceManagerUrl -WinSCPPath $WinSCPPath -CerHostPath $CerHostPath -AdminUserName $AdminUserName -AdminPassword $AdminPassword
            Execute-ConnectionChoice -Route $selectedRoute -Choice $choice -DeviceManagerUrl $deviceManagerUrl -WinSCPPath $WinSCPPath -CerHostPath $CerHostPath -AdminUserName $AdminUserName -AdminPassword $AdminPassword
            
            # Automatically return to device list - clear the screen and force refresh
            $prevTargetListJSON = ''
        } catch {
            Write-Error "Error in discovery loop: $_"
        }
    } while ($true)
}

# Entry point
try {
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "SilentlyContinue"
    
    Test-TcXaeMgmtModule
    Start-ADSDiscovery -TimeoutSeconds $TimeoutSeconds -WinSCPPath $WinSCPPath -CerHostPath $CerHostPath -AdminUserName $AdminUserName -AdminPassword $AdminPassword
} catch {
    Write-Error "Fatal error: $_"
}