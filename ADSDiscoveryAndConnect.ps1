# Store the current PowerShell version for debugging purposes
$CurrentPSVersion = $PSVersionTable.PSVersion;
Write-Host "Powershell Version: $CurrentPSVersion" -ForegroundColor Green

# Function: Wait for key input up to a timeout (without printing a prompt).
function Read-InputWithTimeout {
    param(
        [int]$TimeoutSeconds = 10,
        [switch]$AllowRefresh = $false
    )

    $endTime     = (Get-Date).AddSeconds($TimeoutSeconds)
    $inputString = ""
    while ((Get-Date) -lt $endTime) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($false)
            if ($key.Key -eq "Enter") {
                if ($AllowRefresh) { return "refresh" }
                break
            }
            $inputString += $key.KeyChar
        }
        Start-Sleep -Milliseconds 100
    }
    return $inputString.Trim()
}

# Function to check CERHost port availability quickly
function Test-CERHostAvailability {
    param(
        [string]$IPAddress,
        [int]$Port               = 987,
        [int]$TimeoutMilliseconds = 1000
    )
    try {
        $tcpClient   = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $tcpClient.BeginConnect($IPAddress, $Port, $null, $null)
        $wait        = $asyncResult.AsyncWaitHandle.WaitOne($TimeoutMilliseconds, $false)

        if ($wait) {
            $tcpClient.EndConnect($asyncResult)
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        return $false
    }
    finally {
        if ($tcpClient) { $tcpClient.Close() }
    }
}

$ErrorActionPreference = "Stop"
$ProgressPreference    = 'SilentlyContinue'

# Function to test and load the TcXaeMgmt module
# This function checks the PowerShell edition and installs the appropriate version of the TcXaeMgmt module if it is not already available.
function Test-TcXaeMgmtModule {
    # Determine PowerShell edition and required module version
    if ($PSVersionTable.PSEdition -eq 'Desktop') {
        $versionRange = '3.2.*'
    }
    elseif ($PSVersionTable.PSEdition -eq 'Core') {
        $versionRange = '6.*'
    }
    else {
        Write-Host "Unknown PowerShell edition. Exiting." -ForegroundColor Red
        exit 1
    }

    # Find the latest available version in the desired range
    $module = Get-Module -ListAvailable -Name TcXaeMgmt | Where-Object { $_.Version -like $versionRange } | Sort-Object Version -Descending | Select-Object -First 1

    if (-not $module) {
        Write-Host "TcXaeMgmt module version $versionRange not found. Installing..." -ForegroundColor Yellow
        try {
            # Install the latest version in the range from PSGallery
            Install-Module -Name TcXaeMgmt -RequiredVersion ($versionRange -replace '\*','') -Scope CurrentUser -Force -AcceptLicense -SkipPublisherCheck
            # Try to find it again after install
            $module = Get-Module -ListAvailable -Name TcXaeMgmt | Where-Object { $_.Version -like $versionRange } | Sort-Object Version -Descending | Select-Object -First 1
        }
        catch {
            Write-Host "Failed to install TcXaeMgmt ${versionRange}: $_" -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
    }

    if ($module) {
        Import-Module TcXaeMgmt -RequiredVersion $($module.Version) -Force
        $imported = Get-Module TcXaeMgmt
        if ($imported -and $imported.Version -eq $module.Version) {
            Write-Host "TcXaeMgmt $($imported.Version) loaded." -ForegroundColor Green
        }
        else {
            Write-Host "Failed to load TcXaeMgmt $($module.Version)." -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
    }
    else {
        Write-Host "TcXaeMgmt module not found after attempted install." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Check and load the correct TcXaeMgmt module version
Test-TcXaeMgmtModule

# This variable holds the JSON version of the previous device list.
$prevTargetListJSON = $null

# Function to display no targets found message
function Show-NoTargetsMessage {
    Clear-Host
    Write-Host "No target devices found." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Possible reasons:" -ForegroundColor Cyan
    Write-Host "1. No Beckhoff devices are currently connected or powered on" -ForegroundColor Gray
    Write-Host "2. Network connectivity issues" -ForegroundColor Gray
    Write-Host "3. ADS route discovery is not functioning" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Troubleshooting tips:" -ForegroundColor Cyan
    Write-Host "- Ensure devices are powered on and connected to the network" -ForegroundColor Gray
    Write-Host "- Check network settings and firewall configurations" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Automatically retrying in 10 seconds..." -ForegroundColor Green
    Write-Host "Press Enter to manually refresh now" -ForegroundColor Green
}

# Print the table and prompt once
function Show-TableAndPrompt {
    param($remoteRoutes)
    $table = for ($i = 0; $i -lt $remoteRoutes.Count; $i++) {
        $route = $remoteRoutes[$i]
        [PSCustomObject]@{
            Number   = $i + 1
            Name     = $route.Name
            IP       = $route.Address
            AMSNetID = $route.NetId
            OS       = $route.RTSystem
        }
    }
    $table | Format-Table -AutoSize
    Write-Host ""
    Write-Host "Select a target by entering its number (or type 'exit' to quit):" -ForegroundColor Cyan
}

# Main loop to discover and connect to Beckhoff devices
do {
    try {
      
        # Use broadcast to get all routes and filter out "Local"
        $adsRoutes    = Get-AdsRoute -Broadcast 
        $remoteRoutes = $adsRoutes |
                        Where-Object { $_.IsLocal -ne $True } |
                        Sort-Object Name

        if ($remoteRoutes.Count -eq 0) {
            Show-NoTargetsMessage
            $selection = Read-InputWithTimeout -TimeoutSeconds 10 -AllowRefresh
            if ($selection -eq "refresh") { continue }
            if ($selection -eq "exit")    { break }
            continue
        }

        $currentTargetListJSON = $remoteRoutes | ConvertTo-Json -Compress -Depth 5
        if ($prevTargetListJSON -ne $currentTargetListJSON) {
            Clear-Host
            Show-TableAndPrompt $remoteRoutes
            $prevTargetListJSON = $currentTargetListJSON
        }

        $selection = Read-InputWithTimeout 10

        if ($selection -eq "exit")                     { break }
        if ($selection -eq "")                         { continue }
        if ($selection -notmatch '^\d+$' `
            -or $selection -lt 1 `
            -or $selection -gt $remoteRoutes.Count) {
            Write-Host "Invalid selection. Continuing..." -ForegroundColor Red
            Start-Sleep -Seconds 1
            continue
        }

        $selectedRoute = $remoteRoutes[$selection - 1]
        Clear-Host
        Show-TableAndPrompt $remoteRoutes
        Write-Host ""
        Write-Host "You selected: $($selectedRoute.Name) [$($selectedRoute.Address)] (AMS $($selectedRoute.NetId))" -ForegroundColor Yellow

        # Determine device manager URL
        if    ($selectedRoute.RTSystem -like "Win*")    { $deviceManagerURL = "https://$($selectedRoute.Address)/config" }
        elseif ($selectedRoute.RTSystem -like "TcBSD*") { $deviceManagerURL = "https://$($selectedRoute.Address)" }
        elseif ($selectedRoute.RTSystem -like "TcRTOS*"){ $deviceManagerURL = "http://$($selectedRoute.Address)/config" }
        elseif ($selectedRoute.RTSystem -match "Linux") { $deviceManagerURL = "https://$($selectedRoute.Address)" }
        elseif ($selectedRoute.RTSystem -match "CE")    { $deviceManagerURL = "https://$($selectedRoute.Address)/config" }
        else { continue }

        # Present connection options
        if ($selectedRoute.RTSystem -like "TcBSD*" -or $selectedRoute.RTSystem -match "Linux") {
            Write-Host "Connection options for target '$($selectedRoute.Name)':" -ForegroundColor Cyan
            Write-Host "   1) Open Beckhoff Device Manager webpage ($deviceManagerURL)"
            Write-Host "   2) Start SSH session"
            Write-Host "   3) Open WinSCP connection"
            Write-Host "   4) Open both SSH session and WinSCP"
            $connectionChoice = Read-Host "Enter 1, 2, 3, or 4"
        }
        elseif ($selectedRoute.RTSystem -like "TcRTOS*") {
            Write-Host "Connection options for target '$($selectedRoute.Name)':" -ForegroundColor Cyan
            Write-Host "   1) Open Beckhoff Device Manager webpage ($deviceManagerURL)"
            $connectionChoice = Read-Host "Enter 1"
        }
        elseif ($selectedRoute.RTSystem -like "Win*") {
            Write-Host "Connection options for target '$($selectedRoute.Name)':" -ForegroundColor Cyan
            Write-Host "   1) Open Beckhoff Device Manager webpage ($deviceManagerURL)"
            Write-Host "   2) Start Remote Desktop session"
            $connectionChoice = Read-Host "Enter 1 or 2"
        }
        elseif ($selectedRoute.RTSystem -match "CE") {
            $cerHostAvailable = Test-CERHostAvailability -IPAddress $selectedRoute.Address

            Write-Host "Connection options for target '$($selectedRoute.Name)':" -ForegroundColor Cyan
            Write-Host "   1) Open Beckhoff Device Manager webpage ($deviceManagerURL)"
            if ($cerHostAvailable) {
                Write-Host "   2) Start CERHost Remote Desktop session" -ForegroundColor Green
                $connectionChoice = Read-Host "Enter 1 or 2"
            }
            else {
                Write-Host "   2) Start CERHost Remote Desktop session" -ForegroundColor Red
                Write-Host "      Note: CERHost port (987) is not open. Enable CERHost on the host PC." -ForegroundColor Yellow
                $connectionChoice = Read-Host "Enter 1 (or confirm to continue without CERHost)"
                if ($connectionChoice -eq "2") {
                    Write-Host "" -ForegroundColor Yellow
                    Write-Host "CERHost is not available. To enable:" -ForegroundColor Yellow
                    Write-Host "1. Open Beckhoff Device Manager webpage on the host PC" -ForegroundColor Gray
                    Write-Host "2. Navigate to Device -> Boot -> Remote Display -> Set to ON" -ForegroundColor Gray
                    Read-Host "Press Enter to continue"
                    $connectionChoice = "1"
                }
            }
        }
        else {
            continue
        }

        switch ($connectionChoice) {
            "1" {
                Start-Process $deviceManagerURL
            }
            "2" {
                if ($selectedRoute.RTSystem -like "Win*") {
                    $cmdkeyCommand = "cmdkey /generic:TERMSRV/$($selectedRoute.Address) /user:Administrator /pass:1"
                    Invoke-Expression $cmdkeyCommand | Out-Null

                    $targetName = ($selectedRoute.Name -replace '[\\/:*?"<>|]', '_')
                    $rdpFile    = Join-Path $env:TEMP ("$targetName.rdp")
                    $rdpContent = @"
screen mode id:i:2
full address:s:$($selectedRoute.Address)
desktopwidth:i:1280
desktopheight:i:720
session bpp:i:32
smart sizing:i:1
"@
                    $rdpContent | Set-Content -Path $rdpFile -Encoding ASCII
                    mstsc.exe $rdpFile
                }
                elseif ($selectedRoute.RTSystem -like "TcBSD*") {
                    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-Command", "ssh Administrator@$($selectedRoute.Address)"
                }
                elseif ($selectedRoute.RTSystem -match "Linux") {
                    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-Command", "ssh -m hmac-sha2-512-etm@openssh.com Administrator@$($selectedRoute.Address)"
                }
                elseif ($selectedRoute.RTSystem -match "CE") {
                    try {
                        $scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
                        $cerHostPath     = Join-Path $scriptDirectory "CERHOST.exe"
                        if (Test-Path $cerHostPath) {
                            Start-Process $cerHostPath -ArgumentList "$($selectedRoute.Address)"
                        }
                        else {
                            Write-Host "Error: CERHOST.exe not found in script directory." -ForegroundColor Red
                            Write-Host "Expected location: $cerHostPath" -ForegroundColor Yellow
                        }
                    }
                    catch {
                        Write-Host "Error starting CERHost: $_" -ForegroundColor Red
                    }
                }
            }
            "3" {
                $winscpExePath = "C:\Program Files (x86)\WinSCP\WinSCP.exe"
                $target        = $selectedRoute.Address
                if ($selectedRoute.RTSystem -like "TcBSD*") {
                    try {
                        & $winscpExePath "sftp://Administrator:1@$target/" "/rawsettings" "SftpServer=doas /usr/libexec/sftp-server"
                    }
                    catch {
                        Start-Process "https://winscp.net/eng/download.php"
                    }
                }
                elseif ($selectedRoute.RTSystem -match "Linux") {
                    try {
                        & $winscpExePath "sftp://$target"
                    }
                    catch {
                        Start-Process "https://winscp.net/eng/download.php"
                    }
                }
            }
            "4" {
                if ($selectedRoute.RTSystem -like "TcBSD*") {
                    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-Command", "ssh Administrator@$($selectedRoute.Address)"
                    $winscpExePath = "C:\Program Files (x86)\WinSCP\WinSCP.exe"
                    $target        = $selectedRoute.Address
                    try {
                        & $winscpExePath "sftp://Administrator:1@$target/" "/rawsettings" "SftpServer=doas /usr/libexec/sftp-server"
                    }
                    catch {
                        Start-Process "https://winscp.net/eng/download.php"
                    }
                }
                elseif ($selectedRoute.RTSystem -match "Linux") {
                    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-Command", "ssh -m hmac-sha2-512-etm@openssh.com Administrator@$($selectedRoute.Address)"
                    $winscpExePath = "C:\Program Files (x86)\WinSCP\WinSCP.exe"
                    $target        = $selectedRoute.Address
                    try {
                        & $winscpExePath "sftp://$target"
                    }
                    catch {
                        Start-Process "https://winscp.net/eng/download.php"
                    }
                }
            }
            default { }
        }

        Clear-Host
        Show-TableAndPrompt $remoteRoutes
    }
    catch {
        Write-Host "An error occurred: $_" -ForegroundColor Red
    }
} while ($true)


