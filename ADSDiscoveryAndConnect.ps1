# Function: Wait for key input up to a timeout (without printing a prompt).
function Read-InputWithTimeout {
    param(
        [int]$TimeoutSeconds = 10,
        [switch]$AllowRefresh = $false
    )
    
    $endTime = (Get-Date).AddSeconds($TimeoutSeconds)
    $inputString = ""
    while ((Get-Date) -lt $endTime) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($false)
            if ($key.Key -eq "Enter") { 
                if ($AllowRefresh) {
                    return "refresh"
                }
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
        [int]$Port = 987,
        [int]$TimeoutMilliseconds = 1000
    )
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $tcpClient.BeginConnect($IPAddress, $Port, $null, $null)
        $wait = $asyncResult.AsyncWaitHandle.WaitOne($TimeoutMilliseconds, $false)
        
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
        if ($tcpClient -ne $null) {
            $tcpClient.Close()
        }
    }
}

# Strict pre-flight check for PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "ERROR: This script requires PowerShell 7 or newer. Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    Write-Host "Opening PowerShell 7 download page..." -ForegroundColor Yellow
    Start-Process "https://github.com/PowerShell/PowerShell/releases"
    Read-Host "Press Enter to exit"
    exit 1
}

$ErrorActionPreference = "Stop"
$ProgressPreference = 'SilentlyContinue'

$psgModule = Get-Module -ListAvailable -Name PowerShellGet | 
             Sort-Object Version -Descending | Select-Object -First 1
if (-not $psgModule -or $psgModule.Version -lt [version]"2.2.5") {
    Write-Host "Your PowerShellGet module is outdated (version $($psgModule.Version) found)."
    Write-Host "Please update it to at least version 2.2.5 using:"
    Write-Host "    Install-Module -Name PowerShellGet -Force -AllowClobber"
    Read-Host "Press Enter to exit"
    exit 1
}

function Assert-ModuleInstalled {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Module '$ModuleName' is not installed. Attempting to install..."
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AcceptLicense
            Write-Host "Module '$ModuleName' installed successfully."
        }
        catch {
            Write-Host "Failed to install module '$ModuleName'. Error: $_"
            Read-Host "Press Enter to exit"
            exit 1
        }
    }
}

Assert-ModuleInstalled -ModuleName "TcXaeMgmt"
if (-not (Get-Module -ListAvailable -Name "TcXaeMgmt")) {
    Write-Host "Module 'TcXaeMgmt' is not properly installed."
    Write-Host "Please follow the instructions at:"
    Write-Host "https://infosys.beckhoff.com/content/1033/tc3_ads_ps_tcxaemgmt/5531473547.html?id=5297548616814834841"
    Read-Host "Press Enter to exit"
    exit 1
}

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

# Print the table and prompt once.
function Print-TableAndPrompt {
    param($remoteRoutes)
    # Build a table object.
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

# Main loop.
do {
    try {
        Import-Module TcXaeMgmt -Force
        $localNetID = Get-AmsNetId

        # Perform the broadcast search.
        $adsRoutes = Get-AdsRoute -All
        $remoteRoutes = $adsRoutes | Where-Object { $_.NetId -ne $localNetID } | Sort-Object -Property Name

        # If no remote routes found, show message and wait
        if ($remoteRoutes.Count -eq 0) {
            Show-NoTargetsMessage
            
            # Wait up to 10 seconds, return immediately if Enter is pressed
            $selection = Read-InputWithTimeout -TimeoutSeconds 10 -AllowRefresh
            
            # If manual refresh triggered, continue to next iteration
            if ($selection -eq "refresh") {
                continue
            }
            
            # If exit selected, break the loop
            if ($selection -eq "exit") { 
                break 
            }
            
            # If timeout occurred, continue to next iteration (which will retry search)
            continue
        }

        # Convert the list to JSON with a sufficient depth.
        $currentTargetListJSON = $remoteRoutes | ConvertTo-Json -Compress -Depth 5

        # If the device list has changed (or if first run), update display.
        if ($prevTargetListJSON -ne $currentTargetListJSON) {
            Clear-Host
            Print-TableAndPrompt $remoteRoutes
            $prevTargetListJSON = $currentTargetListJSON
        }
        
        # Wait up to 10 seconds for user input (without reprinting the prompt).
        $selection = Read-InputWithTimeout 10

        if ($selection -eq "exit") { break }
        if ($selection -eq "") {
            # No input: go back to checking broadcast search.
            continue
        }
        if ($selection -notmatch '^\d+$' -or $selection -le 0 -or $selection -gt $remoteRoutes.Count) {
            Write-Host "Invalid selection. Continuing..."
            Start-Sleep -Seconds 1
            continue
        }

        $selectedRoute = $remoteRoutes[$selection - 1]

        # Clear the screen and re-display the table (to keep the table visible)
        Clear-Host
        Print-TableAndPrompt $remoteRoutes
        Write-Host ""
        Write-Host "You selected: $($selectedRoute.Name), IP: $($selectedRoute.Address), AMS Net ID: $($selectedRoute.NetId), OS: $($selectedRoute.RTSystem)" -ForegroundColor Yellow

        # Determine connection method.
        if ($selectedRoute.RTSystem -like "Win*") {
            $deviceManagerURL = "https://$($selectedRoute.Address)/config"
        }
        elseif ($selectedRoute.RTSystem -like "TcBSD*") {
            $deviceManagerURL = "https://$($selectedRoute.Address)"   
        }
        elseif ($selectedRoute.RTSystem -like "TcRTOS*") {
            $deviceManagerURL = "http://$($selectedRoute.Address)/config"
        }
        elseif ($selectedRoute.RTSystem -like "*Linux*") {
            $deviceManagerURL = "https://$($selectedRoute.Address)"
        }
        elseif ($selectedRoute.RTSystem -like "*CE*") {
            $deviceManagerURL = "https://$($selectedRoute.Address)/config"
        }
        else {
            continue
        }

        # Present connection options.
        if ($selectedRoute.RTSystem -like "TcBSD*" -or $selectedRoute.RTSystem -like "*Linux*") {
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
        elseif ($selectedRoute.RTSystem -like "Win*")
        {
            Write-Host "Connection options for target '$($selectedRoute.Name)':" -ForegroundColor Cyan
            Write-Host "   1) Open Beckhoff Device Manager webpage ($deviceManagerURL)"
            Write-Host "   2) Start Remote Desktop session"
            $connectionChoice = Read-Host "Enter 1 or 2"
        }
        elseif ($selectedRoute.RTSystem -like "*CE*") {
            # Check CERHost port availability ONLY for CE devices
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
                
                # If user still wants to try CERHost, provide guidance
                if ($connectionChoice -eq "2") {
                    Write-Host "" -ForegroundColor Yellow
                    Write-Host "CERHost is not available. To enable:" -ForegroundColor Yellow
                    Write-Host "1. Open Beckhoff Device Manager webpage on the host pc" -ForegroundColor Gray
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
                    # Set credentials for the remote host using cmdkey.
                    $cmdkeyCommand = "cmdkey /generic:TERMSRV/$($selectedRoute.Address) /user:Administrator /pass:1"
                    Invoke-Expression $cmdkeyCommand | Out-Null

                    # Create a temporary RDP file with smart sizing enabled.
                    # Sanitize the computer name to generate a valid file name.
                    $targetName = ($selectedRoute.Name -replace '[\\\/:*?"<>|]', '_')
                    $rdpFile = Join-Path $env:TEMP ("$targetName.rdp")
                    $rdpContent = @"
screen mode id:i:2
full address:s:$($selectedRoute.Address)
desktopwidth:i:1280
desktopheight:i:720
session bpp:i:32
smart sizing:i:1
"@
                    $rdpContent | Set-Content -Path $rdpFile -Encoding ASCII

                    # Launch Remote Desktop session using the updated RDP file.
                    mstsc.exe $rdpFile
                }
                elseif ($selectedRoute.RTSystem -like "TcBSD*" -or $selectedRoute.RTSystem -like "*Linux*") {
                    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-Command", "ssh Administrator@$($selectedRoute.Address)"
                }
                elseif ($selectedRoute.RTSystem -like "*CE*") {
                    # Start CERHost for Windows CE device
                    try {
                        # Get the directory of the current script
                        $scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
                        $cerHostPath = Join-Path $scriptDirectory "CERHOST.exe"
                        
                        # Check if the file exists before trying to start it
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
                if ($selectedRoute.RTSystem -like "TcBSD*") {
                    $winscpExePath = "C:\Program Files (x86)\WinSCP\WinSCP.exe"
                    $target = $selectedRoute.Address
                    try {
                        & $winscpExePath "sftp://Administrator:1@$target/" "/rawsettings" "SftpServer=doas /usr/libexec/sftp-server"
                    }
                    catch {
                        Start-Process "https://winscp.net/eng/download.php"
                    }
                }
                elseif ($selectedRoute.RTSystem -like "*Linux*") {
                    $winscpExePath = "C:\Program Files (x86)\WinSCP\WinSCP.exe"
                    $target = $selectedRoute.Address
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
                    # Open SSH session
                    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-Command", "ssh Administrator@$($selectedRoute.Address)"
                    # Open WinSCP connection
                    $winscpExePath = "C:\Program Files (x86)\WinSCP\WinSCP.exe"
                    $target = $selectedRoute.Address
                    try {
                       & $winscpExePath "sftp://Administrator:1@$target/" "/rawsettings" "SftpServer=doas /usr/libexec/sftp-server"
                    }
                    catch {
                        Start-Process "https://winscp.net/eng/download.php"
                    }
                }
                elseif ($selectedRoute.RTSystem -like "*Linux*") {
                    # Open SSH session
                    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-Command", "ssh Administrator@$($selectedRoute.Address)"
                    # Open WinSCP connection
                    $winscpExePath = "C:\Program Files (x86)\WinSCP\WinSCP.exe"
                    $target = $selectedRoute.Address
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
        # Clear the screen and re-display only the table after executing the command.
        Clear-Host
        Print-TableAndPrompt $remoteRoutes
    }
    catch {
        Write-Host "An error occurred: $_"
    }
} while ($true)