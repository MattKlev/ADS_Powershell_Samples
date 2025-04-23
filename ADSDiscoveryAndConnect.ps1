# Function: Wait for key input up to a timeout (without printing a prompt).
function Read-InputWithTimeout {
    param(
        [int]$TimeoutSeconds = 10
    )
    
    $endTime = (Get-Date).AddSeconds($TimeoutSeconds)
    $inputString = ""
    while ((Get-Date) -lt $endTime) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($false)
            if ($key.Key -eq "Enter") { break }
            $inputString += $key.KeyChar
        }
        Start-Sleep -Milliseconds 100
    }
    return $inputString.Trim()
}

# Pre-flight checks.
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or newer. Current version: $($PSVersionTable.PSVersion)"
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Host "Attempting to install PowerShell 7 using winget..."
        try {
            winget install --id Microsoft.PowerShell -e
            Write-Host "PowerShell 7 installation initiated. Please restart the script using PowerShell 7."
        }
        catch {
            Write-Host "Winget encountered an error. Please install PowerShell 7 manually."
        }
    }
    else {
        Write-Host "Winget is not available. Opening the PowerShell 7 download page..."
        Start-Process "https://github.com/PowerShell/PowerShell/releases"
    }
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
        else {
            continue
        }

        # Present connection options.
        if ($selectedRoute.RTSystem -like "TcBSD*") {
            Write-Host "Connection options for target '$($selectedRoute.Name)':" -ForegroundColor Cyan
            Write-Host "   1) Open Beckhoff Device Manager webpage ($deviceManagerURL)"
            Write-Host "   2) Start SSH session"
            Write-Host "   3) Open WinSCP connection as Administrator with root privileges"
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
                elseif ($selectedRoute.RTSystem -like "TcBSD*") {
                    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-Command", "ssh Administrator@$($selectedRoute.Address)"
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
