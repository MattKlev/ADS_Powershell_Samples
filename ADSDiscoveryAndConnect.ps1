# Check if running PowerShell is version 7 or later.
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or newer. Current version: $($PSVersionTable.PSVersion)"
    
    # If winget is available, try to install PowerShell 7.
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

        Write-Host "Here is version 7.5.0"
        Start-Process "https://github.com/PowerShell/PowerShell/releases/tag/v7.5.0"
    }
    
    Read-Host "Press Enter to exit"
    exit 1
}

# Ensure all errors are treated as terminating.
$ErrorActionPreference = "Stop"

# Suppress all progress displays globally.
$ProgressPreference = 'SilentlyContinue'

# Verify that PowerShellGet is up to date.
$psgModule = Get-Module -ListAvailable -Name PowerShellGet | Sort-Object Version -Descending | Select-Object -First 1
if (-not $psgModule -or $psgModule.Version -lt [version]"2.2.5") {
    Write-Host "Your PowerShellGet module is outdated (version $($psgModule.Version) found)."
    Write-Host "Please update it to at least version 2.2.5 using:"
    Write-Host "    Install-Module -Name PowerShellGet -Force -AllowClobber"
    Read-Host "Press Enter to exit"
    exit 1
}

# Function to check and install a required module.
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

# Ensure the required module TcXaeMgmt is installed.
Assert-ModuleInstalled -ModuleName "TcXaeMgmt"

# Verify that TcXaeMgmt is now available.
if (-not (Get-Module -ListAvailable -Name "TcXaeMgmt")) {
    Write-Host "Module 'TcXaeMgmt' is not properly installed."
    Write-Host "Please ensure that you have followed the installation instructions from:"
    Write-Host "https://infosys.beckhoff.com/content/1033/tc3_ads_ps_tcxaemgmt/5531473547.html?id=5297548616814834841"
    Read-Host "Press Enter to exit"
    exit 1
}

try {
    # Import the required module.
    Import-Module TcXaeMgmt -Force

    # Get the AMS Net ID of the local system.
    $localNetID = Get-AmsNetId

    # Perform a broadcast search for all available ADS routes.
    $adsRoutes = Get-AdsRoute -All

    # Filter out the local system using AMS Net ID.
    $remoteRoutes = $adsRoutes | Where-Object { 
        $_.NetId -ne $localNetID 
    }

    # Check if any remote routes were found.
    if ($remoteRoutes.Count -eq 0) {
        Write-Output "No remote ADS routes found."
        Read-Host "Press Enter to exit"
        exit
    }

    # Display the list of targets in a table format.
    Write-Output "Found the following remote targets:"
    $table = @()
    for ($i = 0; $i -lt $remoteRoutes.Count; $i++) {
        $route = $remoteRoutes[$i]
        $table += [PSCustomObject]@{
            Number   = $i + 1
            Name     = $route.Name
            IP       = $route.Address
            AMSNetID = $route.NetId
            OS       = $route.RTSystem
        }
    }
    $table | Format-Table -AutoSize

    # Ask the user to select a target.
    $selection = Read-Host "Select a target by entering the corresponding number"

    # Validate the user input.
    if ($selection -match '^\d+$' -and $selection -gt 0 -and $selection -le $remoteRoutes.Count) {
        $selectedRoute = $remoteRoutes[$selection - 1]
        Write-Output "You selected: $($selectedRoute.Name), IP: $($selectedRoute.Address), AMS Net ID: $($selectedRoute.NetId), OS: $($selectedRoute.RTSystem)"

        # Determine URL and default connection based on the operating system.
        if ($selectedRoute.RTSystem -like "Win*") {
            $deviceManagerURL = "https://$($selectedRoute.Address)/config"
            $defaultAction = "RDP"
        } elseif ($selectedRoute.RTSystem -like "TcBSD*") {
            $deviceManagerURL = "https://$($selectedRoute.Address)"
            $defaultAction = "SSH"
        } else {
            Write-Output "Unsupported operating system type for remote connection: $($selectedRoute.RTSystem)"
            exit 1
        }

        # Ask user for connection type.
        Write-Output "Connection options for target '$($selectedRoute.Name)':"
        Write-Output "   1) Open Beckhoff Device Manager webpage ($deviceManagerURL)"
        Write-Output "   2) Start default remote connection ($defaultAction)"
        $connectionChoice = Read-Host "Enter 1 or 2"

        if ($connectionChoice -eq "1") {
            Write-Output "Opening Beckhoff Device Manager webpage at $deviceManagerURL ..."
            Start-Process $deviceManagerURL
        }
        elseif ($connectionChoice -eq "2") {
            if ($selectedRoute.RTSystem -like "Win*") {
                # If target is Windows, use Remote Desktop.
                $cmdkeyCommand = "cmdkey /generic:TERMSRV/$($selectedRoute.Address) /user:Administrator /pass:1"
                Invoke-Expression $cmdkeyCommand > $null

                $remoteDesktopCommand = "mstsc /v:$($selectedRoute.Address)"
                Write-Output "Starting Remote Desktop session..."
                Invoke-Expression $remoteDesktopCommand
            } elseif ($selectedRoute.RTSystem -like "TcBSD*") {
                # If target is Tc/BSD, use SSH.
                $sshCommand = "ssh Administrator@$($selectedRoute.Address)"
                Write-Output "Starting SSH session..."
                Invoke-Expression $sshCommand
            }
        }
        else {
            Write-Output "Invalid selection. Please choose either 1 or 2."
            Read-Host "Press Enter to exit"
        }

    } else {
        Write-Output "Invalid selection. Please select a number from 1 to $($remoteRoutes.Count) and run the script again."
        Read-Host "Press Enter to exit"
    }
}
catch {
    Write-Host "An error occurred: $_"
    Read-Host "Press Enter to exit"
    exit 1
}
