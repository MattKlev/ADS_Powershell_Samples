# Suppress all progress displays globally
$ProgressPreference = 'SilentlyContinue'

# Get the AMS Net ID of the local system
$localNetID = Get-AmsNetId

# Perform a broadcast search for all available ADS routes
$adsRoutes = Get-AdsRoute -All

# Filter out the local system using AMS Net ID
$remoteRoutes = $adsRoutes | Where-Object { 
    $_.NetId -ne $localNetID 
}

# Check if any remote routes were found
if ($remoteRoutes.Count -eq 0) {
    Write-Output "No remote ADS routes found, excluding the local system."
    exit
}

# Display the list of targets in a table format
Write-Output "Found the following remote targets (excluding local system):"
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

# Ask the user to select a target
$selection = Read-Host "Select a target by entering the corresponding number"

# Validate the user input
if ($selection -match '^\d+$' -and $selection -gt 0 -and $selection -le $remoteRoutes.Count) {
    $selectedRoute = $remoteRoutes[$selection - 1]
    Write-Output "You selected: $($selectedRoute.Name), IP: $($selectedRoute.Address), AMS Net ID: $($selectedRoute.NetId), OS: $($selectedRoute.RTSystem)"

    if ($selectedRoute.RTSystem -like "Win*") {
        # If target is Windows, use Remote Desktop
        $cmdkeyCommand = "cmdkey /generic:TERMSRV/$($selectedRoute.Address) /user:Administrator /pass:1"
        Invoke-Expression $cmdkeyCommand > $null

        $remoteDesktopCommand = "mstsc /v:$($selectedRoute.Address)"
        Write-Output "Starting Remote Desktop session..."
        Invoke-Expression $remoteDesktopCommand

    } elseif ($selectedRoute.RTSystem -like "TcBSD*") {
        # If target is Tc/BSD, use SSH
        $sshCommand = "ssh Administrator@$($selectedRoute.Address)"
        Write-Output "Starting SSH session..."
        Invoke-Expression $sshCommand
    } else {
        Write-Output "Unsupported operating system type for remote connection: $($selectedRoute.RTSystem)"
    }

} else {
    Write-Output "Invalid selection. Please select a number from 1 to $($remoteRoutes.Count) and run the script again."
}
