param(
    [string]$RouteName = "BTN-000f0t53",
    [string]$SourceFolder = "C:\tmp\Boot",
    [string]$DestinationFolder = "\TwinCAT\3.1\Boot\",
    [int]$SetToConfigAndRunMode = 1
)

$DebugPreference = "Continue" #or set to "SilentlyContinue" to ignore debug info

$route = Get-AdsRoute -Name $RouteName
$session = New-TcSession -Route $route -Port 10000 #TwinCAT system server 

# Set target to config mode.
if ($SetToConfigAndRunMode -eq 1) {
    Write-Debug "Setting target to config mode."
    try {
        Set-AdsState -SessionId $session.Id -State Config -Force 2>$null
    } catch [System.Exception] {
        if ($_.Exception.Message -like "*ClientRequestCancelled*") {
            Write-Debug "Encountered 'ClientRequestCancelled' error. Continuing..."
        } else {
            throw
        }
    }
    
    do {
        $result = Get-AdsState -SessionId $session.Id
        $state = $result.State
        Write-Host "Current state: $state"
    
        if ($state -eq "Config") {
            Write-Host "Device is now in config mode!"
            break
        }
    
        Start-Sleep -Seconds 5  # Wait for 5 seconds before checking again.
    } while ($true)
}

# Get a list of all files in a directory
write-Debug "Getting a list of files from $SourceFolder"
$files = Get-ChildItem -Path $SourceFolder -Recurse -Force -File

Write-Debug "copying files to the target"
foreach ($f in $files) {
    Copy-AdsFile -Path $f.FullName -SessionId $session.Id -Destination $f.FullName.Replace($SourceFolder, $DestinationFolder)  -Force -Upload
}

if ($SetToConfigAndRunMode -eq 1) {
    Write-Debug "Setting target to run mode."
    try {
        Set-AdsState -SessionId $session.Id -State Run -Force 2>$null
    } catch [System.Exception] {
        if ($_.Exception.Message -like "*ClientRequestCancelled*") {
            Write-Debug "Encountered 'ClientRequestCancelled' error. Continuing..."
        } else {
            throw
        }
    }
    
    do {
        $result = Get-AdsState -SessionId $session.Id
        $state = $result.State
        Write-Host "Current state: $state"
    
        if ($state -eq "Run") {
            Write-Host "Device is now in run mode!"
            break
        }
    
        Start-Sleep -Seconds 5  # Wait for 5 seconds before checking again.
    } while ($true)
}

Close-TcSession -InputObject $session
Write-Debug "Finished"
