$DebugPreference = "Continue" #or set to "SilentlyContinue" to ignore debug info

$route = Get-AdsRoute -Name "CX-C4ED06"
$SourceFolder = "C:\tmp\" #the local folder were the script is running from
$DestinationFolder = "C:\TwinCAT\3.1\Boot\" #target were the files are uploaded to
$SetToConfigAndRunMode = 0 #1 for true 0 for false - puts TwinCAT into config mode before transfering the files, and back into run mode when finsihed



$session = New-TcSession -Route $route -Port 10000 #TwinCAT system server 

#Set target to config mode.
if($SetToConfigAndRunMode)
{
 Write-Debug "Setting target to config mode."
    $setStateResult = Set-AdsState -SessionId $session.Id -State Config -Force
    
    if(!$setStateResult.Succeeded)
    {
        Write-Error "Failed to set target to config mode"
        return;
    }
}



#Get a list of all files in a directory

write-Debug "Getting a list of files from $SourceFolder"

$files = Get-ChildItem -Path $SourceFolder -Recurse -Force -File

Write-Debug "copying files to the target"
foreach ($f in $files)
{
   
   Copy-AdsFile -Path $f.FullName -SessionId $session.Id -Destination $f.FullName.Replace($SourceFolder,$DestinationFolder)  -Force -Upload  
   
}

if($SetToConfigAndRunMode)
{
    Write-Debug "Setting target to run mode."
    $setStateResult = Set-AdsState -SessionId $session.Id -State Run -Force
    
    if(!$setStateResult.Succeeded)
    {
        Write-Error "Failed to set target to run mode"
        return;
    }
}


Close-TcSession -InputObject $session

Write-Debug "Finished"
