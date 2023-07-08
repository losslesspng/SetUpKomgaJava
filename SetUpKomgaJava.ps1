############################################################################################################
# This script will automatically download:                                                                 #
#     • Java (Adoptium Temurin 17)                                                                         #
#     • NSSM (Non Sucky Service Manager                                                                    #
#     • Komga                                                                                              #
# It will then:                                                                                            #
#     • Create a PowerShell script that can update a komga service                                         #
#     • Use NSSM to create the Komga service and set it to run automatically with Windows                  #
#     • Start the service                                                                                  #
#     • Create a Scheduled Task and a shortcut to allow you to easily check for updates                    #
# Please note that you will need to fill in a few variables beforehand, namely:                            #
#     • $runDIR                                                                                            #
#     • $serviceName                                                                                       #
#     • $javaArguments                                                                                     #
#     • $backupDIR                                                                                         #
# Please read over the code below and make the changes necessary, OR feel free to leave them as is         #
# The settings provided below are Works On My Machine certified and any edits are purely for your own sake.#
# Please use the following line to run this file:                                                          #
# Powershell.exe -ExecutionPolicy Bypass -NoProfile -File .\SetUpKomgaJava.ps1                             #
# This will run the script without requiring you to modify the ExecutionPolicy for your whole computer     #
#                                                                                                          #
#                                                                                                          #
# This script is lovingly provided by png, with help from Diesel and gotson.                               #
############################################################################################################

Write-Host "Thanks for checking out Komga, hope you love it!"

################
# Assign These #
################
# The directory that you will keep the komga jar and everything else that this scripts needs to run
$runDIR = "C:\Utilities\komga\"
# The name that will be given to the komga service
$serviceName = "komga"
# The java arguments that will be used to launch komga. Runs with 4gb of RAM unless modified.
$javaArguments = "-jar -Xmx4g"
# The name of the backup directory - this will hold the current version of komga and two older versions as backup, just in case
$backupDIR = "previous_versions"

# Stop the Komga service if it exists
If ( Get-Service -Name "$serviceName" -ErrorAction 'SilentlyContinue' )
{
    Stop-Service "$serviceName"
}


################################
# Adoptium Temurin 17 LTS Java #
################################
# Java file pattern, download URI, and local save path for the zip file
$assetPattern = "OpenJDK17U-jre_x64_windows_hotspot_17.0.3_7.zip"
$downloadURI = "https://github.com/adoptium/temurin17-binaries/releases/download/jdk-17.0.3%2B7/OpenJDK17U-jre_x64_windows_hotspot_17.0.3_7.zip"
$localZip = ( Join-Path -Path "$runDIR" -ChildPath $assetPattern )

# Make sure that the running directory exists, or create it
If ( !( Test-Path -Path "$runDIR" ) ) 
{ 
   Write-Host Creating directory $runDIR
   New-Item -Path "$runDIR" -ItemType Directory 
}

# Delete the java folder if it exists already
If ( Test-Path -Path "$runDIR\jdk*" )
{
    Write-Host Deleting existing java folder
    Remove-Item "$runDIR\jdk*.*" -Recurse -Force
}

# Download the Java zip file, extract it, delete it, and capture the name of the folder that is extracted
Write-Host Starting download of $assetPattern
Start-BitsTransfer -Source $downloadURI -Destination $runDIR
Write-Host Extracting and deleting zip file $localZip
Expand-Archive -LiteralPath "$localZip" -DestinationPath "$runDIR"
Remove-Item -Path "$localZip"
# Setting the directory that java extracted to, this will get used later
$javaDIR = Get-ChildItem "$runDIR" | Where-Object { $_.PSIsContainer } | Sort CreationTime -Descending | Select -F 1
$javaDIR = ( Join-path -Path "$runDIR" -ChildPath $javaDIR.Name )

#########
# Komga #
#########
# Set the path for the backup directory, and create it if it doesn't already exist
$backupDIR = ( Join-Path -Path $runDIR -ChildPath $backupDIR )
If ( !( Test-Path -Path "$backupDIR" ) ) { New-Item -Path "$backupDIR" -ItemType Directory }

# Download code for komga, grab the latest release
If ( Test-Path -Path "$runDIR\komga.jar" )
{
    Remove-Item "$runDIR\komga.jar"
}
$repoName = "gotson/komga"
$assetPattern = "komga-*.jar"
$releasesURI = "https://api.github.com/repos/$repoName/releases/latest"
$asset = ( Invoke-WebRequest $releasesURI -UseBasicParsing | ConvertFrom-Json ).Assets | Where-Object Name -Like $assetPattern
$downloadURI = $asset.Browser_Download_URL
$latest = [string]$downloadURI.Split('/')[-1]
Write-Host Starting download of $latest
Start-BitsTransfer -Source "$downloadURI" -Destination "$runDIR"
# Make a copy of the jar file, just in case, and then rename it
Copy-Item -Path ( Join-Path -Path $runDIR -ChildPath "$assetPattern" ) -Destination "$backupDIR"
Get-ChildItem -Path ( Join-Path -Path $runDIR -ChildPath $assetPattern ) | Rename-Item -NewName "komga.jar"

# Delete the files if they already exist
If ( Test-Path -Path ( Join-Path -Path "$runDIR" -ChildPath "UpdateKomga.ps1" ) )
{
    Remove-Item -Path ( Join-Path -Path "$runDIR" -ChildPath "UpdateKomga.ps1" )
}
If ( Test-Path -Path ( Join-Path -Path "$runDIR" -ChildPath "KomgaService.bat" ) )
{
    Remove-Item -Path ( Join-Path -Path "$runDIR" -ChildPath "KomgaService.bat" )
}

# Create the PowerShell scripts that will be used to run and update komga
New-Item -ItemType File -Path ( Join-Path -Path "$runDIR" -ChildPath "UpdateKomga.ps1" )
New-Item -ItemType File -Path ( Join-Path -Path "$runDIR" -ChildPath "KomgaService.bat" )

$UpdateKomgaPS1 = @'
# Download code for komga
$repoName = "gotson/komga"
$assetPattern = "komga-*.jar"
$runDIR = "@run"
$backupDIR = "@back"
$serviceName = "@serv"
# Create the running directory, in case it doesn't exist already
If ( !( Test-Path -Path "$runDIR" ) ) { New-Item -Path "$runDIR" -ItemType Directory }
$latest = Get-ChildItem -Path "$backupDIR" -Filter "*.jar" | Sort-Object LastAccessTime -Descending | Select-Object -First 1
$releasesURI = "https://api.github.com/repos/$repoName/releases/latest"
$asset = ( Invoke-WebRequest $releasesURI  -UseBasicParsing | ConvertFrom-Json ).Assets | Where-Object Name -Like $assetPattern
$downloadURI = $asset.Browser_Download_URL

# Check and see if the latest backed up version is the same as the latest posted version
If ( !( [string]$downloadURI.Split('/')[-1] -eq [string]$latest.Name ) )
{
    $latest = [string]$downloadURI.Split('/')[-1]
    Write-Host Update Found! Now updating to $latest
    If ( !( Test-Path -Path "$backupDIR" ) ) { New-Item -Path "$backupDIR" -ItemType Directory }
    Get-ChildItem -Path "$backupDIR" -Filter "*.jar" | Sort-Object LastAccessTime -Descending | Select-Object -Skip 2 | Remove-Item
    Stop-Service -Name "$serviceName"
    Start-BitsTransfer -Source "$downloadURI" -Destination "$runDIR"
    Copy-Item -Path ( Join-Path -Path $runDIR -ChildPath "$assetPattern" ) -Destination "$backupDIR"
    If ( Test-Path -Path ( Join-Path -Path $runDIR -ChildPath "komga.jar" ) ) { Remove-Item -Path ( Join-Path -Path "$runDIR" -ChildPath "komga.jar" ) }
    Get-ChildItem -Path ( Join-Path -Path $runDIR -ChildPath $assetPattern ) | Rename-Item -NewName "komga.jar"
    Start-Service -Name "$serviceName"
    Timeout /T 2
} Else {
    Write-Host No update found!
    Timeout /T 2
}
'@

# Replace the placeholders with the values assigned for use in the rest of the script
$UpdateKomgaPS1 = $UpdateKomgaPS1.Replace( "@run", $runDIR )
$UpdateKomgaPS1 = $UpdateKomgaPS1.Replace( "@back", $backupDIR )
$UpdateKomgaPS1 = $UpdateKomgaPS1.Replace( "@serv", $serviceName )

# Write the file contents
Set-Content -Path ( Join-Path -Path "$runDIR" -ChildPath "UpdateKomga.PS1" ) $UpdateKomgaPS1

# Ahhhh a bat, hopefully not a vampire one
$KomgaServiceBAT = @'
@echo off

@java @arg "@jar"
'@

# Replace the placeholders with the values
$KomgaServiceBAT = $KomgaServiceBAT.Replace( "@java", ( Join-Path -Path $javaDIR -ChildPath "bin" | Join-Path -ChildPath "java" ) )
$KomgaServiceBAT = $KomgaServiceBAT.Replace( "@arg", $javaArguments )
$KomgaServiceBAT = $KomgaServiceBAT.Replace( "@jar", ( Join-Path -Path $runDIR -ChildPath "komga.jar" ) )

# Write the file contents
Set-Content -Path ( Join-Path -Path "$runDIR" -ChildPath "KomgaService.bat" ) $KomgaServiceBAT

##################
# Scheduled Task #
##################
# Create the scheduled task used to update komga
# I really, really wanted this to run in the background but MICROSOFT thinks that ain't cool
# Powershell can't make calls to download files in scripts that aren't focused, so.. sorry but this will launch a shell
$scheduledTaskObject = New-Object -ComObject Schedule.Service
$scheduledTaskObject.Connect()
If ( Get-ScheduledTask -TaskName "Update Komga" -EA 0 ) { Unregister-ScheduledTask -TaskName "Update Komga" -Confirm:$false }
$UpdateKomgaPS1 = ( Join-Path -Path "$runDIR" -ChildPath "UpdateKomga.ps1" )
$action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument "-ExecutionPolicy ByPass -NoProfile -File $UpdateKomgaPS1"
Register-ScheduledTask -Action $action -TaskName "Update Komga" -Description "Check for komga updates" -ErrorAction 'SilentlyContinue'

If ( !( Test-Path "$runDIR\Check For Updates.lnk" ) )
{
    # Build a shortcut to the scheduled task
    $TargetPath = "C:\Windows\System32\schtasks.exe"
    $Arguemnts = '/run /tn "Update Komga"'
    $Destination = ( Join-Path -Path "$runDir" -ChildPath "Check For Updates.lnk" )

    $Shell = New-Object -ComObject WScript.Shell
    $Shortcut = $Shell.CreateShortcut( "$Destination" )
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.Arguments = $Arguemnts
    $Shortcut.Save()
}

########
# NSSM #
########

# Check if NSSM already exists... if it does, this script has probably already been run before. 
If ( !( Test-Path -Path "$runDIR\nssm" ) )
{
# Put together the user credentials needed for NSSM
Write-Host Credentials are required for the NSSM service
$username = "$env:USERDOMAIN\$env:USERNAME"
$credentials = Get-Credential -Credential $username
$password = $credentials.GetNetworkCredential().Password
$sid = ( New-Object System.Security.Principal.NTAccount( $username ) ).Translate( [System.Security.Principal.SecurityIdentifier] ).Value

# NSSM file name, download URI, and local save path for the zip file
$assetPattern = "nssm-2.24.zip"
$downloadURI = "https://nssm.cc/release/$assetPattern"
$localZip = ( Join-Path -Path "$runDIR" -ChildPath $assetPattern )

# Download the NSSM zip file, extract it, move out the x64 version of NSSM, and delete the files not needed
Write-Host Starting download of $assetPattern
Write-Host You will be prompted for administrator rights once it is ready to configure. 
Start-BitsTransfer -Source $downloadURI -Destination "$runDIR"
Write-Host Extracting and deleting zip file $localZip
Expand-Archive -LiteralPath "$localZip" -DestinationPath "$runDIR"
Remove-Item -Path "$localZip"

# It's not the most pleasant code ever, but hey, it works
# Make a new nssm directory and keep the 64bit nssm exe in it, it's all we need
$nssmDIR = Get-ChildItem "$runDIR" | Where-Object { $_.PSIsContainer } | Sort CreationTime -Descending | Select -F 1
$nssmDIR = ( Join-Path -Path "$runDIR" -ChildPath $nssmDIR.Name )
$nssmDIR2 = ( Join-Path -Path $runDIR -ChildPath "nssm" )
If ( !( Test-Path -Path "$nssmDIR2" ) ) { New-Item -Path "$nssmDIR2" -ItemType Directory }
Move-Item -Path ( Join-Path -Path "$nssmDIR" -ChildPath "win64" | Join-Path -ChildPath "nssm.exe" ) -Destination "$nssmDIR2"
Remove-Item -Path "$nssmDIR" -Recurse
$nssmDIR = ( Join-Path -Path $nssmDIR2 -ChildPath "nssm.exe" )

# Gonna need this, at least temporarily. 
New-Item -ItemType File -Path ( Join-Path -Path "$runDIR" -ChildPath "nssm.bat" )
New-Item -ItemType File -Path ( Join-Path -Path "$runDIR" -ChildPath "SetSCPerms.ps1" )

# Create a Batch file to create the service with NSSM. The following steps require a new terminal window to be opened and will prompt for admin.
$nssmBatch = @'
@echo off
@nssm install "@serviceName" @komgaServiceBAT
@nssm set "@serviceName" DisplayName "Komga Service"
@nssm set "@serviceName" Description "Service created with love using SetUpKomgaJava.ps1 from losslesspng and NSSM"
@nssm set "@serviceName" ObjectName "@user" "@pass"
sc.exe sdshow "@serviceName" >@permTXT
PowerShell.exe -ExecutionPolicy Bypass -NoProfile -File @permPS1 -Wait
set /p pe= < @permTXT
sc.exe sdset "@serviceName" "%pe%"
DEL "@permTXT" 
DEL "@permPS1" 
DEL %0
'@

# This is some fancy shit and it's just easier to do this in Powershell, all it needs to do is read the output from sc,
# inject the SID we grabbed earlier into it, and then dump it back in the file so it can be read by the batch file and pushed back
$SetSCPermsPS1 = @'
$sid = "@sid"
$perms = Get-Content -Path @permTXT
$pattern = '(S:)'
$elements = [regex]::Split($perms, $pattern)
$newperms = ( $elements[0] + "(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;" + $sid + ")S:" + $elements[2] ).Trim(" ")
Set-Content -Path @permTXT $newperms
'@

# strap in, we're replacing a lot of things here
$nssmBatch = $nssmBatch.Replace( "@nssm", $nssmDIR )
$nssmBatch = $nssmBatch.Replace( "@serviceName", $serviceName )
$nssmBatch = $nssmBatch.Replace( "@komgaServiceBAT", ( Join-Path -Path $runDIR -ChildPath "KomgaService.bat" ) )
$nssmBatch = $nssmBatch.Replace( "@serviceName", $serviceName )
$nssmBatch = $nssmBatch.Replace( "@user", $username )
$nssmBatch = $nssmBatch.Replace( "@pass", $password )
$nssmBatch = $nssmBatch.Replace( "@permTXT", ( Join-Path -Path $runDIR -ChildPath "perms.txt" ) )
$nssmBatch = $nssmBatch.Replace( "@permPS1", ( Join-Path -Path $runDIR -ChildPath "SetSCPerms.ps1" ) )

$SetSCPermsPS1 = $SetSCPermsPS1.Replace( "@user", $username )
$SetSCPermsPS1 = $SetSCPermsPS1.Replace( "@sid", $sid )
$SetSCPermsPS1 = $SetSCPermsPS1.Replace( "@permTXT", ( Join-Path -Path $runDIR -ChildPath "perms.txt" ) )

# Write the files
Set-Content -Path ( Join-Path -Path "$runDIR" -ChildPath "nssm.bat" ) $nssmBatch
Set-Content -Path ( Join-Path -Path "$runDIR" -ChildPath "SetSCPerms.ps1" ) $SetSCPermsPS1

# Path this batch
$nssmBatch = ( Join-Path -Path "$runDIR" -ChildPath "nssm.bat" )

# Start the batch file as admin
Write-Host Setting up NSSM...
Start-Process cmd.exe -ArgumentList "/C $nssmBatch" -Wait -WindowStyle Hidden -Verb RunAs
}


# Hey, we made it... Start the service, and we're good to go
Write-Host Starting the service... Give Komga some time to start up, and then access your browser at http://localhost:25600/
Start-Service -Name "$serviceName"

Write-Host This script will automatically launch the web UI in 30 seconds, or you can close it and launch it yourself.
Timeout /T 30
Start-Process "http://localhost:25600"
