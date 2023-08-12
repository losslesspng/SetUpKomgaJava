#############################################################################################################
# This script will automatically download:                                                                  #
#     • Scoop                                                                                               #
#       • Shawl                                                                                             #
#       • Java (Adoptium Temurin 17 LTS)                                                                    #
#       • Komga                                                                                             #
# It will then:                                                                                             #
#     • Configure a batch script that uses Shawl to run java with komga jar as a service                    #
#       • Add a description to the service                                                                  #
#       • Update the service to run automatically, and run with the logged in user account                  #
#       • Adjust permissions to allow user account to start and stop service without admin                  #
#     • Create a firewall rule to allow incoming TCP traffic on port 25600                                  #
#     • Start the service                                                                                   #
# Please read over the code below and make the changes necessary, OR feel free to leave them as is          #
# The settings provided below are Works On My Machine certified and any edits are purely for your own sake. #
# You may use the following line to run this file:                                                          #
# Powershell.exe -ExecutionPolicy Bypass -NoProfile -File .\SetUpKomgaJava.ps1                              #
# This will run the script without requiring you to modify the ExecutionPolicy for your whole computer      #
#                                                                                                           #
#                                                                                                           #
# This script is lovingly provided by png, with help from Diesel and gotson.                                #
#############################################################################################################
# Turn off errors from powershell
$ErrorActionPreference = 'SilentlyContinue'

Write-Host "Thanks for checking out Komga, hope you love it!"

#########
# Scoop #
#########
# We are going to clear the powershell errors and check for git being installed.
Write-Host "Installing Scoop"
Invoke-RestMethod get.scoop.sh | Invoke-Expression
$error.Clear()
Write-Host "Checking for git in path (required for scoop)"
git -v
if ($error)
{ 
    Write-Host "git not found, scoop will install and maintain its own version. This will also install 7zip if it is not in your path already."
    scoop install git 
}
# Add the buckets and the applications we need
Write-Host "Adding scoop buckets for java, komga, and shawl"
scoop bucket add java
scoop bucket add extras
Write-Host "Installing packages"
scoop install shawl komga temurin17-jre

###########
# Scripts #
###########
# This script will start java to run the komga jar file, with 4GB of RAM configured as the maximum memory size. 
$KomgaServiceBAT = @'
@echo off
REM You can modify the below line to adjust komga's memory use. Just stop the service and start it again when changing it. 
%HOMEPATH%\scoop\apps\temurin17-jre\current\bin\java -jar -Xmx4g %HOMEPATH%\scoop\apps\komga\current\komga.jar
'@
# This little guy is going to use shawl to create the service, modify it to run as the logged in user, set a description, 
# and then it will set it so that the user can start and stop the service without admin
$KomgaServicePS1 = @'
shawl add --name KomgaService -- C:\Users\$env:USERNAME\.komga\service.bat
$cred = Get-Credential -Credential $env:USERDOMAIN\$env:USERNAME
sc.exe config KomgaService obj= $cred.Username password= $cred.GetNetworkCredential().Password start= auto type= own
Set-Service -Name KomgaService -Description 'A Service for Komga.'
$sid = ( New-Object System.Security.Principal.NTAccount( $env:USERNAME ) ).Translate( [System.Security.Principal.SecurityIdentifier] ).Value
$perms = sc.exe sdshow KomgaService
$pattern = "(S:)"
$elements = [regex]::Split($perms, $pattern)
$newperms = ( $elements[0] + '(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;' + $sid + ')S:' + $elements[2] ).Trim(' ')
sc.exe sdset KomgaService $newperms
New-NetFirewallRule -DisplayName "Allow Komga" -Direction Inbound -Action Allow -LocalPort 25600 -Protocol TCP
Timeout /T 10
'@
# Create a base application.yml file for komga.
$ApplicationYML = @'
# Lines starting with # are comments
# Make sure indentation is correct (2 spaces at every indentation level), yaml is very sensitive!
komga:
  libraries-scan-cron: "0 0 */8 * * ?"  # periodic scan every 8 hours
  # libraries-scan-cron: "-"              # disable periodic scan
  # libraries-scan-startup: false         # scan libraries at startup
  database:
    file: ${user.home}/.komga/database.sqlite
  delete-empty-collections: true
  delete-empty-read-lists: true
server:
  # If you change this port, please update the firewall rule "Allow Komga" to match
  port: 25600
'@
# Pre-create the .komga folder if it doesn't exist
If ( !( Test-Path -Path "~\.komga\" ) ) { New-Item -Path "~\.komga\" -ItemType Directory 2>&1 | Out-Null }

# Save the above files
Set-Content -Path "~\.komga\service.bat" -Value $KomgaServiceBAT
Set-Content -Path "~\.komga\setup-service.ps1" -Value $KomgaServicePS1
Set-Content -Path "~\.komga\application.yml" -Value $ApplicationYML

# Run the script to make the service, this requires an admin prompt. 
Start-Process Powershell.exe -Verb RunAs "~\.komga\setup-service.ps1" -Wait
Remove-Item -Path "~\.komga\setup-service.ps1"

# Hey, we made it... Start the service, and we're good to go
Write-Host Starting the service... Give Komga some time to start up, and then access your browser at http://localhost:25600/
Start-Service -Name "KomgaService"
Write-Host This script will automatically launch the web UI in 30 seconds, or you can close it and launch it yourself.
Timeout /T 30
Start-Process "http://localhost:25600"
