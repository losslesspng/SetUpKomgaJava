# SetUpKomgaJava
A Powershell script to get up and running with Komga

## So, what the heck is this?
So glad you asked! This is a powershell script t-OH WAIT it says that right there.

This script is all you need to do the following:
- Download the latest Komga jar
- Download java
- Download NSSM
- Configure NSSM to run Komga as a Windows service
- Create a script that will allow you to easily check for and apply updates to Komga

### This script makes use of the following:
- NSSM, from https://nssm.cc/
- Adoptium Temurin 17 LTS, from https://adoptium.net/
- Komga, from https://github.com/gotson/komga

Please read the instructions in the big 'ol comment block at the top of the script, but I'll quickly spell it out here for you:
You can run this script as provided, no edits necessary, and it should work. It has a "Works on my machine" badge, but you can find me here or on the Komga discord server if you run into trouble.

### *** IMPORTANT *** If you're running a Microsoft Account, please check out the [Wiki article](https://github.com/losslesspng/SetUpKomgaJava/wiki/Running-With-a-Microsoft-Account)! 

The script should do a good job at keeping itself organized, but you can easily change the directory it gets set up in, the name of the rolling backup folder for komga versions (it keeps the latest 3 versions that you've run, in case something goes wrong and you need to revert), the name of the service that runs komga, and the arguments that get passed to java as it runs. The only default arguement runs komga at 4gb of ram. 

### You do not need to run this script as administrator, but you **do** need to have an admin account to setup NSSM
I recommend running the script like this:
```powershell
Powershell.exe -ExecutionPolicy Bypass -NoProfile -File .\SetUpKomgaJava.ps1
```
I tested it against a brand new Windows 10 install, and this runs without having to change any powershell settings. 
This script will prompt you once for your password, and then again to allow it to launch part of the configuration as administrator. This is necessary to install the service with NSSM, as well as make it run with your account, and allow you to stop and start the service without admin access in the future, which makes the update script run smoother.
