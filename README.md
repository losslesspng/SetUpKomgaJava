# SetUpKomgaJava
A Powershell script to get up and running with Komga

## So, what the heck is this?
So glad you asked! This is a powershell script t-OH WAIT it says that right there.

This script is all you need to do the following:
- Download and install Scoop
- Add Scoop buckets Java and Extras
- Install Shawl, Java, and Komga apps from Scoop
- Configure Shawl to run Komga as a Windows service
- Create a script that will allow you to easily check for and apply updates to Komga
- Create a firewall rule to allow incoming traffic on port 25600 (new komga default)

### This script makes use of the following:
- [Scoop](https://github.com/ScoopInstaller/Scoop)
- [Adoptium Temurin 17 LTS](https://adoptium.net/)
- [Komga](https://github.com/gotson/komga)
- [Shawl](https://github.com/mtkennerly/shawl)

Please read the instructions in the big 'ol comment block at the top of the script, but I'll quickly spell it out here for you:
You can run this script as provided, no edits necessary, and it should work. It has a "Works on my machine" badge, but you can find me here or (more likely) on the Komga discord server (@png) if you run into trouble.

### *** IMPORTANT *** If you're running a Microsoft Account, please check out the [Wiki article](https://github.com/losslesspng/SetUpKomgaJava/wiki/Running-With-a-Microsoft-Account)! 

### You do not need to run this script as administrator, but you **do** need to have an admin account to setup the background service
I recommend running the script like this:
```powershell
Powershell.exe -ExecutionPolicy Bypass -NoProfile -File .\SetUpKomgaJava.ps1
```
I tested it against a brand new Windows 10 install, and this runs without having to change any powershell settings. 
This script will prompt you once for your password, which is uses to set your account as the one running komga. This is to make it easier to set up your library, komga will be able to access anything that your user account can. 
