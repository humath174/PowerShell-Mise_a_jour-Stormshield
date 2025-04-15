### Stormshield Firewall Updater

This PowerShell script, is provided as-is, with no warranty.
Created by Dregnoxx:
[Dregnoxx.tech](https://dregnoxx.tech) 
[@dregnoxx](https://twitter.com/dregnoxx).

# Requirements
- [WinSCP](https://winscp.net/eng/download.php)
- [Posh-SSH](https://github.com/darkoperator/Posh-SSH) (Install using `Install-Module -Name Posh-SSH`, in a Powershell window)
- Your root folder should look like this:
![image](https://github.com/Dregnoxx/StormshieldFWUpdate/assets/40840621/f685a5a4-35e9-42c1-a190-c9a8db428515)  
# IMPORTANT
Dont forget to create a subfolder where you will host the update file.
In the sript the default name is "FichierMAJ".
In the subfolder, I recommend you to always rename the update LATESTXXX.maj and overwrite each time you put another .maj file.
Example: LATESTSN210.maj for every SN210 firewall.

# Compatibility
This script is designed for SN160 ,200, 210, 220, 310, 320, 510, and 710 series. It has not been tested on other models.

# Installation
In your Task Manager create a new task, add the name, planify when you want to execute it, and in action specify:
-  Start a program:  `%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe`
-  In argument add:  `-file C:\YOURROOTFOLDER\SN210.ps1 -User USERNAME -PSWD PASSWORD -IP 1.2.3.4 -Client CLIENT_NAME`
