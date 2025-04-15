####
#   Powershell script made by Dregnoxx
#   Provided as is, no warranty and no support will be added
#   Dregnoxx.tech | @dregnoxx
####
#   This script require the use of WinSCP and PoshSSH
#   https://winscp.net/eng/download.php
#   https://github.com/darkoperator/Posh-SSH    |   Install-Module -Name Posh-SSH
###
#   Made for SN160,210,220,310,320,510,710 series
#   Not yet tested on other models
###

##  VAR for each setting you will need to specify in your task
# Each block is it own VAR
param(

    [Parameter(Mandatory=$True, Position=0, ValueFromPipeline=$false)]
    [System.String]
    $User,

    [Parameter(Mandatory=$True, Position=1, ValueFromPipeline=$false)]
    [System.String]
    $PSWD,
    
    [Parameter(Mandatory=$True, Position=3, ValueFromPipeline=$false)]
    [System.String]
    $IP,

    [Parameter(Mandatory=$True, Position=4, ValueFromPipeline=$false)]
    [System.String]
    $Client
)

$PASSWORD = ConvertTo-SecureString -String $PSWD -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential ($User, $PASSWORD)
$IPSSH = $IP + ":22"
$WSCPLogin = "$User"+":"+"$PSWD"

## LOG Firewall version + Date in the format day, month, year, hour, minute
$date = Get-Date -Format "dd-MM-yyyy-HH-mm"
$logfile = "C:\YOURROOTFOLDER\LOG\" + $Client+ "-" + $date + ".LOG"


Get-SSHTrustedHost | Remove-SSHTrustedHost 

##  SFTP Connexion + upload the update file
ADD-content -path $logfile -value "------------------------------------------"
ADD-content -path $logfile -value "WINSCP upload"
ADD-content -path $logfile -value "------------------------------------------"
C:\YOURROOTFOLDER\WinSCP.com /command "open $WSCPLogin@$IPSSH -certificate=*" "put C:\YOURROOTFOLDER\FichierMAJ\LATESTSN210.maj /usr/Firewall/" "exit" | Out-File $logfile -Append -encoding UTF8
ADD-content -path $logfile -value ""

### SSH connection + update
$sessionssh=New-SSHSession -ComputerName "$IP" -AcceptKey: $true -Credential $Credential
ADD-content -path $logfile -value "------------------------------------------"
ADD-content -path $logfile -value "Firwmare version before the update"
ADD-content -path $logfile -value "------------------------------------------"
Invoke-SSHCommand -SSHSession $sessionssh -Command "getversion" | Out-File $logfile -Append -encoding UTF8
ADD-content -path $logfile -value ""
ADD-content -path $logfile -value "------------------------------------------"
ADD-content -path $logfile -value "Firewall update"
ADD-content -path $logfile -value "------------------------------------------"  
Invoke-SSHCommand -SSHSession $sessionssh -Command "fwupdate -r -f /usr/Firewall/LATESTSN210.maj" | Out-File $logfile -Append -encoding UTF8
Remove-SSHSession -SSHSession $sessionssh

### Sleep for 10 mins needed to let the update do it job
Start-Sleep -Seconds 600


$sessionssh=New-SSHSession -ComputerName "$IP" -AcceptKey: $true -Credential $Credential
ADD-content -path $logfile -value "------------------------------------------"
ADD-content -path $logfile -value "Firmware version after the update"
ADD-content -path $logfile -value "------------------------------------------"
Invoke-SSHCommand -SSHSession $sessionssh -Command "getversion" | Out-File $logfile -Append -encoding UTF8
Remove-SSHSession -SSHSession $sessionssh
