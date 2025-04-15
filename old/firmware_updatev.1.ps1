param(
    [Parameter(Mandatory=$True, Position=0, ValueFromPipeline=$false)]
    [System.String]$User,

    [Parameter(Mandatory=$True, Position=1, ValueFromPipeline=$false)]
    [System.String]$PSWD,
    
    [Parameter(Mandatory=$True, Position=2, ValueFromPipeline=$false)]
    [System.String]$IP,

    [Parameter(Mandatory=$True, Position=3, ValueFromPipeline=$false)]
    [System.String]$Client,

    [Parameter(Mandatory=$false)]
    [System.String]$Port = "13422"
)

# Charger l'assembly pour l'interface graphique
Add-Type -AssemblyName System.Windows.Forms

# Afficher la boîte de dialogue pour sélectionner le fichier
$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Title = "Sélectionnez le fichier de mise à jour (.maj)"
$fileDialog.Filter = "Fichiers de mise à jour (*.maj)|*.maj"
$fileDialog.Multiselect = $false

# Si l'utilisateur annule la sélection
if ($fileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Aucun fichier sélectionné. Le script va s'arrêter."
    exit 1
}

$UpdateFilePath = $fileDialog.FileName

# Convert password to secure string and create credentials
$PASSWORD = ConvertTo-SecureString -String $PSWD -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential ($User, $PASSWORD)
$WSCPLogin = "$User" + ":" + "$PSWD"

## LOG Firewall version + Date in the format day, month, year, hour, minute
$date = Get-Date -Format "dd-MM-yyyy-HH-mm"
$logPath = "C:\logs"
$logfile = "$logPath\$Client-$date.log"

# Create logs directory if it doesn't exist
if (-not (Test-Path -Path $logPath)) {
    New-Item -ItemType Directory -Path $logPath -Force
}

# Clear any existing trusted hosts
Get-SSHTrustedHost | Remove-SSHTrustedHost 

## SFTP Connection + upload the update file
Add-Content -Path $logfile -Value "------------------------------------------"
Add-Content -Path $logfile -Value "WINSCP upload started at $(Get-Date)"
Add-Content -Path $logfile -Value "------------------------------------------"

try {
    # WinSCP upload command
    & "C:\Program Files (x86)\WinSCP\WinSCP.com" /command `
        "open scp://$WSCPLogin@$IP`:$Port -hostkey=`"*`"" `
        "put `"$UpdateFilePath`" `"/usr/Firewall/Update/`"" `
        "exit"
    
    Add-Content -Path $logfile -Value "File uploaded successfully"
} catch {
    Add-Content -Path $logfile -Value "Error during WinSCP upload: $_"
    exit 1
}

## SSH connection + update
try {
    $sessionParams = @{
        ComputerName = $IP
        Credential   = $Credential
        AcceptKey    = $true
        Port         = 13422  # Default SSH port
    }

    $sessionssh = New-SSHSession @sessionParams -ErrorAction Stop
    
    Add-Content -Path $logfile -Value "------------------------------------------"
    Add-Content -Path $logfile -Value "Firmware version before the update"
    Add-Content -Path $logfile -Value "------------------------------------------"
    
    $preVersion = Invoke-SSHCommand -SSHSession $sessionssh -Command "getversion" -ErrorAction Stop
    Add-Content -Path $logfile -Value $preVersion.Output
    
    Add-Content -Path $logfile -Value "------------------------------------------"
    Add-Content -Path $logfile -Value "Starting firewall update at $(Get-Date)"
    Add-Content -Path $logfile -Value "------------------------------------------"
    
    $updateResult = Invoke-SSHCommand -SSHSession $sessionssh -Command "fwupdate -r -f /usr/Firewall/Update/$(Split-Path $UpdateFilePath -Leaf)" -ErrorAction Stop
    Add-Content -Path $logfile -Value $updateResult.Output
    
    Remove-SSHSession -SSHSession $sessionssh | Out-Null
    
    Add-Content -Path $logfile -Value "Update command sent successfully. Waiting for 10 minutes..."
    
    ## Sleep for 10 mins needed to let the update do its job
    Start-Sleep -Seconds 600
    
    ## Verify update
    $sessionssh = New-SSHSession @sessionParams -ErrorAction Stop
    
    Add-Content -Path $logfile -Value "------------------------------------------"
    Add-Content -Path $logfile -Value "Firmware version after the update"
    Add-Content -Path $logfile -Value "------------------------------------------"
    
    $postVersion = Invoke-SSHCommand -SSHSession $sessionssh -Command "getversion" -ErrorAction Stop
    Add-Content -Path $logfile -Value $postVersion.Output
    
    Remove-SSHSession -SSHSession $sessionssh | Out-Null
    
} catch {
    Add-Content -Path $logfile -Value "Error during SSH operations: $_"
    exit 1
}

Add-Content -Path $logfile -Value "------------------------------------------"
Add-Content -Path $logfile -Value "Script completed successfully at $(Get-Date)"
Add-Content -Path $logfile -Value "------------------------------------------"