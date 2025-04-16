# === Configuration ===
$firewallIP = "172.16.70.57"  # Adresse IP du Stormshield
$sshPort = 13422                 # Port SSH (par défaut : 22)
$username = "admin"           # Nom d'utilisateur
$password = "admin"           # Mot de passe
$localBackupDir = "C:\Sauvegardes"  # Répertoire local pour stocker la sauvegarde
$remoteBackupFile = "mybackup.na"  # Nom du fichier de sauvegarde sur le Stormshield

# === Préparation du répertoire local ===
if (-not (Test-Path -Path $localBackupDir)) {
    New-Item -ItemType Directory -Path $localBackupDir -Force | Out-Null
}

$localBackupFile = Join-Path $localBackupDir $remoteBackupFile

# === Commande de sauvegarde ===
$backupCommand = "cli; $password; modify on force; CONFIG BACKUP list=all [password=$password] > $remoteBackupFile"

# === Connexion SSH et exécution ===
try {
    # Charger le module SSH.NET
    Import-Module -Name SSH-Sessions -ErrorAction Stop

    Write-Host "Connexion au pare-feu Stormshield ($sshPort)..."

    # Créer une session SSH
    $sessionParams = @{
        ComputerName = $firewallIP
        Credential   = (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, (ConvertTo-SecureString -String $password -AsPlainText -Force))
        Port         = $sshPort
        AcceptKey    = $true
    }

    $sshSession = New-SSHSession @sessionParams -ErrorAction Stop

    Write-Host "Exécution de la commande de sauvegarde..."
    Invoke-SSHCommand -SSHSession $sshSession -Command $backupCommand -ErrorAction Stop

    Write-Host "Téléchargement du fichier de sauvegarde ($remoteBackupFile)..."
    Get-SFTPFile -SSHSession $sshSession -RemoteFile "/$remoteBackupFile" -LocalPath $localBackupFile -ErrorAction Stop

    Write-Host "✅ Sauvegarde réussie ! Fichier enregistré : $localBackupFile" -ForegroundColor Green

    # Supprimer le fichier distant après téléchargement
    Invoke-SSHCommand -SSHSession $sshSession -Command "rm /$remoteBackupFile" -ErrorAction SilentlyContinue

    # Fermer la session SSH
    Remove-SSHSession -SSHSession $sshSession | Out-Null
} catch {
    Write-Host "❌ Une erreur est survenue : $_" -ForegroundColor Red
}