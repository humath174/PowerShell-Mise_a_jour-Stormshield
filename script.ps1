# Configuration de base
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Paramètres de connexion
$hostname = "172.16.70.57" # ou Read-Host pour la saisie interactive
$username = "admin"        # ou Read-Host pour la saisie interactive
$port = 13422
$localFilePath = "C:\Users\mathe\Downloads\fwupd-4.3.36-SNS-amd64-M.maj"
$remoteFilePath = "/var/tmp/" + [System.IO.Path]::GetFileName($localFilePath)

# Chemins des outils
$pscpPath = "C:\Program Files\PuTTY\pscp.exe"
$plinkPath = "C:\Program Files\PuTTY\plink.exe"

# Fonction pour exécuter des commandes SSH
function Invoke-SSHCommand {
    param(
        [string]$Command,
        [switch]$IgnoreErrors = $false
    )
    
    try {
        $output = & $plinkPath -ssh -P $port -batch "$username@$hostname" $Command 2>&1
        if (-not $IgnoreErrors -and ($LASTEXITCODE -ne 0 -or $output -match "error|fail")) {
            throw $output
        }
        return $output
    } catch {
        if (-not $IgnoreErrors) {
            Write-Host "❌ Erreur SSH: $_" -ForegroundColor Red
            exit 1
        }
        return $null
    }
}

# 1. Vérification de l'espace disque
Write-Host "🔍 Vérification de l'espace disque..."
$diskSpace = Invoke-SSHCommand "df -h /var/tmp | tail -1 | awk '{print \$4}'"
Write-Host "Espace disponible: $diskSpace"

if ([int]$diskSpace.TrimEnd('M') -lt 100) {
    Write-Host "⚠️ Nettoyage de l'espace disque..."
    Invoke-SSHCommand "rm -f /var/tmp/*.maj" -IgnoreErrors
}

# 2. Transfert du fichier
Write-Host "📤 Transfert du fichier..."
try {
    & $pscpPath -P $port -batch "$localFilePath" "$username@$hostname`:$remoteFilePath"
    Write-Host "✅ Transfert réussi" -ForegroundColor Green
} catch {
    Write-Host "❌ Échec du transfert: $_" -ForegroundColor Red
    exit 1
}

# 3. Vérification du fichier
Write-Host "🔎 Vérification du fichier..."
$fileCheck = Invoke-SSHCommand "ls -lh $remoteFilePath"
if (-not $fileCheck) {
    Write-Host "❌ Fichier non trouvé" -ForegroundColor Red
    exit 1
}
Write-Host "✅ Fichier présent: $fileCheck" -ForegroundColor Green

# 4. Application de la mise à jour
Write-Host "⚙️ Démarrage de la mise à jour..."
try {
    $updateOutput = Invoke-SSHCommand "update $remoteFilePath --no-verify"
    Write-Host "📋 Résultat de la mise à jour:"
    $updateOutput
    
    if ($updateOutput -match "successfully installed|installation réussie") {
        Write-Host "✅ Mise à jour installée" -ForegroundColor Green
    } else {
        Write-Host "❌ Échec de l'installation" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "❌ Erreur critique: $_" -ForegroundColor Red
    exit 1
}

# 5. Redémarrage optionnel
$choice = Read-Host "Redémarrer le Stormshield maintenant? (O/N)"
if ($choice -eq "O" -or $choice -eq "o") {
    Write-Host "🔄 Redémarrage en cours..."
    Invoke-SSHCommand "reboot" -IgnoreErrors
    Write-Host "✅ Redémarrage initié" -ForegroundColor Green
}

Write-Host "🏁 Opération terminée avec succès" -ForegroundColor Green