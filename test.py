# exportconfig.ps1 - Version SRP Workaround
param(
    [string]$IP = "172.16.70.57",
    [int]$Port = 13422
)

# 1. Configuration
$Username = "admin"
$Password = Read-Host "Mot de passe SRP" -AsSecureString
$PlainPass = [System.Net.NetworkCredential]::new("", $Password).Password
$BackupPass = Read-Host "Mot de passe backup" -AsSecureString
$PlainBck = [System.Net.NetworkCredential]::new("", $BackupPass).Password

# 2. Fonction spéciale pour SRP
function Invoke-StormshieldSSH {
    $plink = Start-Process "plink" -ArgumentList "-ssh $IP -P $Port -l $Username -pw $PlainPass" -NoNewWindow -PassThru -RedirectStandardInput ".\input.txt"
    
    # Envoi séquentiel avec timing
    @"
cli
$PlainPass
modify on force
CONFIG BACKUP list=all [password=$PlainBck]> backup_$(Get-Date -Format yyyyMMdd).na
exit
"@ | Out-File ".\input.txt" -Encoding ASCII

    Start-Sleep -Seconds 15  # Temps pour le backup
    if (!$plink.HasExited) { $plink.Kill() }
}

# 3. Exécution
try {
    Write-Host "Lancement du processus SRP..."
    Invoke-StormshieldSSH
    Write-Host "Backup devrait être terminé. Vérifiez sur le firewall." -ForegroundColor Green
}
catch {
    Write-Host "ERREUR: $_" -ForegroundColor Red
}
finally {
    Remove-Item ".\input.txt" -ErrorAction SilentlyContinue
}