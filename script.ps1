# Demander les informations à l'utilisateur
$hostname = Read-Host "Entrez l'adresse IP ou le nom de l'hôte"
$username = Read-Host "Entrez votre nom d'utilisateur"
$port = 13422
$localFilePath = Read-Host "Entrez le chemin local vers le fichier de mise à jour"
$remoteFilePath = "/var/tmp/" + [System.IO.Path]::GetFileName($localFilePath)
 
# Chemin vers pscp.exe et plink.exe
$pscpPath = "C:\Program Files\PuTTY\pscp.exe"
$plinkPath = "C:\Program Files\PuTTY\plink.exe"
 
# Commande pour envoyer le fichier via SCP
Write-Host "Transfert du fichier de mise à jour..."
$args = "-P $port $localFilePath $username@${hostname}:${remoteFilePath}"
Start-Process -FilePath "$pscpPath" -ArgumentList $args -NoNewWindow -Wait
 
# Pause pour vérifier le transfert du fichier
Read-Host "Vérifiez si le fichier a été transféré correctement, puis appuyez sur Entrée pour continuer."
 
# Commande pour vérifier la présence du fichier et ses permissions
Write-Host "Vérification de la présence du fichier sur l'équipement..."
$checkFileCommand = "ls -l ${remoteFilePath} && if [ -f ${remoteFilePath} ]; then echo 'Fichier transféré avec succès'; else echo 'Échec du transfert du fichier'; fi"
$checkFileArgs = "-ssh $username@$hostname -P $port $checkFileCommand"
Start-Process -FilePath "$plinkPath" -ArgumentList $checkFileArgs -NoNewWindow -Wait
 
# Pause pour vérifier la présence du fichier
Read-Host "Vérifiez si le fichier est présent sur l'équipement, puis appuyez sur Entrée pour continuer."
 
# Commande pour appliquer la mise à jour avec diagnostic
Write-Host "Application de la mise à jour..."
$applyUpdateCommand = "chmod +x ${remoteFilePath} && ${remoteFilePath} > /var/tmp/update_log.txt 2>&1"
$applyUpdateArgs = "-ssh $username@$hostname -P $port $applyUpdateCommand"
Start-Process -FilePath "$plinkPath" -ArgumentList $applyUpdateArgs -NoNewWindow -Wait
 
# Pause pour vérifier l'application de la mise à jour
Read-Host "Vérifiez si la mise à jour a été appliquée correctement, puis appuyez sur Entrée pour continuer."
 
# Commande pour afficher le journal de mise à jour
Write-Host "Affichage du journal de mise à jour..."
$viewLogCommand = "cat /var/tmp/update_log.txt"
$viewLogArgs = "-ssh $username@$hostname -P $port $viewLogCommand"
Start-Process -FilePath "$plinkPath" -ArgumentList $viewLogArgs -NoNewWindow -Wait
 
# Commande pour redémarrer le firewall via SSH
Write-Host "Redémarrage du firewall..."
$restartArgs = "-ssh $username@$hostname -P $port reboot"
Start-Process -FilePath "$plinkPath" -ArgumentList $restartArgs -NoNewWindow -Wait
 
Write-Host "Processus terminé."