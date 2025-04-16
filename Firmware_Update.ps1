# Charger les assemblies pour l'interface graphique
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Créer le formulaire principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Sauvegarde de configuration Stormshield"
$form.Size = New-Object System.Drawing.Size(500,400)
$form.StartPosition = "CenterScreen"

# Position verticale courante pour les contrôles
$yPos = 10

# Champ pour l'adresse IP
$labelIP = New-Object System.Windows.Forms.Label
$labelIP.Location = New-Object System.Drawing.Point(10,$yPos)
$labelIP.Size = New-Object System.Drawing.Size(200,20)
$labelIP.Text = "Adresse IP du firewall:"
$form.Controls.Add($labelIP)

$textIP = New-Object System.Windows.Forms.TextBox
$textIP.Location = New-Object System.Drawing.Point(220,$yPos)
$textIP.Size = New-Object System.Drawing.Size(250,20)
$form.Controls.Add($textIP)
$yPos += 30

# Champ pour l'utilisateur admin
$labelUser = New-Object System.Windows.Forms.Label
$labelUser.Location = New-Object System.Drawing.Point(10,$yPos)
$labelUser.Size = New-Object System.Drawing.Size(200,20)
$labelUser.Text = "Nom d'utilisateur admin:"
$form.Controls.Add($labelUser)

$textUser = New-Object System.Windows.Forms.TextBox
$textUser.Location = New-Object System.Drawing.Point(220,$yPos)
$textUser.Size = New-Object System.Drawing.Size(250,20)
$form.Controls.Add($textUser)
$yPos += 30

# Champ pour le mot de passe
$labelPSWD = New-Object System.Windows.Forms.Label
$labelPSWD.Location = New-Object System.Drawing.Point(10,$yPos)
$labelPSWD.Size = New-Object System.Drawing.Size(200,20)
$labelPSWD.Text = "Mot de passe:"
$form.Controls.Add($labelPSWD)

$textPSWD = New-Object System.Windows.Forms.TextBox
$textPSWD.Location = New-Object System.Drawing.Point(220,$yPos)
$textPSWD.Size = New-Object System.Drawing.Size(250,20)
$textPSWD.PasswordChar = '*'
$form.Controls.Add($textPSWD)
$yPos += 30

# Champ pour le port
$labelPort = New-Object System.Windows.Forms.Label
$labelPort.Location = New-Object System.Drawing.Point(10,$yPos)
$labelPort.Size = New-Object System.Drawing.Size(200,20)
$labelPort.Text = "Port (défaut: 13422):"
$form.Controls.Add($labelPort)

$textPort = New-Object System.Windows.Forms.TextBox
$textPort.Location = New-Object System.Drawing.Point(220,$yPos)
$textPort.Size = New-Object System.Drawing.Size(250,20)
$textPort.Text = "13422"
$form.Controls.Add($textPort)
$yPos += 40

# Case à cocher pour sauvegarder la partition
$checkPartition = New-Object System.Windows.Forms.CheckBox
$checkPartition.Location = New-Object System.Drawing.Point(10,$yPos)
$checkPartition.Size = New-Object System.Drawing.Size(300,20)
$checkPartition.Text = "Sauvegarder la partition (nécessite plus d'espace)"
$form.Controls.Add($checkPartition)
$yPos += 30

# Bouton de sauvegarde
$buttonBackup = New-Object System.Windows.Forms.Button
$buttonBackup.Location = New-Object System.Drawing.Point(150,$yPos)
$buttonBackup.Size = New-Object System.Drawing.Size(200,30)
$buttonBackup.Text = "Lancer la sauvegarde"
$form.Controls.Add($buttonBackup)
$yPos += 40

# Barre de progression
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,$yPos)
$progressBar.Size = New-Object System.Drawing.Size(460,20)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$form.Controls.Add($progressBar)
$yPos += 30

# Label de statut
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(10,$yPos)
$labelStatus.Size = New-Object System.Drawing.Size(460,20)
$labelStatus.Text = "Prêt"
$form.Controls.Add($labelStatus)

# Gestionnaire d'événements pour le bouton de sauvegarde
$buttonBackup.Add_Click({
    # Récupérer les valeurs des champs
    $IP = $textIP.Text
    $User = $textUser.Text
    $PSWD = $textPSWD.Text
    $Port = $textPort.Text
    $backupPartition = $checkPartition.Checked

    # Validation des champs obligatoires
    if ([string]::IsNullOrEmpty($IP) -or [string]::IsNullOrEmpty($User) -or [string]::IsNullOrEmpty($PSWD)) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez remplir tous les champs obligatoires!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Créer le dossier de sauvegarde s'il n'existe pas
    $backupDir = "C:\BackupStormshield"
    if (-not (Test-Path -Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    }

    # Générer un nom de fichier avec date/heure
    $date = Get-Date -Format "yyyyMMdd-HHmmss"
    $backupFile = "$backupDir\StormshieldBackup_$date.conf"
    $partitionFile = "$backupDir\StormshieldPartition_$date.tar.gz"

    try {
        # Convert password to secure string and create credentials
        $PASSWORD = ConvertTo-SecureString -String $PSWD -AsPlainText -Force
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential ($User, $PASSWORD)
        $WSCPLogin = "$User" + ":" + "$PSWD"

        # Configuration des paramètres SSH
        $sessionParams = @{
            ComputerName = $IP
            Credential   = $Credential
            AcceptKey    = $true
            Port         = $Port
        }

        $labelStatus.Text = "Connexion au firewall..."
        $progressBar.Value = 10

        # Étape 1: Exporter la configuration
        $labelStatus.Text = "Export de la configuration..."
        $progressBar.Value = 30

        $sessionssh = New-SSHSession @sessionParams -ErrorAction Stop
        $exportResult = Invoke-SSHCommand -SSHSession $sessionssh -Command "export configuration" -ErrorAction Stop
        
        # Vérifier si l'export a réussi
        if ($exportResult.Output -notmatch "Export configuration: OK") {
            throw "Échec de l'export de la configuration"
        }

        # Étape 2: Télécharger le fichier de configuration
        $labelStatus.Text = "Téléchargement de la configuration..."
        $progressBar.Value = 50

        & "C:\Program Files (x86)\WinSCP\WinSCP.com" /command `
            "open scp://$WSCPLogin@$IP`:$Port -hostkey=`"*`"" `
            "get `"/usr/Firewall/Update/export.conf`" `"$backupFile`"" `
            "exit"

        if (-not (Test-Path -Path $backupFile)) {
            throw "Échec du téléchargement du fichier de configuration"
        }

        # Étape 3: Sauvegarder la partition si demandé
        if ($backupPartition) {
            $labelStatus.Text = "Sauvegarde de la partition..."
            $progressBar.Value = 70

            $partitionResult = Invoke-SSHCommand -SSHSession $sessionssh -Command "backup partition /tmp/partition.tar.gz" -ErrorAction Stop
            
            if ($partitionResult.Output -notmatch "Backup partition: OK") {
                throw "Échec de la sauvegarde de la partition"
            }

            & "C:\Program Files (x86)\WinSCP\WinSCP.com" /command `
                "open scp://$WSCPLogin@$IP`:$Port -hostkey=`"*`"" `
                "get `"/tmp/partition.tar.gz`" `"$partitionFile`"" `
                "rm `"/tmp/partition.tar.gz`"" `
                "exit"

            if (-not (Test-Path -Path $partitionFile)) {
                throw "Échec du téléchargement du fichier de partition"
            }
        }

        # Nettoyage
        Invoke-SSHCommand -SSHSession $sessionssh -Command "rm /usr/Firewall/Update/export.conf" -ErrorAction SilentlyContinue
        Remove-SSHSession -SSHSession $sessionssh | Out-Null

        $labelStatus.Text = "Sauvegarde terminée avec succès!"
        $progressBar.Value = 100

        # Afficher un message de succès
        $message = "Sauvegarde terminée avec succès!`n`n"
        $message += "Fichier de configuration: $backupFile`n"
        if ($backupPartition) {
            $message += "Fichier de partition: $partitionFile`n"
        }

        [System.Windows.Forms.MessageBox]::Show($message, "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    } catch {
        $labelStatus.Text = "Erreur lors de la sauvegarde"
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de la sauvegarde: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $progressBar.Value = 0
    }
})

# Afficher le formulaire
[void]$form.ShowDialog()