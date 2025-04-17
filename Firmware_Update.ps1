# Charger les assemblies pour l'interface graphique
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Définir les modèles Stormshield et leurs catégories
$modeles = @{
    "VM" = @("EVA1", "EVA2", "EVA3", "EVA4", "EVAU", "VPAYG")
    "Taille S" = @("SN160-A", "SN160W-A", "SN210-A", "SN210W-A", "SN310-A")
    "Taille M" = @("SN510-A", "SN710-A", "SNi40-A", "SNi20-A", "SN-S-Series-220", "SN-S-Series-320", "SN-XS-Series-170", "SNi10")
    "Taille L" = @("SN6100-A", "SN3100-A", "SN2100-A", "SN910-A", "SN1100-A", "SN6000-A", "SN3000-A", "SN2000-A", "SNxr1200-A", "SN-M-Series-720", "SN-M-Series-920", "SN520-A", "SN-L-Series-2200", "SN-L-Series-3200", "SN-XL-Series-5200", "SN-XL-Series-6200")
}

# Fonction pour déterminer la catégorie d'un modèle
function Get-StormshieldCategory {
    param (
        [string]$model
    )
    
    foreach ($category in $modeles.Keys) {
        if ($modeles[$category] -contains $model) {
            return $category
        }
    }
    
    # Vérification partielle pour les modèles qui pourraient avoir des suffixes différents
    foreach ($category in $modeles.Keys) {
        foreach ($pattern in $modeles[$category]) {
            $basePattern = $pattern -replace '-.*$', ''
            if ($model -match "^$basePattern") {
                return $category
            }
        }
    }
    
    return $null
}

# Fonction pour charger les clients depuis un fichier CSV
function Load-ClientsFromCSV {
    param (
        [string]$filePath
    )
    
    try {
        $clients = Import-Csv -Path $filePath -Delimiter ";"
        return $clients
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors du chargement du fichier CSV: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
}

# Fonction pour sauvegarder un client dans le fichier CSV
function Save-ClientToCSV {
    param (
        [string]$filePath,
        [PSCustomObject]$client
    )
    
    try {
        # Vérifier si le fichier existe déjà
        if (Test-Path -Path $filePath) {
            $existingClients = Import-Csv -Path $filePath -Delimiter ";"
            # Vérifier si le client existe déjà
            $existingClient = $existingClients | Where-Object { $_.IP -eq $client.IP }
            if ($existingClient) {
                # Mettre à jour le client existant
                $existingClient.Nom = $client.Nom
                $existingClient.Utilisateur = $client.Utilisateur
                $existingClient.MotDePasse = $client.MotDePasse
                $existingClient.Port = $client.Port
            } else {
                # Ajouter le nouveau client
                $existingClients += $client
            }
            $existingClients | Export-Csv -Path $filePath -Delimiter ";" -NoTypeInformation -Force
        } else {
            # Créer un nouveau fichier avec le client
            $client | Export-Csv -Path $filePath -Delimiter ";" -NoTypeInformation -Force
        }
        return $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de la sauvegarde du client: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# Créer le formulaire principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Mise à jour de firmware Stormshield"
$form.Size = New-Object System.Drawing.Size(800,700) # Augmenté la taille pour la liste des clients
$form.StartPosition = "CenterScreen"

# Position verticale courante pour les contrôles
$yPos = 10

# Liste des clients
$labelClients = New-Object System.Windows.Forms.Label
$labelClients.Location = New-Object System.Drawing.Point(10,$yPos)
$labelClients.Size = New-Object System.Drawing.Size(200,20)
$labelClients.Text = "Liste des clients:"
$form.Controls.Add($labelClients)

$listClients = New-Object System.Windows.Forms.ListBox
$listClients.Location = New-Object System.Drawing.Point(10,($yPos + 20))
$listClients.Size = New-Object System.Drawing.Size(300,150)
$listClients.SelectionMode = "One"
$form.Controls.Add($listClients)
$yPos += 180

# Boutons pour la gestion des clients
$buttonLoadCSV = New-Object System.Windows.Forms.Button
$buttonLoadCSV.Location = New-Object System.Drawing.Point(320,$yPos)
$buttonLoadCSV.Size = New-Object System.Drawing.Size(150,30)
$buttonLoadCSV.Text = "Charger depuis CSV"
$form.Controls.Add($buttonLoadCSV)

$buttonSaveCSV = New-Object System.Windows.Forms.Button
$buttonSaveCSV.Location = New-Object System.Drawing.Point(480,$yPos)
$buttonSaveCSV.Size = New-Object System.Drawing.Size(150,30)
$buttonSaveCSV.Text = "Sauvegarder vers CSV"
$form.Controls.Add($buttonSaveCSV)
$yPos += 40

# Boutons pour la gestion des entrées
$buttonAddClient = New-Object System.Windows.Forms.Button
$buttonAddClient.Location = New-Object System.Drawing.Point(320,$yPos)
$buttonAddClient.Size = New-Object System.Drawing.Size(150,30)
$buttonAddClient.Text = "Ajouter client"
$form.Controls.Add($buttonAddClient)

$buttonRemoveClient = New-Object System.Windows.Forms.Button
$buttonRemoveClient.Location = New-Object System.Drawing.Point(480,$yPos)
$buttonRemoveClient.Size = New-Object System.Drawing.Size(150,30)
$buttonRemoveClient.Text = "Supprimer client"
$form.Controls.Add($buttonRemoveClient)
$yPos += 40

# Séparateur
$separator = New-Object System.Windows.Forms.Label
$separator.Location = New-Object System.Drawing.Point(10,$yPos)
$separator.Size = New-Object System.Drawing.Size(760,2)
$separator.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$form.Controls.Add($separator)
$yPos += 20

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

# Champ pour le nom du client
$labelClient = New-Object System.Windows.Forms.Label
$labelClient.Location = New-Object System.Drawing.Point(10,$yPos)
$labelClient.Size = New-Object System.Drawing.Size(200,20)
$labelClient.Text = "Nom du client:"
$form.Controls.Add($labelClient)

$textClient = New-Object System.Windows.Forms.TextBox
$textClient.Location = New-Object System.Drawing.Point(220,$yPos)
$textClient.Size = New-Object System.Drawing.Size(250,20)
$form.Controls.Add($textClient)
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

# Bouton de test et détection
$buttonTest = New-Object System.Windows.Forms.Button
$buttonTest.Location = New-Object System.Drawing.Point(150,$yPos)
$buttonTest.Size = New-Object System.Drawing.Size(200,30)
$buttonTest.Text = "Tester la connexion et détecter"
$form.Controls.Add($buttonTest)
$yPos += 40

# Bouton de backup de configuration
$buttonBackup = New-Object System.Windows.Forms.Button
$buttonBackup.Location = New-Object System.Drawing.Point(150,$yPos)
$buttonBackup.Size = New-Object System.Drawing.Size(200,30)
$buttonBackup.Text = "Backup de la configuration"
$form.Controls.Add($buttonBackup)
$yPos += 40

# Séparateur
$separator = New-Object System.Windows.Forms.Label
$separator.Location = New-Object System.Drawing.Point(10,$yPos)
$separator.Size = New-Object System.Drawing.Size(760,2)
$separator.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$form.Controls.Add($separator)
$yPos += 20

# Label pour afficher les infos détectées
$labelDetected = New-Object System.Windows.Forms.Label
$labelDetected.Location = New-Object System.Drawing.Point(10,$yPos)
$labelDetected.Size = New-Object System.Drawing.Size(760,40)
$labelDetected.Text = "Aucune information détectée"
$form.Controls.Add($labelDetected)
$yPos += 50

# Label et ComboBox pour la catégorie
$labelCategorie = New-Object System.Windows.Forms.Label
$labelCategorie.Location = New-Object System.Drawing.Point(10,$yPos)
$labelCategorie.Size = New-Object System.Drawing.Size(200,20)
$labelCategorie.Text = "Sélectionnez la catégorie:"
$form.Controls.Add($labelCategorie)

$comboCategorie = New-Object System.Windows.Forms.ComboBox
$comboCategorie.Location = New-Object System.Drawing.Point(220,$yPos)
$comboCategorie.Size = New-Object System.Drawing.Size(250,20)
$comboCategorie.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$modeles.Keys | ForEach-Object { $comboCategorie.Items.Add($_) }
$form.Controls.Add($comboCategorie)
$yPos += 30

# Label et ComboBox pour le modèle spécifique
$labelModele = New-Object System.Windows.Forms.Label
$labelModele.Location = New-Object System.Drawing.Point(10,$yPos)
$labelModele.Size = New-Object System.Drawing.Size(200,20)
$labelModele.Text = "Sélectionnez le modèle:"
$form.Controls.Add($labelModele)

$comboModele = New-Object System.Windows.Forms.ComboBox
$comboModele.Location = New-Object System.Drawing.Point(220,$yPos)
$comboModele.Size = New-Object System.Drawing.Size(250,20)
$comboModele.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$form.Controls.Add($comboModele)
$yPos += 30

# Options de version
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Location = New-Object System.Drawing.Point(10,$yPos)
$labelVersion.Size = New-Object System.Drawing.Size(200,20)
$labelVersion.Text = "Options de version:"
$form.Controls.Add($labelVersion)

$radioDerniere = New-Object System.Windows.Forms.RadioButton
$radioDerniere.Location = New-Object System.Drawing.Point(220,$yPos)
$radioDerniere.Size = New-Object System.Drawing.Size(250,20)
$radioDerniere.Text = "Utiliser la dernière version"
$radioDerniere.Checked = $true
$form.Controls.Add($radioDerniere)
$yPos += 25

$radioChoisir = New-Object System.Windows.Forms.RadioButton
$radioChoisir.Location = New-Object System.Drawing.Point(220,$yPos)
$radioChoisir.Size = New-Object System.Drawing.Size(250,20)
$radioChoisir.Text = "Choisir une version spécifique"
$form.Controls.Add($radioChoisir)
$yPos += 30

# Bouton pour sélectionner le fichier
$buttonSelect = New-Object System.Windows.Forms.Button
$buttonSelect.Location = New-Object System.Drawing.Point(220,$yPos)
$buttonSelect.Size = New-Object System.Drawing.Size(250,30)
$buttonSelect.Text = "Sélectionner le fichier .maj"
$buttonSelect.Enabled = $false
$form.Controls.Add($buttonSelect)
$yPos += 40

# Bouton OK
$buttonOK = New-Object System.Windows.Forms.Button
$buttonOK.Location = New-Object System.Drawing.Point(150,$yPos)
$buttonOK.Size = New-Object System.Drawing.Size(100,30)
$buttonOK.Text = "Lancer la mise à jour"
$buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $buttonOK
$form.Controls.Add($buttonOK)

# Bouton pour lancer la mise à jour pour tous les clients
$buttonUpdateAll = New-Object System.Windows.Forms.Button
$buttonUpdateAll.Location = New-Object System.Drawing.Point(260,$yPos)
$buttonUpdateAll.Size = New-Object System.Drawing.Size(150,30)
$buttonUpdateAll.Text = "Mettre à jour tous"
$form.Controls.Add($buttonUpdateAll)

# Bouton Annuler
$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(420,$yPos)
$buttonCancel.Size = New-Object System.Drawing.Size(100,30)
$buttonCancel.Text = "Annuler"
$buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $buttonCancel
$form.Controls.Add($buttonCancel)

# Gestionnaire d'événements pour la sélection de catégorie
$comboCategorie.Add_SelectedIndexChanged({
    $comboModele.Items.Clear()
    $selectedCategorie = $comboCategorie.SelectedItem
    $modeles[$selectedCategorie] | ForEach-Object { $comboModele.Items.Add($_) }
    $comboModele.SelectedIndex = 0
})

# Gestionnaire d'événements pour le choix de version
$radioChoisir.Add_CheckedChanged({
    $buttonSelect.Enabled = $radioChoisir.Checked
})

# Gestionnaire d'événements pour le bouton de sélection
$buttonSelect.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Sélectionnez le fichier de mise à jour (.maj)"
    $fileDialog.Filter = "Fichiers de mise à jour (*.maj)|*.maj"
    $fileDialog.InitialDirectory = Join-Path $PSScriptRoot "version"
    
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:UpdateFilePath = $fileDialog.FileName
        $buttonSelect.Text = "Fichier sélectionné: " + (Split-Path $script:UpdateFilePath -Leaf)
    }
})

# Gestionnaire d'événements pour le bouton de test
$buttonTest.Add_Click({
    # Récupérer les valeurs des champs
    $IP = $textIP.Text
    $User = $textUser.Text
    $PSWD = $textPSWD.Text
    $Port = $textPort.Text

    # Validation des champs obligatoires
    if ([string]::IsNullOrEmpty($IP) -or [string]::IsNullOrEmpty($User) -or [string]::IsNullOrEmpty($PSWD)) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez remplir l'IP, l'utilisateur et le mot de passe pour le test!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    try {
        # Convert password to secure string and create credentials
        $PASSWORD = ConvertTo-SecureString -String $PSWD -AsPlainText -Force
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential ($User, $PASSWORD)

        # SSH connection
        $sessionParams = @{
            ComputerName = $IP
            Credential   = $Credential
            AcceptKey    = $true
            Port         = $Port
        }

        $sessionssh = New-SSHSession @sessionParams -ErrorAction Stop
        
        # Récupérer les informations du firewall
        $versionInfo = Invoke-SSHCommand -SSHSession $sessionssh -Command "getversion" -ErrorAction Stop
        $licenseInfo = Invoke-SSHCommand -SSHSession $sessionssh -Command "getlicense" -ErrorAction Stop
        $systemInfo = Invoke-SSHCommand -SSHSession $sessionssh -Command "system info" -ErrorAction Stop
        $modelInfo = Invoke-SSHCommand -SSHSession $sessionssh -Command "getmodel" -ErrorAction Stop
        
        Remove-SSHSession -SSHSession $sessionssh | Out-Null

        # Analyser les informations
        $version = $versionInfo.Output | Where-Object { $_ -match "Version" } | Select-Object -First 1
        $model = $systemInfo.Output | Where-Object { $_ -match "Model" } | Select-Object -First 1
        $serial = $systemInfo.Output | Where-Object { $_ -match "Serial" } | Select-Object -First 1
        $licenseStatus = $licenseInfo.Output | Where-Object { $_ -match "Status" } | Select-Object -First 1
        $modelName = $modelInfo.Output | Select-Object -First 1

        # Nettoyer le nom du modèle
        $modelName = $modelName.Trim()
        
        # Déterminer la catégorie
        $category = Get-StormshieldCategory -model $modelName
        
        if ($null -eq $category) {
            $category = "Inconnue"
            $labelDetected.Text = "Modèle détecté: $modelName`nCatégorie: $category`nImpossible de déterminer automatiquement la catégorie"
        } else {
            $labelDetected.Text = "Modèle détecté: $modelName`nCatégorie: $category"
            
            # Sélectionner automatiquement la catégorie et le modèle
            $comboCategorie.SelectedItem = $category
            $comboModele.SelectedItem = $modelName
        }

        # Afficher les informations dans une boîte de dialogue
        $message = @"
Informations du firewall:
$version
Modèle: $modelName
$model
$serial
Catégorie: $category
État de la licence: $licenseStatus
"@

        [System.Windows.Forms.MessageBox]::Show($message, "Résultats du test", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors du test de connexion: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Gestionnaire d'événements pour le bouton de backup
$buttonBackup.Add_Click({
    # Récupérer les valeurs des champs
    $IP = $textIP.Text
    $Client = $textClient.Text
    $User = $textUser.Text
    $PSWD = $textPSWD.Text
    $Port = $textPort.Text

    # Validation des champs obligatoires
    if ([string]::IsNullOrEmpty($IP) -or [string]::IsNullOrEmpty($Client) -or [string]::IsNullOrEmpty($User) -or [string]::IsNullOrEmpty($PSWD)) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez remplir tous les champs obligatoires pour le backup!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    try {
        # Créer le dossier de backup s'il n'existe pas
        $backupDir = Join-Path $PSScriptRoot "backups"
        if (-not (Test-Path -Path $backupDir)) {
            New-Item -ItemType Directory -Path $backupDir -Force
        }

        # Générer le nom de fichier avec la date
        $date = Get-Date -Format "yyyyMMdd-HHmmss"
        $backupFile = "$backupDir\$Client-config-$date.ncfg"

        # Convert password to secure string and create credentials
        $PASSWORD = ConvertTo-SecureString -String $PSWD -AsPlainText -Force
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential ($User, $PASSWORD)
        $WSCPLogin = "$User" + ":" + "$PSWD"

        # Exporter la configuration via SSH
        $sessionParams = @{
            ComputerName = $IP
            Credential   = $Credential
            AcceptKey    = $true
            Port         = $Port
        }

        $sessionssh = New-SSHSession @sessionParams -ErrorAction Stop
        
        # Exécuter la commande d'export
        $exportResult = Invoke-SSHCommand -SSHSession $sessionssh -Command "export configuration export.ncfg" -ErrorAction Stop
        
        # Télécharger le fichier via WinSCP
        & "C:\Program Files (x86)\WinSCP\WinSCP.com" /command `
            "open scp://$WSCPLogin@$IP`:$Port -hostkey=`"*`"" `
            "get `"/usr/Firewall/Update/export.ncfg`" `"$backupFile`"" `
            "exit"
        
        # Supprimer le fichier temporaire sur le firewall
        Invoke-SSHCommand -SSHSession $sessionssh -Command "rm /usr/Firewall/Update/export.ncfg" -ErrorAction SilentlyContinue
        
        Remove-SSHSession -SSHSession $sessionssh | Out-Null

        [System.Windows.Forms.MessageBox]::Show("Backup de la configuration réussi!`nFichier sauvegardé: $backupFile", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lors du backup de la configuration: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Gestionnaire d'événements pour le bouton de chargement CSV
$buttonLoadCSV.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Sélectionnez le fichier CSV des clients"
    $fileDialog.Filter = "Fichiers CSV (*.csv)|*.csv"
    $fileDialog.InitialDirectory = $PSScriptRoot
    
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $clients = Load-ClientsFromCSV -filePath $fileDialog.FileName
        if ($clients) {
            $listClients.Items.Clear()
            $clients | ForEach-Object { $listClients.Items.Add($_) }
            $script:ClientsCSVPath = $fileDialog.FileName
            [System.Windows.Forms.MessageBox]::Show("Clients chargés avec succès!", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
})

# Gestionnaire d'événements pour le bouton de sauvegarde CSV
$buttonSaveCSV.Add_Click({
    if ($listClients.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucun client à sauvegarder!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $fileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $fileDialog.Title = "Enregistrer le fichier CSV des clients"
    $fileDialog.Filter = "Fichiers CSV (*.csv)|*.csv"
    $fileDialog.InitialDirectory = $PSScriptRoot
    $fileDialog.FileName = "clients_stormshield.csv"
    
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $clients = @()
        foreach ($item in $listClients.Items) {
            $clients += $item
        }
        $clients | Export-Csv -Path $fileDialog.FileName -Delimiter ";" -NoTypeInformation -Force
        $script:ClientsCSVPath = $fileDialog.FileName
        [System.Windows.Forms.MessageBox]::Show("Clients sauvegardés avec succès!", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Gestionnaire d'événements pour le bouton d'ajout de client
$buttonAddClient.Add_Click({
    # Vérifier que tous les champs sont remplis
    if ([string]::IsNullOrEmpty($textIP.Text) -or [string]::IsNullOrEmpty($textClient.Text) -or [string]::IsNullOrEmpty($textUser.Text) -or [string]::IsNullOrEmpty($textPSWD.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez remplir tous les champs obligatoires pour ajouter un client!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Créer un nouvel objet client
    $newClient = [PSCustomObject]@{
        IP = $textIP.Text
        Nom = $textClient.Text
        Utilisateur = $textUser.Text
        MotDePasse = $textPSWD.Text
        Port = $textPort.Text
    }

    # Ajouter le client à la liste
    $listClients.Items.Add($newClient)
    
    # Si un fichier CSV est déjà chargé, sauvegarder automatiquement
    if ($script:ClientsCSVPath) {
        Save-ClientToCSV -filePath $script:ClientsCSVPath -client $newClient
    }

    [System.Windows.Forms.MessageBox]::Show("Client ajouté avec succès!", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Gestionnaire d'événements pour le bouton de suppression de client
$buttonRemoveClient.Add_Click({
    if ($listClients.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner un client à supprimer!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $listClients.Items.Remove($listClients.SelectedItem)
    
    # Si un fichier CSV est déjà chargé, sauvegarder automatiquement
    if ($script:ClientsCSVPath) {
        $clients = @()
        foreach ($item in $listClients.Items) {
            $clients += $item
        }
        $clients | Export-Csv -Path $script:ClientsCSVPath -Delimiter ";" -NoTypeInformation -Force
    }

    [System.Windows.Forms.MessageBox]::Show("Client supprimé avec succès!", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Gestionnaire d'événements pour la sélection d'un client dans la liste
$listClients.Add_SelectedIndexChanged({
    if ($listClients.SelectedItem -ne $null) {
        $selectedClient = $listClients.SelectedItem
        $textIP.Text = $selectedClient.IP
        $textClient.Text = $selectedClient.Nom
        $textUser.Text = $selectedClient.Utilisateur
        $textPSWD.Text = $selectedClient.MotDePasse
        $textPort.Text = $selectedClient.Port
    }
})

# Gestionnaire d'événements pour le bouton de mise à jour pour tous les clients
$buttonUpdateAll.Add_Click({
    if ($listClients.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucun client à mettre à jour!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # Vérifier les paramètres de mise à jour
    if ($comboCategorie.SelectedItem -eq $null -or $comboModele.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner une catégorie et un modèle!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Vérifier si un fichier est nécessaire et sélectionné
    if ($radioChoisir.Checked -and (-not $script:UpdateFilePath)) {
        [System.Windows.Forms.MessageBox]::Show("Aucun fichier de mise à jour sélectionné!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Si on utilise la dernière version, déterminer le chemin automatiquement
    if ($radioDerniere.Checked) {
        $categorie = $comboCategorie.SelectedItem
        $modeles = $comboModele.SelectedItem
        
        # Déterminer le sous-dossier en fonction de la catégorie
        $sousDossier = switch ($categorie) {
            "VM" { "VM" }
            "Taille S" { "s" }
            "Taille M" { "m" }
            "Taille L" { "l" }
        }
        
        # Chemin vers le dossier des firmwares
        $firmwareDir = Join-Path $PSScriptRoot "version\$sousDossier"
        
        # Trouver le firmware le plus récent
        $latestFirmware = Get-ChildItem -Path $firmwareDir -Filter "*.maj" | 
                         Sort-Object LastWriteTime -Descending | 
                         Select-Object -First 1
        
        if ($latestFirmware) {
            $script:UpdateFilePath = $latestFirmware.FullName
            Write-Host "Firmware sélectionné: $($latestFirmware.Name)"
        } else {
            [System.Windows.Forms.MessageBox]::Show("Aucun firmware trouvé dans le dossier $firmwareDir", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }

    # Demander confirmation
    $confirm = [System.Windows.Forms.MessageBox]::Show("Êtes-vous sûr de vouloir mettre à jour tous les clients ($($listClients.Items.Count) clients)?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    # Créer un formulaire de progression
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Progression des mises à jour"
    $progressForm.Size = New-Object System.Drawing.Size(500,200)
    $progressForm.StartPosition = "CenterScreen"
    $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false

    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(10,20)
    $progressLabel.Size = New-Object System.Drawing.Size(460,20)
    $progressLabel.Text = "Préparation des mises à jour..."
    $progressForm.Controls.Add($progressLabel)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10,50)
    $progressBar.Size = New-Object System.Drawing.Size(460,20)
    $progressBar.Minimum = 0
    $progressBar.Maximum = $listClients.Items.Count
    $progressForm.Controls.Add($progressBar)

    $progressDetails = New-Object System.Windows.Forms.TextBox
    $progressDetails.Location = New-Object System.Drawing.Point(10,80)
    $progressDetails.Size = New-Object System.Drawing.Size(460,80)
    $progressDetails.Multiline = $true
    $progressDetails.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $progressDetails.ReadOnly = $true
    $progressForm.Controls.Add($progressDetails)

    # Afficher le formulaire de progression
    $progressForm.Show()
    $progressForm.Refresh()

    # Traiter chaque client
    $successCount = 0
    $errorCount = 0
    $currentIndex = 0

    foreach ($client in $listClients.Items) {
        $currentIndex++
        $progressBar.Value = $currentIndex
        $progressLabel.Text = "Traitement du client $currentIndex/$($listClients.Items.Count) - $($client.Nom)"
        $progressDetails.AppendText("Début de la mise à jour pour $($client.Nom) ($($client.IP))...`r`n")
        $progressForm.Refresh()

        try {
            # Convert password to secure string and create credentials
            $PASSWORD = ConvertTo-SecureString -String $client.MotDePasse -AsPlainText -Force
            $Credential = New-Object -TypeName System.Management.Automation.PSCredential ($client.Utilisateur, $PASSWORD)
            $WSCPLogin = "$($client.Utilisateur)" + ":" + "$($client.MotDePasse)"

            ## LOG Firewall version + Date in the format day, month, year, hour, minute
            $date = Get-Date -Format "dd-MM-yyyy-HH-mm"
            $logPath = "C:\logs"
            $logfile = "$logPath\$($client.Nom)-$date.log"

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
            Add-Content -Path $logfile -Value "Client: $($client.Nom)"
            Add-Content -Path $logfile -Value "IP: $($client.IP)"
            Add-Content -Path $logfile -Value "Utilisateur: $($client.Utilisateur)"
            Add-Content -Path $logfile -Value "Modèle sélectionné: $($comboModele.SelectedItem)"
            Add-Content -Path $logfile -Value "Firmware utilisé: $(Split-Path $script:UpdateFilePath -Leaf)"

            try {
                # WinSCP upload command
                & "C:\Program Files (x86)\WinSCP\WinSCP.com" /command `
                    "open scp://$WSCPLogin@$($client.IP)`:$($client.Port) -hostkey=`"*`"" `
                    "put `"$script:UpdateFilePath`" `"/usr/Firewall/Update/`"" `
                    "exit"
                
                Add-Content -Path $logfile -Value "File uploaded successfully"
                $progressDetails.AppendText("Fichier de mise à jour uploadé avec succès`r`n")
            } catch {
                Add-Content -Path $logfile -Value "Error during WinSCP upload: $_"
                $progressDetails.AppendText("ERREUR lors de l'upload du fichier: $_`r`n")
                throw "Erreur upload"
            }

            ## SSH connection + update
            try {
                $sessionParams = @{
                    ComputerName = $client.IP
                    Credential   = $Credential
                    AcceptKey    = $true
                    Port         = $client.Port
                }

                $sessionssh = New-SSHSession @sessionParams -ErrorAction Stop
                
                Add-Content -Path $logfile -Value "------------------------------------------"
                Add-Content -Path $logfile -Value "Firmware version before the update"
                Add-Content -Path $logfile -Value "------------------------------------------"
                
                $preVersion = Invoke-SSHCommand -SSHSession $sessionssh -Command "getversion" -ErrorAction Stop
                Add-Content -Path $logfile -Value $preVersion.Output
                $progressDetails.AppendText("Version actuelle:`r`n$($preVersion.Output)`r`n")
                
                Add-Content -Path $logfile -Value "------------------------------------------"
                Add-Content -Path $logfile -Value "Starting firewall update at $(Get-Date)"
                Add-Content -Path $logfile -Value "------------------------------------------"
                
                $updateResult = Invoke-SSHCommand -SSHSession $sessionssh -Command "fwupdate -r -f /usr/Firewall/Update/$(Split-Path $script:UpdateFilePath -Leaf)" -ErrorAction Stop
                Add-Content -Path $logfile -Value $updateResult.Output
                $progressDetails.AppendText("Commande de mise à jour envoyée`r`n")
                
                Remove-SSHSession -SSHSession $sessionssh | Out-Null
                
                Add-Content -Path $logfile -Value "Update command sent successfully. Waiting for 10 minutes..."
                $progressDetails.AppendText("Mise à jour lancée, attente de 10 minutes...`r`n")
                $progressForm.Refresh()
                
                ## Sleep for 10 mins needed to let the update do its job
                Start-Sleep -Seconds 600
                
                ## Verify update
                $sessionssh = New-SSHSession @sessionParams -ErrorAction Stop
                
                Add-Content -Path $logfile -Value "------------------------------------------"
                Add-Content -Path $logfile -Value "Firmware version after the update"
                Add-Content -Path $logfile -Value "------------------------------------------"
                
                $postVersion = Invoke-SSHCommand -SSHSession $sessionssh -Command "getversion" -ErrorAction Stop
                Add-Content -Path $logfile -Value $postVersion.Output
                $progressDetails.AppendText("Nouvelle version:`r`n$($postVersion.Output)`r`n")
                
                Remove-SSHSession -SSHSession $sessionssh | Out-Null
                
                Add-Content -Path $logfile -Value "------------------------------------------"
                Add-Content -Path $logfile -Value "Script completed successfully at $(Get-Date)"
                Add-Content -Path $logfile -Value "------------------------------------------"
                
                $progressDetails.AppendText("Mise à jour terminée avec succès!`r`nVoir le log: $logfile`r`n")
                $successCount++
            } catch {
                Add-Content -Path $logfile -Value "Error during SSH operations: $_"
                $progressDetails.AppendText("ERREUR lors des opérations SSH: $_`r`n")
                $errorCount++
            }
        } catch {
            $progressDetails.AppendText("ERREUR lors du traitement du client $($client.Nom): $_`r`n")
            $errorCount++
        }
        
        $progressForm.Refresh()
    }

    # Fermer le formulaire de progression
    $progressForm.Close()

    # Afficher un récapitulatif
    $message = "Mises à jour terminées!`r`n"
    $message += "Clients traités: $($listClients.Items.Count)`r`n"
    $message += "Mises à jour réussies: $successCount`r`n"
    $message += "Échecs: $errorCount"
    
    [System.Windows.Forms.MessageBox]::Show($message, "Récapitulatif", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Gestionnaire d'événements pour le bouton OK (mise à jour d'un seul client)
$buttonOK.Add_Click({
    # Récupérer les valeurs des champs
    $IP = $textIP.Text
    $Client = $textClient.Text
    $User = $textUser.Text
    $PSWD = $textPSWD.Text
    $Port = $textPort.Text

    # Validation des champs obligatoires
    if ([string]::IsNullOrEmpty($IP) -or [string]::IsNullOrEmpty($Client) -or [string]::IsNullOrEmpty($User) -or [string]::IsNullOrEmpty($PSWD)) {
        [System.Windows.Forms.MessageBox]::Show("Tous les champs obligatoires doivent être remplis!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Vérifier si un fichier est nécessaire et sélectionné
    if ($radioChoisir.Checked -and (-not $script:UpdateFilePath)) {
        [System.Windows.Forms.MessageBox]::Show("Aucun fichier de mise à jour sélectionné!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Si on utilise la dernière version, déterminer le chemin automatiquement
    if ($radioDerniere.Checked) {
        $categorie = $comboCategorie.SelectedItem
        $modeles = $comboModele.SelectedItem
        
        # Déterminer le sous-dossier en fonction de la catégorie
        $sousDossier = switch ($categorie) {
            "VM" { "VM" }
            "Taille S" { "s" }
            "Taille M" { "m" }
            "Taille L" { "l" }
        }
        
        # Chemin vers le dossier des firmwares
        $firmwareDir = Join-Path $PSScriptRoot "version\$sousDossier"
        
        # Trouver le firmware le plus récent
        $latestFirmware = Get-ChildItem -Path $firmwareDir -Filter "*.maj" | 
                         Sort-Object LastWriteTime -Descending | 
                         Select-Object -First 1
        
        if ($latestFirmware) {
            $script:UpdateFilePath = $latestFirmware.FullName
            Write-Host "Firmware sélectionné: $($latestFirmware.Name)"
        } else {
            [System.Windows.Forms.MessageBox]::Show("Aucun firmware trouvé dans le dossier $firmwareDir", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }

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
    Add-Content -Path $logfile -Value "Client: $Client"
    Add-Content -Path $logfile -Value "IP: $IP"
    Add-Content -Path $logfile -Value "Utilisateur: $User"
    Add-Content -Path $logfile -Value "Modèle sélectionné: $($comboModele.SelectedItem)"
    Add-Content -Path $logfile -Value "Firmware utilisé: $(Split-Path $script:UpdateFilePath -Leaf)"

    try {
        # WinSCP upload command
        & "C:\Program Files (x86)\WinSCP\WinSCP.com" /command `
            "open scp://$WSCPLogin@$IP`:$Port -hostkey=`"*`"" `
            "put `"$script:UpdateFilePath`" `"/usr/Firewall/Update/`"" `
            "exit"
        
        Add-Content -Path $logfile -Value "File uploaded successfully"
    } catch {
        Add-Content -Path $logfile -Value "Error during WinSCP upload: $_"
        [System.Windows.Forms.MessageBox]::Show("Erreur lors de l'upload du fichier: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }

    ## SSH connection + update
    try {
        $sessionParams = @{
            ComputerName = $IP
            Credential   = $Credential
            AcceptKey    = $true
            Port         = $Port
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
        
        $updateResult = Invoke-SSHCommand -SSHSession $sessionssh -Command "fwupdate -r -f /usr/Firewall/Update/$(Split-Path $script:UpdateFilePath -Leaf)" -ErrorAction Stop
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
        
        # Afficher un message de succès
        [System.Windows.Forms.MessageBox]::Show("Mise à jour terminée avec succès!`nVoir le log: $logfile", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
    } catch {
        Add-Content -Path $logfile -Value "Error during SSH operations: $_"
        [System.Windows.Forms.MessageBox]::Show("Erreur lors des opérations SSH: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }

    Add-Content -Path $logfile -Value "------------------------------------------"
    Add-Content -Path $logfile -Value "Script completed successfully at $(Get-Date)"
    Add-Content -Path $logfile -Value "------------------------------------------"
})

# Afficher le formulaire et attendre la sélection
$result = $form.ShowDialog()