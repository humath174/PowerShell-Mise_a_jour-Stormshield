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

# Chemin du fichier CSV par défaut
$defaultCsvPath = Join-Path $PSScriptRoot "clients_stormshield.csv"

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
        if (Test-Path -Path $filePath) {
            $clients = Import-Csv -Path $filePath -Delimiter ";"
            return $clients
        }
        return @()
    } catch {
        Write-Error "Erreur lors du chargement du fichier CSV: $_"
        return @()
    }
}

# Fonction pour sauvegarder les clients dans un fichier CSV
function Save-ClientsToCSV {
    param (
        [string]$filePath,
        [array]$clients
    )
    
    try {
        $clients | Export-Csv -Path $filePath -Delimiter ";" -NoTypeInformation -Force
        return $true
    } catch {
        Write-Error "Erreur lors de la sauvegarde du fichier CSV: $_"
        return $false
    }
}

# Fonction pour tester la connexion à un firewall
function Test-FirewallConnection {
    param (
        [PSCustomObject]$client
    )
    
    try {
        # Convert password to secure string and create credentials
        $PASSWORD = ConvertTo-SecureString -String $client.MotDePasse -AsPlainText -Force
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential ($client.Utilisateur, $PASSWORD)

        # SSH connection parameters
        $sessionParams = @{
            ComputerName = $client.IP
            Credential   = $Credential
            AcceptKey    = $true
            Port         = $client.Port
        }

        # Test SSH connection
        $sessionssh = New-SSHSession @sessionParams -ErrorAction Stop
        
        # Get basic info
        $versionInfo = Invoke-SSHCommand -SSHSession $sessionssh -Command "getversion" -ErrorAction Stop
        $modelInfo = Invoke-SSHCommand -SSHSession $sessionssh -Command "getmodel" -ErrorAction Stop
        
        Remove-SSHSession -SSHSession $sessionssh | Out-Null

        # Parse model info
        $modelName = ($modelInfo.Output | Select-Object -First 1).Trim()
        
        return @{
            Status = "Success"
            Model = $modelName
            Version = ($versionInfo.Output | Where-Object { $_ -match "Version" } | Select-Object -First 1)
        }
    } catch {
        return @{
            Status = "Error"
            Message = $_.Exception.Message
        }
    }
}

# Créer le formulaire principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Mise à jour de firmware Stormshield"
$form.Size = New-Object System.Drawing.Size(900,700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Position verticale courante pour les contrôles
$yPos = 10

# Liste des clients avec couleur
$listClients = New-Object System.Windows.Forms.ListView
$listClients.Location = New-Object System.Drawing.Point(10, $yPos)
$listClients.Size = New-Object System.Drawing.Size(600, 200)
$listClients.View = [System.Windows.Forms.View]::Details
$listClients.FullRowSelect = $true
$listClients.GridLines = $true
$listClients.MultiSelect = $true

# Ajouter les colonnes
$listClients.Columns.Add("Client", 150) | Out-Null
$listClients.Columns.Add("IP", 120) | Out-Null
$listClients.Columns.Add("Modèle", 150) | Out-Null
$listClients.Columns.Add("Version", 150) | Out-Null
$listClients.Columns.Add("Statut", 100) | Out-Null

$form.Controls.Add($listClients)
$yPos += 210

# Boutons pour la gestion des clients
$buttonTestAll = New-Object System.Windows.Forms.Button
$buttonTestAll.Location = New-Object System.Drawing.Point(620, 10)
$buttonTestAll.Size = New-Object System.Drawing.Size(150, 30)
$buttonTestAll.Text = "Tester tous"
$form.Controls.Add($buttonTestAll)

$buttonAddClient = New-Object System.Windows.Forms.Button
$buttonAddClient.Location = New-Object System.Drawing.Point(620, 50)
$buttonAddClient.Size = New-Object System.Drawing.Size(150, 30)
$buttonAddClient.Text = "Ajouter client"
$form.Controls.Add($buttonAddClient)

$buttonRemoveClient = New-Object System.Windows.Forms.Button
$buttonRemoveClient.Location = New-Object System.Drawing.Point(620, 90)
$buttonRemoveClient.Size = New-Object System.Drawing.Size(150, 30)
$buttonRemoveClient.Text = "Supprimer client(s)"
$form.Controls.Add($buttonRemoveClient)

$buttonSave = New-Object System.Windows.Forms.Button
$buttonSave.Location = New-Object System.Drawing.Point(620, 130)
$buttonSave.Size = New-Object System.Drawing.Size(150, 30)
$buttonSave.Text = "Sauvegarder"
$form.Controls.Add($buttonSave)

$buttonLoad = New-Object System.Windows.Forms.Button
$buttonLoad.Location = New-Object System.Drawing.Point(620, 170)
$buttonLoad.Size = New-Object System.Drawing.Size(150, 30)
$buttonLoad.Text = "Charger"
$form.Controls.Add($buttonLoad)

# Séparateur
$separator = New-Object System.Windows.Forms.Label
$separator.Location = New-Object System.Drawing.Point(10, $yPos)
$separator.Size = New-Object System.Drawing.Size(860, 2)
$separator.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$form.Controls.Add($separator)
$yPos += 20

# Détails du client sélectionné
$labelDetails = New-Object System.Windows.Forms.Label
$labelDetails.Location = New-Object System.Drawing.Point(10, $yPos)
$labelDetails.Size = New-Object System.Drawing.Size(200, 20)
$labelDetails.Text = "Détails du client sélectionné:"
$form.Controls.Add($labelDetails)
$yPos += 25

# Champ pour l'adresse IP
$labelIP = New-Object System.Windows.Forms.Label
$labelIP.Location = New-Object System.Drawing.Point(10, $yPos)
$labelIP.Size = New-Object System.Drawing.Size(150, 20)
$labelIP.Text = "Adresse IP:"
$form.Controls.Add($labelIP)

$textIP = New-Object System.Windows.Forms.TextBox
$textIP.Location = New-Object System.Drawing.Point(170, $yPos)
$textIP.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textIP)

$labelPort = New-Object System.Windows.Forms.Label
$labelPort.Location = New-Object System.Drawing.Point(380, $yPos)
$labelPort.Size = New-Object System.Drawing.Size(50, 20)
$labelPort.Text = "Port:"
$form.Controls.Add($labelPort)

$textPort = New-Object System.Windows.Forms.TextBox
$textPort.Location = New-Object System.Drawing.Point(440, $yPos)
$textPort.Size = New-Object System.Drawing.Size(80, 20)
$textPort.Text = "13422"
$form.Controls.Add($textPort)
$yPos += 30

# Champ pour le nom du client
$labelClient = New-Object System.Windows.Forms.Label
$labelClient.Location = New-Object System.Drawing.Point(10, $yPos)
$labelClient.Size = New-Object System.Drawing.Size(150, 20)
$labelClient.Text = "Nom du client:"
$form.Controls.Add($labelClient)

$textClient = New-Object System.Windows.Forms.TextBox
$textClient.Location = New-Object System.Drawing.Point(170, $yPos)
$textClient.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($textClient)
$yPos += 30

# Champ pour l'utilisateur admin
$labelUser = New-Object System.Windows.Forms.Label
$labelUser.Location = New-Object System.Drawing.Point(10, $yPos)
$labelUser.Size = New-Object System.Drawing.Size(150, 20)
$labelUser.Text = "Utilisateur admin:"
$form.Controls.Add($labelUser)

$textUser = New-Object System.Windows.Forms.TextBox
$textUser.Location = New-Object System.Drawing.Point(170, $yPos)
$textUser.Size = New-Object System.Drawing.Size(200, 20)
$textUser.Text = "admin"
$form.Controls.Add($textUser)
$yPos += 30

# Champ pour le mot de passe
$labelPSWD = New-Object System.Windows.Forms.Label
$labelPSWD.Location = New-Object System.Drawing.Point(10, $yPos)
$labelPSWD.Size = New-Object System.Drawing.Size(150, 20)
$labelPSWD.Text = "Mot de passe:"
$form.Controls.Add($labelPSWD)

$textPSWD = New-Object System.Windows.Forms.TextBox
$textPSWD.Location = New-Object System.Drawing.Point(170, $yPos)
$textPSWD.Size = New-Object System.Drawing.Size(200, 20)
$textPSWD.PasswordChar = '*'
$form.Controls.Add($textPSWD)
$yPos += 40

# Bouton de test individuel
$buttonTest = New-Object System.Windows.Forms.Button
$buttonTest.Location = New-Object System.Drawing.Point(10, $yPos)
$buttonTest.Size = New-Object System.Drawing.Size(200, 30)
$buttonTest.Text = "Tester ce client"
$form.Controls.Add($buttonTest)

# Bouton de backup
$buttonBackup = New-Object System.Windows.Forms.Button
$buttonBackup.Location = New-Object System.Drawing.Point(220, $yPos)
$buttonBackup.Size = New-Object System.Drawing.Size(200, 30)
$buttonBackup.Text = "Backup configuration"
$form.Controls.Add($buttonBackup)
$yPos += 40

# Séparateur
$separator = New-Object System.Windows.Forms.Label
$separator.Location = New-Object System.Drawing.Point(10, $yPos)
$separator.Size = New-Object System.Drawing.Size(860, 2)
$separator.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$form.Controls.Add($separator)
$yPos += 20

# Section mise à jour
$labelUpdate = New-Object System.Windows.Forms.Label
$labelUpdate.Location = New-Object System.Drawing.Point(10, $yPos)
$labelUpdate.Size = New-Object System.Drawing.Size(200, 20)
$labelUpdate.Text = "Options de mise à jour:"
$form.Controls.Add($labelUpdate)
$yPos += 25

# Catégorie et modèle
$labelCategorie = New-Object System.Windows.Forms.Label
$labelCategorie.Location = New-Object System.Drawing.Point(10, $yPos)
$labelCategorie.Size = New-Object System.Drawing.Size(150, 20)
$labelCategorie.Text = "Catégorie:"
$form.Controls.Add($labelCategorie)

$comboCategorie = New-Object System.Windows.Forms.ComboBox
$comboCategorie.Location = New-Object System.Drawing.Point(170, $yPos)
$comboCategorie.Size = New-Object System.Drawing.Size(200, 20)
$comboCategorie.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$modeles.Keys | ForEach-Object { $comboCategorie.Items.Add($_) }
$form.Controls.Add($comboCategorie)

$labelModele = New-Object System.Windows.Forms.Label
$labelModele.Location = New-Object System.Drawing.Point(380, $yPos)
$labelModele.Size = New-Object System.Drawing.Size(150, 20)
$labelModele.Text = "Modèle:"
$form.Controls.Add($labelModele)

$comboModele = New-Object System.Windows.Forms.ComboBox
$comboModele.Location = New-Object System.Drawing.Point(440, $yPos)
$comboModele.Size = New-Object System.Drawing.Size(200, 20)
$comboModele.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$form.Controls.Add($comboModele)
$yPos += 30

# Options de version
$radioDerniere = New-Object System.Windows.Forms.RadioButton
$radioDerniere.Location = New-Object System.Drawing.Point(10, $yPos)
$radioDerniere.Size = New-Object System.Drawing.Size(250, 20)
$radioDerniere.Text = "Utiliser la dernière version"
$radioDerniere.Checked = $true
$form.Controls.Add($radioDerniere)

$radioChoisir = New-Object System.Windows.Forms.RadioButton
$radioChoisir.Location = New-Object System.Drawing.Point(260, $yPos)
$radioChoisir.Size = New-Object System.Drawing.Size(250, 20)
$radioChoisir.Text = "Choisir une version spécifique"
$form.Controls.Add($radioChoisir)
$yPos += 25

# Bouton pour sélectionner le fichier
$buttonSelect = New-Object System.Windows.Forms.Button
$buttonSelect.Location = New-Object System.Drawing.Point(10, $yPos)
$buttonSelect.Size = New-Object System.Drawing.Size(300, 30)
$buttonSelect.Text = "Sélectionner le fichier .maj"
$buttonSelect.Enabled = $false
$form.Controls.Add($buttonSelect)
$yPos += 40

# Boutons d'action
$buttonUpdateSelected = New-Object System.Windows.Forms.Button
$buttonUpdateSelected.Location = New-Object System.Drawing.Point(10, $yPos)
$buttonUpdateSelected.Size = New-Object System.Drawing.Size(200, 30)
$buttonUpdateSelected.Text = "Mettre à jour sélection"
$form.Controls.Add($buttonUpdateSelected)

$buttonUpdateAllOK = New-Object System.Windows.Forms.Button
$buttonUpdateAllOK.Location = New-Object System.Drawing.Point(220, $yPos)
$buttonUpdateAllOK.Size = New-Object System.Drawing.Size(200, 30)
$buttonUpdateAllOK.Text = "Mettre à jour tous (OK)"
$form.Controls.Add($buttonUpdateAllOK)

$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(430, $yPos)
$buttonCancel.Size = New-Object System.Drawing.Size(100, 30)
$buttonCancel.Text = "Fermer"
$form.Controls.Add($buttonCancel)

# Gestionnaire d'événements pour la sélection de catégorie
$comboCategorie.Add_SelectedIndexChanged({
    $comboModele.Items.Clear()
    $selectedCategorie = $comboCategorie.SelectedItem
    $modeles[$selectedCategorie] | ForEach-Object { $comboModele.Items.Add($_) }
    if ($comboModele.Items.Count -gt 0) {
        $comboModele.SelectedIndex = 0
    }
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

# Gestionnaire d'événements pour le bouton de test individuel
$buttonTest.Add_Click({
    if ($listClients.SelectedItems.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner un seul client à tester", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $selectedItem = $listClients.SelectedItems[0]
    $client = [PSCustomObject]@{
        IP = $selectedItem.SubItems[1].Text
        Nom = $selectedItem.Text
        Utilisateur = $textUser.Text
        MotDePasse = $textPSWD.Text
        Port = $textPort.Text
    }

    $result = Test-FirewallConnection -client $client

    if ($result.Status -eq "Success") {
        $selectedItem.SubItems[2].Text = $result.Model
        $selectedItem.SubItems[3].Text = $result.Version
        $selectedItem.SubItems[4].Text = "OK"
        $selectedItem.BackColor = [System.Drawing.Color]::LightGreen
    } else {
        $selectedItem.SubItems[4].Text = "Erreur"
        $selectedItem.BackColor = [System.Drawing.Color]::LightCoral
    }
})

# Gestionnaire d'événements pour le bouton de test tous
$buttonTestAll.Add_Click({
    if ($listClients.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Aucun client à tester", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # Créer un formulaire de progression
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Test des connexions"
    $progressForm.Size = New-Object System.Drawing.Size(400, 150)
    $progressForm.StartPosition = "CenterScreen"
    $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false

    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(10, 20)
    $progressLabel.Size = New-Object System.Drawing.Size(380, 20)
    $progressLabel.Text = "Test des connexions en cours..."
    $progressForm.Controls.Add($progressLabel)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 50)
    $progressBar.Size = New-Object System.Drawing.Size(380, 20)
    $progressBar.Minimum = 0
    $progressBar.Maximum = $listClients.Items.Count
    $progressForm.Controls.Add($progressBar)

    # Afficher le formulaire de progression
    $progressForm.Show()
    $progressForm.Refresh()

    # Tester chaque client
    for ($i = 0; $i -lt $listClients.Items.Count; $i++) {
        $item = $listClients.Items[$i]
        $progressBar.Value = $i + 1
        $progressLabel.Text = "Test de $($item.Text) ($($item.SubItems[1].Text))..."
        $progressForm.Refresh()

        $client = [PSCustomObject]@{
            IP = $item.SubItems[1].Text
            Nom = $item.Text
            Utilisateur = $textUser.Text # Utilise les credentials du formulaire
            MotDePasse = $textPSWD.Text
            Port = $textPort.Text
        }

        $result = Test-FirewallConnection -client $client

        if ($result.Status -eq "Success") {
            $item.SubItems[2].Text = $result.Model
            $item.SubItems[3].Text = $result.Version
            $item.SubItems[4].Text = "OK"
            $item.BackColor = [System.Drawing.Color]::LightGreen
        } else {
            $item.SubItems[4].Text = "Erreur"
            $item.BackColor = [System.Drawing.Color]::LightCoral
        }

        $listClients.Refresh()
    }

    $progressForm.Close()
    [System.Windows.Forms.MessageBox]::Show("Test des connexions terminé", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Gestionnaire d'événements pour le bouton d'ajout de client
$buttonAddClient.Add_Click({
    if ([string]::IsNullOrEmpty($textClient.Text) -or [string]::IsNullOrEmpty($textIP.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez saisir au moins un nom et une adresse IP", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $item = New-Object System.Windows.Forms.ListViewItem($textClient.Text)
    $item.SubItems.Add($textIP.Text) | Out-Null
    $item.SubItems.Add("") | Out-Null # Modèle
    $item.SubItems.Add("") | Out-Null # Version
    $item.SubItems.Add("Non testé") | Out-Null # Statut
    $listClients.Items.Add($item)
})

# Gestionnaire d'événements pour le bouton de suppression
$buttonRemoveClient.Add_Click({
    if ($listClients.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner au moins un client à supprimer", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    foreach ($item in $listClients.SelectedItems) {
        $listClients.Items.Remove($item)
    }
})

# Gestionnaire d'événements pour le bouton de sauvegarde
$buttonSave.Add_Click({
    $clients = @()
    foreach ($item in $listClients.Items) {
        $client = [PSCustomObject]@{
            Nom = $item.Text
            IP = $item.SubItems[1].Text
            Modele = $item.SubItems[2].Text
            Version = $item.SubItems[3].Text
            Statut = $item.SubItems[4].Text
        }
        $clients += $client
    }

    if (Save-ClientsToCSV -filePath $defaultCsvPath -clients $clients) {
        [System.Windows.Forms.MessageBox]::Show("Clients sauvegardés avec succès", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Gestionnaire d'événements pour le bouton de chargement
$buttonLoad.Add_Click({
    $clients = Load-ClientsFromCSV -filePath $defaultCsvPath
    $listClients.Items.Clear()
    
    foreach ($client in $clients) {
        $item = New-Object System.Windows.Forms.ListViewItem($client.Nom)
        $item.SubItems.Add($client.IP) | Out-Null
        $item.SubItems.Add($client.Modele) | Out-Null
        $item.SubItems.Add($client.Version) | Out-Null
        $item.SubItems.Add($client.Statut) | Out-Null
        
        # Colorer en fonction du statut
        if ($client.Statut -eq "OK") {
            $item.BackColor = [System.Drawing.Color]::LightGreen
        } elseif ($client.Statut -eq "Erreur") {
            $item.BackColor = [System.Drawing.Color]::LightCoral
        }
        
        $listClients.Items.Add($item)
    }
})

# Gestionnaire d'événements pour la sélection dans la liste
$listClients.Add_SelectedIndexChanged({
    if ($listClients.SelectedItems.Count -eq 1) {
        $selectedItem = $listClients.SelectedItems[0]
        $textClient.Text = $selectedItem.Text
        $textIP.Text = $selectedItem.SubItems[1].Text
        $textPort.Text = "13422" # Réinitialiser le port par défaut
    }
})

# Gestionnaire d'événements pour le bouton de backup
$buttonBackup.Add_Click({
    if ($listClients.SelectedItems.Count -ne 1) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner un seul client", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $selectedItem = $listClients.SelectedItems[0]
    $client = [PSCustomObject]@{
        IP = $selectedItem.SubItems[1].Text
        Nom = $selectedItem.Text
        Utilisateur = $textUser.Text
        MotDePasse = $textPSWD.Text
        Port = $textPort.Text
    }

    try {
        # Créer le dossier de backup s'il n'existe pas
        $backupDir = Join-Path $PSScriptRoot "backups"
        if (-not (Test-Path -Path $backupDir)) {
            New-Item -ItemType Directory -Path $backupDir -Force
        }

        # Générer le nom de fichier avec la date
        $date = Get-Date -Format "yyyyMMdd-HHmmss"
        $backupFile = "$backupDir\$($client.Nom)-config-$date.ncfg"

        # Convert password to secure string and create credentials
        $PASSWORD = ConvertTo-SecureString -String $client.MotDePasse -AsPlainText -Force
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential ($client.Utilisateur, $PASSWORD)
        $WSCPLogin = "$($client.Utilisateur)" + ":" + "$($client.MotDePasse)"

        # Exporter la configuration via SSH
        $sessionParams = @{
            ComputerName = $client.IP
            Credential   = $Credential
            AcceptKey    = $true
            Port         = $client.Port
        }

        $sessionssh = New-SSHSession @sessionParams -ErrorAction Stop
        
        # Exécuter la commande d'export
        $exportResult = Invoke-SSHCommand -SSHSession $sessionssh -Command "export configuration export.ncfg" -ErrorAction Stop
        
        # Télécharger le fichier via WinSCP
        & "C:\Program Files (x86)\WinSCP\WinSCP.com" /command `
            "open scp://$WSCPLogin@$($client.IP)`:$($client.Port) -hostkey=`"*`"" `
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

# Gestionnaire d'événements pour le bouton de mise à jour sélection
$buttonUpdateSelected.Add_Click({
    if ($listClients.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner au moins un client", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # Vérifier les paramètres de mise à jour
    if ($comboCategorie.SelectedItem -eq $null -or $comboModele.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner une catégorie et un modèle", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Vérifier si un fichier est nécessaire et sélectionné
    if ($radioChoisir.Checked -and (-not $script:UpdateFilePath)) {
        [System.Windows.Forms.MessageBox]::Show("Aucun fichier de mise à jour sélectionné", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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
        } else {
            [System.Windows.Forms.MessageBox]::Show("Aucun firmware trouvé dans le dossier $firmwareDir", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }

    # Créer un formulaire de progression
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Mise à jour en cours"
    $progressForm.Size = New-Object System.Drawing.Size(500, 200)
    $progressForm.StartPosition = "CenterScreen"
    $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false

    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(10, 20)
    $progressLabel.Size = New-Object System.Drawing.Size(480, 20)
    $progressForm.Controls.Add($progressLabel)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 50)
    $progressBar.Size = New-Object System.Drawing.Size(480, 20)
    $progressBar.Minimum = 0
    $progressBar.Maximum = $listClients.SelectedItems.Count
    $progressForm.Controls.Add($progressBar)

    $progressDetails = New-Object System.Windows.Forms.TextBox
    $progressDetails.Location = New-Object System.Drawing.Point(10, 80)
    $progressDetails.Size = New-Object System.Drawing.Size(480, 100)
    $progressDetails.Multiline = $true
    $progressDetails.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $progressDetails.ReadOnly = $true
    $progressForm.Controls.Add($progressDetails)

    # Afficher le formulaire de progression
    $progressForm.Show()
    $progressForm.Refresh()

    $successCount = 0
    $errorCount = 0
    $currentIndex = 0

    foreach ($item in $listClients.SelectedItems) {
        $currentIndex++
        $progressBar.Value = $currentIndex
        $progressLabel.Text = "Mise à jour de $($item.Text) ($($item.SubItems[1].Text))..."
        $progressForm.Refresh()

        $client = [PSCustomObject]@{
            IP = $item.SubItems[1].Text
            Nom = $item.Text
            Utilisateur = $textUser.Text
            MotDePasse = $textPSWD.Text
            Port = $textPort.Text
        }

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
            Get
                        # SFTP Connection + upload the update file
                        $progressDetails.AppendText("Upload du fichier de mise à jour...`r`n")
                        $progressForm.Refresh()
            
                        try {
                            # WinSCP upload command
                            & "C:\Program Files (x86)\WinSCP\WinSCP.com" /command `
                                "open scp://$WSCPLogin@$($client.IP)`:$($client.Port) -hostkey=`"*`"" `
                                "put `"$script:UpdateFilePath`" `"/usr/Firewall/Update/`"" `
                                "exit"
                            
                            $progressDetails.AppendText("Fichier uploadé avec succès`r`n")
                        } catch {
                            $progressDetails.AppendText("ERREUR lors de l'upload: $_`r`n")
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
                            
                            $progressDetails.AppendText("Récupération de la version actuelle...`r`n")
                            $preVersion = Invoke-SSHCommand -SSHSession $sessionssh -Command "getversion" -ErrorAction Stop
                            
                            $progressDetails.AppendText("Lancement de la mise à jour...`r`n")
                            $updateResult = Invoke-SSHCommand -SSHSession $sessionssh -Command "fwupdate -r -f /usr/Firewall/Update/$(Split-Path $script:UpdateFilePath -Leaf)" -ErrorAction Stop
                            
                            Remove-SSHSession -SSHSession $sessionssh | Out-Null
                            
                            $progressDetails.AppendText("Mise à jour lancée, attente de 10 minutes...`r`n")
                            $progressForm.Refresh()
                            
                            ## Sleep for 10 mins needed to let the update do its job
                            Start-Sleep -Seconds 600
                            
                            ## Verify update
                            $sessionssh = New-SSHSession @sessionParams -ErrorAction Stop
                            $postVersion = Invoke-SSHCommand -SSHSession $sessionssh -Command "getversion" -ErrorAction Stop
                            Remove-SSHSession -SSHSession $sessionssh | Out-Null
                            
                            $progressDetails.AppendText("Mise à jour terminée avec succès!`r`n")
                            $progressDetails.AppendText("Ancienne version: $($preVersion.Output)`r`n")
                            $progressDetails.AppendText("Nouvelle version: $($postVersion.Output)`r`n")
                            
                            $item.SubItems[3].Text = ($postVersion.Output | Where-Object { $_ -match "Version" } | Select-Object -First 1)
                            $item.SubItems[4].Text = "Mis à jour"
                            $item.BackColor = [System.Drawing.Color]::LightGreen
                            $successCount++
                        } catch {
                            $progressDetails.AppendText("ERREUR lors de la mise à jour: $_`r`n")
                            $item.SubItems[4].Text = "Échec mise à jour"
                            $item.BackColor = [System.Drawing.Color]::LightCoral
                            $errorCount++
                        }
                    } catch {
                        $progressDetails.AppendText("ERREUR lors du traitement: $_`r`n")
                        $item.SubItems[4].Text = "Erreur"
                        $item.BackColor = [System.Drawing.Color]::LightCoral
                        $errorCount++
                    }
                    
                    $listClients.Refresh()
                    $progressForm.Refresh()
                }
            
                # Fermer le formulaire de progression
                $progressForm.Close()
            
                # Afficher un récapitulatif
                $message = "Mises à jour terminées!`r`n"
                $message += "Clients traités: $($listClients.SelectedItems.Count)`r`n"
                $message += "Mises à jour réussies: $successCount`r`n"
                $message += "Échecs: $errorCount"
                
                [System.Windows.Forms.MessageBox]::Show($message, "Récapitulatif", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            })
            
            # Gestionnaire d'événements pour le bouton de mise à jour tous (OK)
            $buttonUpdateAllOK.Add_Click({
                # Filtrer seulement les clients qui ont un statut OK
                $clientsToUpdate = @()
                foreach ($item in $listClients.Items) {
                    if ($item.SubItems[4].Text -eq "OK") {
                        $clientsToUpdate += $item
                    }
                }
            
                if ($clientsToUpdate.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("Aucun client avec statut OK à mettre à jour", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    return
                }
            
                # Vérifier les paramètres de mise à jour
                if ($comboCategorie.SelectedItem -eq $null -or $comboModele.SelectedItem -eq $null) {
                    [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner une catégorie et un modèle", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
            
                # Vérifier si un fichier est nécessaire et sélectionné
                if ($radioChoisir.Checked -and (-not $script:UpdateFilePath)) {
                    [System.Windows.Forms.MessageBox]::Show("Aucun fichier de mise à jour sélectionné", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Aucun firmware trouvé dans le dossier $firmwareDir", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        return
                    }
                }
            
                # Créer un formulaire de progression
                $progressForm = New-Object System.Windows.Forms.Form
                $progressForm.Text = "Mise à jour en cours"
                $progressForm.Size = New-Object System.Drawing.Size(500, 200)
                $progressForm.StartPosition = "CenterScreen"
                $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
                $progressForm.MaximizeBox = $false
                $progressForm.MinimizeBox = $false
            
                $progressLabel = New-Object System.Windows.Forms.Label
                $progressLabel.Location = New-Object System.Drawing.Point(10, 20)
                $progressLabel.Size = New-Object System.Drawing.Size(480, 20)
                $progressForm.Controls.Add($progressLabel)
            
                $progressBar = New-Object System.Windows.Forms.ProgressBar
                $progressBar.Location = New-Object System.Drawing.Point(10, 50)
                $progressBar.Size = New-Object System.Drawing.Size(480, 20)
                $progressBar.Minimum = 0
                $progressBar.Maximum = $clientsToUpdate.Count
                $progressForm.Controls.Add($progressBar)
            
                $progressDetails = New-Object System.Windows.Forms.TextBox
                $progressDetails.Location = New-Object System.Drawing.Point(10, 80)
                $progressDetails.Size = New-Object System.Drawing.Size(480, 100)
                $progressDetails.Multiline = $true
                $progressDetails.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
                $progressDetails.ReadOnly = $true
                $progressForm.Controls.Add($progressDetails)
            
                # Afficher le formulaire de progression
                $progressForm.Show()
                $progressForm.Refresh()
            
                $successCount = 0
                $errorCount = 0
                $currentIndex = 0
            
                foreach ($item in $clientsToUpdate) {
                    $currentIndex++
                    $progressBar.Value = $currentIndex
                    $progressLabel.Text = "Mise à jour de $($item.Text) ($($item.SubItems[1].Text))..."
                    $progressForm.Refresh()
            
                    $client = [PSCustomObject]@{
                        IP = $item.SubItems[1].Text
                        Nom = $item.Text
                        Utilisateur = $textUser.Text
                        MotDePasse = $textPSWD.Text
                        Port = $textPort.Text
                    }
            
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
            
                        # SFTP Connection + upload the update file
                        $progressDetails.AppendText("Upload du fichier de mise à jour...`r`n")
                        $progressForm.Refresh()
            
                        try {
                            # WinSCP upload command
                            & "C:\Program Files (x86)\WinSCP\WinSCP.com" /command `
                                "open scp://$WSCPLogin@$($client.IP)`:$($client.Port) -hostkey=`"*`"" `
                                "put `"$script:UpdateFilePath`" `"/usr/Firewall/Update/`"" `
                                "exit"
                            
                            $progressDetails.AppendText("Fichier uploadé avec succès`r`n")
                        } catch {
                            $progressDetails.AppendText("ERREUR lors de l'upload: $_`r`n")
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
                            
                            $progressDetails.AppendText("Récupération de la version actuelle...`r`n")
                            $preVersion = Invoke-SSHCommand -SSHSession $sessionssh -Command "getversion" -ErrorAction Stop
                            
                            $progressDetails.AppendText("Lancement de la mise à jour...`r`n")
                            $updateResult = Invoke-SSHCommand -SSHSession $sessionssh -Command "fwupdate -r -f /usr/Firewall/Update/$(Split-Path $script:UpdateFilePath -Leaf)" -ErrorAction Stop
                            
                            Remove-SSHSession -SSHSession $sessionssh | Out-Null
                            
                            $progressDetails.AppendText("Mise à jour lancée, attente de 10 minutes...`r`n")
                            $progressForm.Refresh()
                            
                            ## Sleep for 10 mins needed to let the update do its job
                            Start-Sleep -Seconds 600
                            
                            ## Verify update
                            $sessionssh = New-SSHSession @sessionParams -ErrorAction Stop
                            $postVersion = Invoke-SSHCommand -SSHSession $sessionssh -Command "getversion" -ErrorAction Stop
                            Remove-SSHSession -SSHSession $sessionssh | Out-Null
                            
                            $progressDetails.AppendText("Mise à jour terminée avec succès!`r`n")
                            $progressDetails.AppendText("Ancienne version: $($preVersion.Output)`r`n")
                            $progressDetails.AppendText("Nouvelle version: $($postVersion.Output)`r`n")
                            
                            $item.SubItems[3].Text = ($postVersion.Output | Where-Object { $_ -match "Version" } | Select-Object -First 1)
                            $item.SubItems[4].Text = "Mis à jour"
                            $item.BackColor = [System.Drawing.Color]::LightGreen
                            $successCount++
                        } catch {
                            $progressDetails.AppendText("ERREUR lors de la mise à jour: $_`r`n")
                            $item.SubItems[4].Text = "Échec mise à jour"
                            $item.BackColor = [System.Drawing.Color]::LightCoral
                            $errorCount++
                        }
                    } catch {
                        $progressDetails.AppendText("ERREUR lors du traitement: $_`r`n")
                        $item.SubItems[4].Text = "Erreur"
                        $item.BackColor = [System.Drawing.Color]::LightCoral
                        $errorCount++
                    }
                    
                    $listClients.Refresh()
                    $progressForm.Refresh()
                }
            
                # Fermer le formulaire de progression
                $progressForm.Close()
            
                # Afficher un récapitulatif
                $message = "Mises à jour terminées!`r`n"
                $message += "Clients traités: $($clientsToUpdate.Count)`r`n"
                $message += "Mises à jour réussies: $successCount`r`n"
                $message += "Échecs: $errorCount"
                
                [System.Windows.Forms.MessageBox]::Show($message, "Récapitulatif", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            })
            
            # Gestionnaire d'événements pour le bouton Fermer
            $buttonCancel.Add_Click({
                $form.Close()
            })
            
            # Charger automatiquement les clients au démarrage
            $initialClients = Load-ClientsFromCSV -filePath $defaultCsvPath
            foreach ($client in $initialClients) {
                $item = New-Object System.Windows.Forms.ListViewItem($client.Nom)
                $item.SubItems.Add($client.IP) | Out-Null
                $item.SubItems.Add($client.Modele) | Out-Null
                $item.SubItems.Add($client.Version) | Out-Null
                $item.SubItems.Add($client.Statut) | Out-Null
                
                # Colorer en fonction du statut
                if ($client.Statut -eq "OK") {
                    $item.BackColor = [System.Drawing.Color]::LightGreen
                } elseif ($client.Statut -eq "Erreur") {
                    $item.BackColor = [System.Drawing.Color]::LightCoral
                }
                
                $listClients.Items.Add($item)
            }
            
            # Afficher le formulaire
            $form.ShowDialog() | Out-Null
