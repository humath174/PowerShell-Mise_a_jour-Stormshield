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

# Créer le formulaire principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Mise à jour de firmware Stormshield"
$form.Size = New-Object System.Drawing.Size(500,600) # Augmenté la hauteur pour le nouveau bouton
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

# Séparateur
$separator = New-Object System.Windows.Forms.Label
$separator.Location = New-Object System.Drawing.Point(10,$yPos)
$separator.Size = New-Object System.Drawing.Size(460,2)
$separator.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$form.Controls.Add($separator)
$yPos += 20

# Label pour afficher les infos détectées
$labelDetected = New-Object System.Windows.Forms.Label
$labelDetected.Location = New-Object System.Drawing.Point(10,$yPos)
$labelDetected.Size = New-Object System.Drawing.Size(460,40)
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

# Bouton Annuler
$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(260,$yPos)
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

# Afficher le formulaire et attendre la sélection
$result = $form.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Opération annulée par l'utilisateur."
    exit 1
}

# Récupérer les valeurs des champs
$IP = $textIP.Text
$Client = $textClient.Text
$User = $textUser.Text
$PSWD = $textPSWD.Text
$Port = $textPort.Text

# Validation des champs obligatoires
if ([string]::IsNullOrEmpty($IP) -or [string]::IsNullOrEmpty($Client) -or [string]::IsNullOrEmpty($User) -or [string]::IsNullOrEmpty($PSWD)) {
    [System.Windows.Forms.MessageBox]::Show("Tous les champs obligatoires doivent être remplis!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

# Vérifier si un fichier est nécessaire et sélectionné
if ($radioChoisir.Checked -and (-not $script:UpdateFilePath)) {
    [System.Windows.Forms.MessageBox]::Show("Aucun fichier de mise à jour sélectionné!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
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
        exit 1
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