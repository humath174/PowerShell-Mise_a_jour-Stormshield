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

# Créer le formulaire principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Sélection du firmware Stormshield"
$form.Size = New-Object System.Drawing.Size(500,350)
$form.StartPosition = "CenterScreen"

# Label et ComboBox pour la catégorie
$labelCategorie = New-Object System.Windows.Forms.Label
$labelCategorie.Location = New-Object System.Drawing.Point(10,20)
$labelCategorie.Size = New-Object System.Drawing.Size(200,20)
$labelCategorie.Text = "Sélectionnez la catégorie:"
$form.Controls.Add($labelCategorie)

$comboCategorie = New-Object System.Windows.Forms.ComboBox
$comboCategorie.Location = New-Object System.Drawing.Point(220,20)
$comboCategorie.Size = New-Object System.Drawing.Size(250,20)
$comboCategorie.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$modeles.Keys | ForEach-Object { $comboCategorie.Items.Add($_) }
$form.Controls.Add($comboCategorie)

# Label et ComboBox pour le modèle spécifique
$labelModele = New-Object System.Windows.Forms.Label
$labelModele.Location = New-Object System.Drawing.Point(10,60)
$labelModele.Size = New-Object System.Drawing.Size(200,20)
$labelModele.Text = "Sélectionnez le modèle:"
$form.Controls.Add($labelModele)

$comboModele = New-Object System.Windows.Forms.ComboBox
$comboModele.Location = New-Object System.Drawing.Point(220,60)
$comboModele.Size = New-Object System.Drawing.Size(250,20)
$comboModele.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$form.Controls.Add($comboModele)

# Options de version
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Location = New-Object System.Drawing.Point(10,100)
$labelVersion.Size = New-Object System.Drawing.Size(200,20)
$labelVersion.Text = "Options de version:"
$form.Controls.Add($labelVersion)

$radioDerniere = New-Object System.Windows.Forms.RadioButton
$radioDerniere.Location = New-Object System.Drawing.Point(220,100)
$radioDerniere.Size = New-Object System.Drawing.Size(250,20)
$radioDerniere.Text = "Utiliser la dernière version"
$radioDerniere.Checked = $true
$form.Controls.Add($radioDerniere)

$radioChoisir = New-Object System.Windows.Forms.RadioButton
$radioChoisir.Location = New-Object System.Drawing.Point(220,130)
$radioChoisir.Size = New-Object System.Drawing.Size(250,20)
$radioChoisir.Text = "Choisir une version spécifique"
$form.Controls.Add($radioChoisir)

# Bouton pour sélectionner le fichier
$buttonSelect = New-Object System.Windows.Forms.Button
$buttonSelect.Location = New-Object System.Drawing.Point(220,170)
$buttonSelect.Size = New-Object System.Drawing.Size(250,30)
$buttonSelect.Text = "Sélectionner le fichier .maj"
$buttonSelect.Enabled = $false
$form.Controls.Add($buttonSelect)

# Bouton OK
$buttonOK = New-Object System.Windows.Forms.Button
$buttonOK.Location = New-Object System.Drawing.Point(150,220)
$buttonOK.Size = New-Object System.Drawing.Size(100,30)
$buttonOK.Text = "OK"
$buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $buttonOK
$form.Controls.Add($buttonOK)

# Bouton Annuler
$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(260,220)
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
        $buttonSelect.Text = "Fichier sélectionné"
    }
})

# Afficher le formulaire et attendre la sélection
$result = $form.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Opération annulée par l'utilisateur."
    exit 1
}

# Vérifier si un fichier est nécessaire et sélectionné
if ($radioChoisir.Checked -and (-not $script:UpdateFilePath)) {
    Write-Host "Aucun fichier sélectionné. Le script va s'arrêter."
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
        Write-Host "Aucun firmware trouvé dans $firmwareDir"
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
    
} catch {
    Add-Content -Path $logfile -Value "Error during SSH operations: $_"
    exit 1
}

Add-Content -Path $logfile -Value "------------------------------------------"
Add-Content -Path $logfile -Value "Script completed successfully at $(Get-Date)"
Add-Content -Path $logfile -Value "------------------------------------------"

Write-Host "Mise à jour terminée avec succès. Voir le log: $logfile"