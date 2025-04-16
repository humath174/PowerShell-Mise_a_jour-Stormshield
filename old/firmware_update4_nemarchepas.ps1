# Charger les assemblies pour l'interface graphique
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Configuration initiale
$scriptDir = $PSScriptRoot
$logDir = Join-Path $scriptDir "logs"

# Créer le dossier logs si nécessaire
if (-not (Test-Path -Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# Modèles Stormshield
$modeles = @{
    "VM" = @("EVA1", "EVA2", "EVA3", "EVA4", "EVAU", "VPAYG")
    "S" = @("SN160-A", "SN160W-A", "SN210-A", "SN210W-A", "SN310-A")
    "M" = @("SN510-A", "SN710-A", "SNi40-A", "SNi20-A", "SN-S-Series-220", "SN-S-Series-320", "SN-XS-Series-170", "SNi10")
    "L" = @("SN6100-A", "SN3100-A", "SN2100-A", "SN910-A", "SN1100-A", "SN6000-A", "SN3000-A", "SN2000-A", "SNxr1200-A", "SN-M-Series-720", "SN-M-Series-920", "SN520-A", "SN-L-Series-2200", "SN-L-Series-3200", "SN-XL-Series-5200", "SN-XL-Series-6200")
}

# Créer le formulaire principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "Gestionnaire Stormshield"
$form.Size = New-Object System.Drawing.Size(1200, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Panel pour le tableau
$tablePanel = New-Object System.Windows.Forms.Panel
$tablePanel.Location = New-Object System.Drawing.Point(10, 10)
$tablePanel.Size = New-Object System.Drawing.Size(1160, 600)
$tablePanel.AutoScroll = $true
$form.Controls.Add($tablePanel)

# En-têtes du tableau
$headers = @(
    @{Text="IP"; Width=120},
    @{Text="Utilisateur"; Width=100},
    @{Text="Mot de passe"; Width=100},
    @{Text="Version"; Width=150},
    @{Text="Modèle"; Width=120},
    @{Text="Taille"; Width=60},
    @{Text="Licence"; Width=80},
    @{Text="Actions"; Width=200}
)

# Créer les en-têtes
$yPos = 0
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, $yPos)
$headerPanel.Size = New-Object System.Drawing.Size(1140, 30)
$headerPanel.BackColor = [System.Drawing.Color]::LightGray
$tablePanel.Controls.Add($headerPanel)

$xPos = 0
foreach ($header in $headers) {
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point($xPos, 5)
    $label.Size = New-Object System.Drawing.Size($header.Width, 20)
    $label.Text = $header.Text
    $label.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [System.Drawing.FontStyle]::Bold)
    $headerPanel.Controls.Add($label)
    $xPos += $header.Width
}

$yPos += 35

# Liste pour stocker les lignes d'appareils
$appareilsRows = @()

# Fonction pour ajouter une ligne d'appareil
function Add-AppareilRow {
    $rowPanel = New-Object System.Windows.Forms.Panel
    $rowPanel.Location = New-Object System.Drawing.Point(0, $yPos)
    $rowPanel.Size = New-Object System.Drawing.Size(1140, 40)
    $rowPanel.BackColor = if ($appareilsRows.Count % 2 -eq 0) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::WhiteSmoke }
    $tablePanel.Controls.Add($rowPanel)
    
    $controls = @{}

    # Champ IP
    $textIP = New-Object System.Windows.Forms.TextBox
    $textIP.Location = New-Object System.Drawing.Point(0, 10)
    $textIP.Size = New-Object System.Drawing.Size(120, 20)
    $rowPanel.Controls.Add($textIP)

    # Champ Utilisateur
    $textUser = New-Object System.Windows.Forms.TextBox
    $textUser.Location = New-Object System.Drawing.Point(120, 10)
    $textUser.Size = New-Object System.Drawing.Size(100, 20)
    $rowPanel.Controls.Add($textUser)

    # Champ Mot de passe
    $textPSWD = New-Object System.Windows.Forms.TextBox
    $textPSWD.Location = New-Object System.Drawing.Point(220, 10)
    $textPSWD.Size = New-Object System.Drawing.Size(100, 20)
    $textPSWD.PasswordChar = '*'
    $rowPanel.Controls.Add($textPSWD)

    # Champ Port
    $textPort = New-Object System.Windows.Forms.TextBox
    $textPort.Location = New-Object System.Drawing.Point(320, 10)
    $textPort.Size = New-Object System.Drawing.Size(60, 20)
    $textPort.Text = "13422"  # Valeur par défaut
    $rowPanel.Controls.Add($textPort)

    # Champ Client
    $textClient = New-Object System.Windows.Forms.TextBox
    $textClient.Location = New-Object System.Drawing.Point(380, 10)
    $textClient.Size = New-Object System.Drawing.Size(100, 20)
    $rowPanel.Controls.Add($textClient)

    # Label Version
    $labelVersion = New-Object System.Windows.Forms.Label
    $labelVersion.Location = New-Object System.Drawing.Point(480, 10)
    $labelVersion.Size = New-Object System.Drawing.Size(150, 20)
    $labelVersion.Text = ""
    $rowPanel.Controls.Add($labelVersion)

    # Label Modèle
    $labelModele = New-Object System.Windows.Forms.Label
    $labelModele.Location = New-Object System.Drawing.Point(630, 10)
    $labelModele.Size = New-Object System.Drawing.Size(120, 20)
    $labelModele.Text = ""
    $rowPanel.Controls.Add($labelModele)

    # Label Taille
    $labelTaille = New-Object System.Windows.Forms.Label
    $labelTaille.Location = New-Object System.Drawing.Point(750, 10)
    $labelTaille.Size = New-Object System.Drawing.Size(60, 20)
    $labelTaille.Text = ""
    $rowPanel.Controls.Add($labelTaille)

    # Label Licence
    $labelLicence = New-Object System.Windows.Forms.Label
    $labelLicence.Location = New-Object System.Drawing.Point(810, 10)
    $labelLicence.Size = New-Object System.Drawing.Size(80, 20)
    $labelLicence.Text = ""
    $rowPanel.Controls.Add($labelLicence)

    # Bouton Get Info
    $buttonGetInfo = New-Object System.Windows.Forms.Button
    $buttonGetInfo.Location = New-Object System.Drawing.Point(890, 8)
    $buttonGetInfo.Size = New-Object System.Drawing.Size(80, 25)
    $buttonGetInfo.Text = "Infos"
    $rowPanel.Controls.Add($buttonGetInfo)

    # Bouton Mise à jour
    $buttonUpdate = New-Object System.Windows.Forms.Button
    $buttonUpdate.Location = New-Object System.Drawing.Point(980, 8)
    $buttonUpdate.Size = New-Object System.Drawing.Size(80, 25)
    $buttonUpdate.Text = "Màj"
    $buttonUpdate.Enabled = $false
    $rowPanel.Controls.Add($buttonUpdate)

    # Bouton Supprimer
    $buttonRemove = New-Object System.Windows.Forms.Button
    $buttonRemove.Location = New-Object System.Drawing.Point(1070, 8)
    $buttonRemove.Size = New-Object System.Drawing.Size(80, 25)
    $buttonRemove.Text = "Supprimer"
    $buttonRemove.ForeColor = [System.Drawing.Color]::Red
    $rowPanel.Controls.Add($buttonRemove)

    # Stocker les contrôles
    $controls["RowPanel"] = $rowPanel
    $controls["IP"] = $textIP
    $controls["User"] = $textUser
    $controls["PSWD"] = $textPSWD
    $controls["Port"] = $textPort
    $controls["Client"] = $textClient
    $controls["Version"] = $labelVersion
    $controls["Modele"] = $labelModele
    $controls["Taille"] = $labelTaille
    $controls["Licence"] = $labelLicence
    $controls["ButtonGetInfo"] = $buttonGetInfo
    $controls["ButtonUpdate"] = $buttonUpdate
    $controls["ButtonRemove"] = $buttonRemove

    # Ajouter à la liste des lignes
    $appareilsRows += $controls

    # Mettre à jour la position
    $yPos += 45
    Update-RowsPosition
}
    
    # Gestionnaire d'événements pour le bouton Get Info
    $buttonGetInfo.Add_Click({
        $ip = $controls["IP"].Text.Trim()
        $user = $controls["User"].Text.Trim()
        $pswd = $controls["PSWD"].Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($ip) -or [string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pswd)) {
            [System.Windows.Forms.MessageBox]::Show("Veuillez remplir tous les champs de connexion!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        try {
            # Convertir le mot de passe et créer les credentials
            $securePswd = ConvertTo-SecureString -String $pswd -AsPlainText -Force
            $credential = New-Object -TypeName System.Management.Automation.PSCredential ($user, $securePswd)
            
            # Connexion SSH
            $sessionParams = @{
                ComputerName = $ip
                Credential   = $credential
                AcceptKey    = $true
                Port         = 13422
            }
            
            $session = New-SSHSession @sessionParams -ErrorAction Stop
            
            # Récupérer les informations
            $versionInfo = Invoke-SSHCommand -SSHSession $session -Command "getversion" -ErrorAction Stop
            $controls["Version"].Text = ($versionInfo.Output -join " ") -replace "`n", " " -replace "`r", ""
            
            # Détection du modèle et de la taille
            $modelInfo = Invoke-SSHCommand -SSHSession $session -Command "getsystem" -ErrorAction Stop
            $modelOutput = $modelInfo.Output -join " "
            
            $modelDetected = $null
            $tailleDetected = $null
            
            foreach ($category in $modeles.Keys) {
                foreach ($model in $modeles[$category]) {
                    if ($modelOutput -match $model) {
                        $modelDetected = $model
                        $tailleDetected = $category
                        break
                    }
                }
                if ($modelDetected) { break }
            }
            
            if ($modelDetected) {
                $controls["Modele"].Text = $modelDetected
                $controls["Taille"].Text = $tailleDetected
                
                # Vérification de la licence
                $licenceInfo = Invoke-SSHCommand -SSHSession $session -Command "getlicence" -ErrorAction Stop
                if ($licenceInfo.Output -match "Maintenance: Valid") {
                    $controls["Licence"].Text = "Valide"
                    $controls["Licence"].ForeColor = [System.Drawing.Color]::Green
                    $controls["ButtonUpdate"].Enabled = $true
                } else {
                    $controls["Licence"].Text = "Invalide"
                    $controls["Licence"].ForeColor = [System.Drawing.Color]::Red
                    $controls["ButtonUpdate"].Enabled = $false
                }
            } else {
                $controls["Modele"].Text = "Inconnu"
                $controls["Taille"].Text = "Inconnue"
                $controls["Licence"].Text = "Inconnue"
                $controls["ButtonUpdate"].Enabled = $false
            }
            
            Remove-SSHSession -SSHSession $session | Out-Null
            
        } catch {
            $controls["Version"].Text = "Erreur"
            $controls["Modele"].Text = ""
            $controls["Taille"].Text = ""
            $controls["Licence"].Text = ""
            $controls["ButtonUpdate"].Enabled = $false
            [System.Windows.Forms.MessageBox]::Show("Erreur de connexion: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    
    # Gestionnaire d'événements pour le bouton Mise à jour
    $buttonUpdate.Add_Click({
        $ip = $controls["IP"].Text.Trim()
        $user = $controls["User"].Text.Trim()
        $pswd = $controls["PSWD"].Text.Trim()
        $taille = $controls["Taille"].Text
        
        if ([string]::IsNullOrWhiteSpace($ip) -or [string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pswd)) {
            [System.Windows.Forms.MessageBox]::Show("Veuillez remplir tous les champs de connexion!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        if ($controls["Licence"].Text -ne "Valide") {
            [System.Windows.Forms.MessageBox]::Show("La licence de maintenance n'est pas valide!", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        # Déterminer le sous-dossier en fonction de la taille
        $sousDossier = $taille
        $firmwareDir = Join-Path $scriptDir "version\$sousDossier"
        
        # Trouver le firmware le plus récent
        $latestFirmware = Get-ChildItem -Path $firmwareDir -Filter "*.maj" | 
                         Sort-Object LastWriteTime -Descending | 
                         Select-Object -First 1
        
        if (-not $latestFirmware) {
            [System.Windows.Forms.MessageBox]::Show("Aucun firmware trouvé dans le dossier $firmwareDir", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        
        $updateFilePath = $latestFirmware.FullName
        
        try {
            # Créer le fichier log
            $date = Get-Date -Format "dd-MM-yyyy-HH-mm"
            $logfile = Join-Path $logDir "$ip-$date.log"
            
            Add-Content -Path $logfile -Value "------------------------------------------"
            Add-Content -Path $logfile -Value "Mise à jour démarrée à $(Get-Date)"
            Add-Content -Path $logfile -Value "------------------------------------------"
            Add-Content -Path $logfile -Value "IP: $ip"
            Add-Content -Path $logfile -Value "Utilisateur: $user"
            Add-Content -Path $logfile -Value "Modèle: $($controls["Modele"].Text)"
            Add-Content -Path $logfile -Value "Taille: $taille"
            Add-Content -Path $logfile -Value "Firmware utilisé: $(Split-Path $updateFilePath -Leaf)"
            
            # WinSCP upload
            $wscpLogin = "$user" + ":" + "$pswd"
            & "C:\Program Files (x86)\WinSCP\WinSCP.com" /command `
                "open scp://$wscpLogin@$ip`:13422 -hostkey=`"*`"" `
                "put `"$updateFilePath`" `"/usr/Firewall/Update/`"" `
                "exit"
            
            Add-Content -Path $logfile -Value "Fichier uploadé avec succès"
            
            # SSH update
            $securePswd = ConvertTo-SecureString -String $pswd -AsPlainText -Force
            $credential = New-Object -TypeName System.Management.Automation.PSCredential ($user, $securePswd)
            
            $sessionParams = @{
                ComputerName = $ip
                Credential   = $credential
                AcceptKey    = $true
                Port         = 13422
            }
            
            $session = New-SSHSession @sessionParams -ErrorAction Stop
            
            $updateResult = Invoke-SSHCommand -SSHSession $session -Command "fwupdate -r -f /usr/Firewall/Update/$(Split-Path $updateFilePath -Leaf)" -ErrorAction Stop
            Add-Content -Path $logfile -Value ($updateResult.Output -join "`r`n")
            
            Remove-SSHSession -SSHSession $session | Out-Null
            
            Add-Content -Path $logfile -Value "Commande de mise à jour envoyée. Attente de 10 minutes..."
            
            # Attendre la mise à jour
            Start-Sleep -Seconds 600
            
            # Vérifier la version après mise à jour
            $session = New-SSHSession @sessionParams -ErrorAction Stop
            $postVersion = Invoke-SSHCommand -SSHSession $session -Command "getversion" -ErrorAction Stop
            Add-Content -Path $logfile -Value "------------------------------------------"
            Add-Content -Path $logfile -Value "Version après mise à jour:"
            Add-Content -Path $logfile -Value ($postVersion.Output -join "`r`n")
            
            Remove-SSHSession -SSHSession $session | Out-Null
            
            # Mettre à jour l'affichage
            $controls["Version"].Text = ($postVersion.Output -join " ") -replace "`n", " " -replace "`r", ""
            
            [System.Windows.Forms.MessageBox]::Show("Mise à jour terminée avec succès!`nVoir le log: $logfile", "Succès", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
        } catch {
            Add-Content -Path $logfile -Value "Erreur: $_"
            [System.Windows.Forms.MessageBox]::Show("Erreur lors de la mise à jour: $_", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    
    # Gestionnaire d'événements pour le bouton Supprimer
    $buttonRemove.Add_Click({
        $tablePanel.Controls.Remove($rowPanel)
        $appareilsRows = $appareilsRows | Where-Object { $_ -ne $controls }
        Update-RowsPosition
    })
    
    # Ajouter à la liste des lignes
    $appareilsRows += $controls
    
    # Mettre à jour la position
    $yPos += 45
    Update-RowsPosition

# Fonction pour mettre à jour les positions des lignes
function Update-RowsPosition {
    $y = 35 # Après les en-têtes
    foreach ($row in $appareilsRows) {
        $row["RowPanel"].Location = New-Object System.Drawing.Point(0, $y)
        $y += 45
    }
    $tablePanel.Height = [Math]::Max(600, $y + 10)
}

# Bouton Ajouter en bas du formulaire
$buttonAdd = New-Object System.Windows.Forms.Button
$buttonAdd.Location = New-Object System.Drawing.Point(10, 620)
$buttonAdd.Size = New-Object System.Drawing.Size(100, 30)
$buttonAdd.Text = "Ajouter"
$buttonAdd.Add_Click({ Add-AppareilRow })
$form.Controls.Add($buttonAdd)

# Bouton Quitter
$buttonQuit = New-Object System.Windows.Forms.Button
$buttonQuit.Location = New-Object System.Drawing.Point(120, 620)
$buttonQuit.Size = New-Object System.Drawing.Size(100, 30)
$buttonQuit.Text = "Quitter"
$buttonQuit.Add_Click({ $form.Close() })
$form.Controls.Add($buttonQuit)

# Ajouter une ligne vide au démarrage
Add-AppareilRow

# Afficher le formulaire
$form.ShowDialog() | Out-Null