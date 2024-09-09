Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Functie om CSV-bestand te verwerken
function Process-CSVFile {
    param (
        [string]$filePath,
        [object]$config
    )

    if (Test-Path $filePath) {
        Write-Host "CSV-bestand gevonden: $filePath" -ForegroundColor Green
        # Importeren van CSV-gegevens en gebruikers toevoegen
        $users = Import-Csv -Path $filePath

        foreach ($user in $users) {
            $FullName = $user.FullName
            $Title = $user.Title

            # Controleer of alle vereiste velden aanwezig zijn
            if ($FullName -eq "" -or $Title -eq "") {
                Write-Host "Gebruiker met incomplete gegevens overgeslagen: $FullName" -ForegroundColor Yellow
                continue
            }

            # Voeg de gebruiker toe
            Add-ADUserFromData -FullName $FullName -Title $Title -config $config
        }
    } else {
        Write-Host "Het bestand bestaat niet of kan niet worden geopend." -ForegroundColor Red
    }
}

# Functie om een gebruiker toe te voegen aan AD
function Add-ADUserFromData {
    param (
        [string]$FullName,
        [string]$Title,
        [object]$config
    )

    # Haal voor- en achternaam op uit de volledige naam
    $NameParts = $FullName -split ' '
    $FirstName = $NameParts[0]
    $LastName = $NameParts[-1]

    # Stel de gebruikersnaam samen (bijv. voornaam.achternaam)
    $UserName = "$FirstName.$LastName"

    # Stel de UPN en andere AD-parameters in
    $UserPrincipalName = "$UserName@$($config.domain)"
    $SamAccountName = $UserName
    $OU = $config.ou_path

    try {
        New-ADUser `
        -Name $FullName `
        -GivenName $FirstName `
        -Surname $LastName `
        -UserPrincipalName $UserPrincipalName `
        -SamAccountName $SamAccountName `
        -Title $Title `
        -Path $OU `
        -Enabled $true `
        -AccountPassword (ConvertTo-SecureString $config.default_password -AsPlainText -Force) `
        -ChangePasswordAtLogon $config.require_password_change

        Write-Host "Gebruiker $FullName succesvol aangemaakt!" -ForegroundColor Green
    } catch {
        Write-Host "Er is een fout opgetreden bij het aanmaken van de gebruiker ${FullName}: $_" -ForegroundColor Red
    }

    # Functie om het configuratiebestand in te laden
    function Load-Config {
        param (
            [string]$configFilePath
        )

        if (Test-Path $configFilePath) {
            $configContent = Get-Content -Path $configFilePath | Out-String | ConvertFrom-Json
            return $configContent
        } else {
            Write-Host "Configuratiebestand niet gevonden. Script wordt afgebroken." -ForegroundColor Red
            exit
        }
    }

    # Laad het configuratiebestand
    $config = Load-Config -configFilePath ".\config.json"

    # Maak een nieuw venster aan
    $form = New-Object Windows.Forms.Form
    $form.Text = "Sleep het CSV-bestand hier"
    $form.Size = New-Object Drawing.Size(400, 200)
    $form.StartPosition = "CenterScreen"
    $form.AllowDrop = $true

    # Label toevoegen aan het formulier
    $label = New-Object Windows.Forms.Label
    $label.Text = "Sleep een CSV-bestand naar dit venster"
    $label.AutoSize = $true
    $label.Location = New-Object Drawing.Point(100, 80)
    $form.Controls.Add($label)

    # Functie voor het afhandelen van het slepen van bestanden
    $form.Add_DragEnter({
        param($sender, $e)
        if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
            $e.Effect = [Windows.Forms.DragDropEffects]::Copy
        }
    })

    # Functie voor het afhandelen van het neerzetten van het bestand
    $form.Add_DragDrop({
        param($sender, $e)
        $files = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
        $filePath = $files[0] # Neem het eerste bestand als er meerdere zijn
        Write-Host "Bestand gedropt: $filePath" -ForegroundColor Green

        # Controleer of het een CSV-bestand is
        if ($filePath -like "*.csv") {
            Process-CSVFile -filePath $filePath -config $config
        } else {
            Write-Host "Dit is geen CSV-bestand. Probeer opnieuw." -ForegroundColor Red
        }
    })

    # Toon het venster
    $form.Topmost = $true
    $form.ShowDialog()
