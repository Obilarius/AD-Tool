# Load ActiveDirectory module
Import-Module ActiveDirectory


 'Programm wird ausgeführt bitte warten bis Eingabe erscheint.'


# Die ersten beiden Befehle holen sich die .NET-Erweiterungen (sog. Assemblies) für die grafische Gestaltung in den RAM.
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 


# Die nächste Zeile erstellt aus der Formsbibliothek das Fensterobjekt.
$objForm = New-Object System.Windows.Forms.Form

# Hintergrundfarbe für das Fenster festlegen
$objForm.Backcolor=“white“

# Icon in die Titelleiste setzen
$objForm.Icon="img\adtool.ico"  #kann selbst definiert werden

# Hintergrundbild mit Formatierung Zentral = 2
$objForm.BackgroundImageLayout = 2
$objForm.BackgroundImage = [System.Drawing.Image]::FromFile('img\Arges_viereck.jpg')  #kann selbst definiert werden

# Position des Fensters festlegen
$objForm.StartPosition = "CenterScreen"

# Fenstergröße festlegen
#$objForm.Size = New-Object System.Drawing.Size(380,350)
$objForm.Size = New-Object System.Drawing.Size(550,350)

# Titelleiste festlegen
$objForm.Text = "AD Tool"



#############################################################################################################



# Vorhandene Kontakte auslesen

#$Initials = Get-ADUser -Filter * -Properties Initials | select * -ExpandProperty Initials
#$AccName = Get-ADUser -Filter * | select * -ExpandProperty SamAccountName


# KÜRZEL SUCHE
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(25,30)
$objLabel.Size = New-Object System.Drawing.Size(150,20)
$objLabel.BackColor = "Transparent"
$objLabel.Text = "KÜRZEL SUCHEN"
$objForm.Controls.Add($objLabel)

$objTextBox = New-Object System.Windows.Forms.TextBox
$objTextBox.Location = New-Object System.Drawing.Size(25,60)
$objTextBox.Size = New-Object System.Drawing.Size(150,20)
$objForm.Controls.Add($objTextBox)

$objLabel3 = New-Object System.Windows.Forms.Label
$objLabel3.Location = New-Object System.Drawing.Size(25,90)
$objLabel3.Size = New-Object System.Drawing.Size(150,100)
$objLabel3.BackColor = "Transparent"
$objForm.Controls.Add($objLabel3)

# NAMEN SUCHE
$objNLabel = New-Object System.Windows.Forms.Label
$objNLabel.Location = New-Object System.Drawing.Size(190,30)
$objNLabel.Size = New-Object System.Drawing.Size(150,20)
$objNLabel.BackColor = "Transparent"
$objNLabel.Text = "BENUTZERNAME SUCHEN"
$objForm.Controls.Add($objNLabel)

$objNTextBox = New-Object System.Windows.Forms.TextBox
$objNTextBox.Location = New-Object System.Drawing.Size(190,60)
$objNTextBox.Size = New-Object System.Drawing.Size(150,20)
$objForm.Controls.Add($objNTextBox)

$objNLabel3 = New-Object System.Windows.Forms.Label
$objNLabel3.Location = New-Object System.Drawing.Size(190,90)
$objNLabel3.Size = New-Object System.Drawing.Size(150,100)
$objNLabel3.BackColor = "Transparent"
$objForm.Controls.Add($objNLabel3)

# ANGEMELDET AN PC SUCHEN
$objPCLabel = New-Object System.Windows.Forms.Label
$objPCLabel.Location = New-Object System.Drawing.Size(355,30)
$objPCLabel.Size = New-Object System.Drawing.Size(150,20)
$objPCLabel.BackColor = "Transparent"
$objPCLabel.Text = "ANGEMELDET AN PC"
$objForm.Controls.Add($objPCLabel)

$objPCTextBox = New-Object System.Windows.Forms.TextBox
$objPCTextBox.Location = New-Object System.Drawing.Size(355,60)
$objPCTextBox.Size = New-Object System.Drawing.Size(150,20)
$objForm.Controls.Add($objPCTextBox)

$objPCLabel3 = New-Object System.Windows.Forms.Label
$objPCLabel3.Location = New-Object System.Drawing.Size(355,90)
$objPCLabel3.Size = New-Object System.Drawing.Size(150,100)
$objPCLabel3.BackColor = "Transparent"
$objForm.Controls.Add($objPCLabel3)


#############################################################################################################


    #OK Button anzeigen lassen
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(170,260)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "Überprüfen"
    $OKButton.Name = "Überprüfen"
    #$OKButton.DialogResult = "OK" # Ansonsten wird Fenster geschlossen
    $OKButton.Add_Click({

        # Suche Kürzel im AD
        if($objTextBox.Text) {
            $command = 'Get-ADUser -Filter "Initials -like ''' + $objTextBox.Text + '''" -Properties Initials'
            $User = Invoke-Expression $command
            if($User) {
                $objLabel3.Text = $User.Name + "`n" + $User.Initials + "`n" + $User.Enabled + "`n" + $User.ObjectGUID
            } else {
                $objLabel3.Text =  "Kürzel nicht vorhanden"
            }
        }
        if($objTextBox.Text -eq "") {
            $objLabel3.Text = ""
        }

        # Suche Name im AD
         if($objNTextBox.Text) {
             $command = 'Get-ADUser -Filter "SamAccountName -like ''' + $objNTextBox.Text + '''"'
             $User = Invoke-Expression $command
             if($User) {
                 $objNLabel3.Text = $User.Name + "`n" + $User.Initials + "`n" + $User.Enabled + "`n" + $User.ObjectGUID
             } else {
                 $objNLabel3.Text =  "Kürzel nicht vorhanden"
             }
         }
         if($objNTextBox.Text -eq "") {
             $objNLabel3.Text = ""
         }

         # Sucht User der an bestimmten PC angemeldet ist
         if($objPCTextBox.Text) {
             $objPCLabel3.Text = ""
             $command = 'Get-WmiObject -ComputerName ''' + $objPCTextBox.Text + ''' -Class Win32_ComputerSystem  | Select-Object username'
             $User = Invoke-Expression $command
             if($User) {
                 $objPCLabel3.Text = $User.username
             } else {
                 $objPCLabel3.Text =  "Kürzel nicht vorhanden"
             }
         }
         if($objPCTextBox.Text -eq "") {
             $objPCLabel3.Text = ""
         }


         $objForm.Refresh()
    })
    $objForm.Controls.Add($OKButton) 


    #Abbrechen Button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(265,260)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Beenden"
    $CancelButton.Name = "Beenden"
    $CancelButton.DialogResult = "Cancel"
    $CancelButton.Add_Click({$objForm.Close()})
    $objForm.Controls.Add($CancelButton) 


######################################################################################################

         
# Die letzte Zeile sorgt dafür, dass unser Fensterobjekt auf dem Bildschirm angezeigt wird.
[void] $objForm.ShowDialog()