###############################################################################################
# filename          : printuserUI.ps1
# description       : installation Imprimantes en GUI
# Creation date     : 15/11/2023
# Update Date       : 27/12/2023
# version           : 2.1
# usage             : printuserUI.ps1
###############################################################################################
# 1     : Initial
# 2     : Ajout choix du driver si le nom de l'imprimante n'est pas dans la nomenclature groupe 
# 2.1   : Ajout fenetre installation en cours
# 2.2   : Modification texte
###############################################################################################

Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# Variables
$DriverNameCanon = 'Canon Generic Plus PCL6'
$DriverNameHP = 'HP Universal Printing PCL 6'
$DriverNameRicoh = 'PCL6 Driver for Universal Print'
$MsgInstal = 'installation en cours de '

# Main Form
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Installation Imprimantes'
$main_form.Width = 330
$main_form.Height = 180
$main_form.AutoSize = $true
$main_form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

# Main Form - Texte Ecran acceuil
$Label_main_form = New-Object System.Windows.Forms.Label
$Label_main_form.Text = "Entrer le nom de l'imprimante que vous voulez installer"
$Label_main_form.Location  = New-Object System.Drawing.Point(10,10)
$Label_main_form.AutoSize = $true
$main_form.Controls.Add($Label_main_form)

# Main Form -  TextBox pour saisie nom imprimante
$TextBox_NomImp = New-Object System.Windows.Forms.TextBox
$TextBox_NomImp.Width = 200
$TextBox_NomImp.Location  = New-Object System.Drawing.Point(40,50)
$main_form.Controls.Add($TextBox_NomImp)

# Main Form -  Bouton Validation pour installer 
$Button_Valide = New-Object System.Windows.Forms.Button
$Button_Valide.Location = New-Object System.Drawing.Point(170,100)
$Button_Valide.Size = New-Object System.Drawing.Size(60,20)
$Button_Valide.Text = "Installer"
$main_form.AcceptButton = $Button_Valide
$main_form.DialogResult = "OK"
$Button_Valide.Add_Click({

    if (($null -eq $TextBox_NomImp.Text) -or ("" -eq $TextBox_NomImp.Text)  ) { 
        [void] [System.Windows.MessageBox]::Show( "Veuillez saisir un nom d'imprimante", "Erreur", "OK", "Information" )
    }
    else {
        Instal_Imp -PrinterName $TextBox_NomImp.Text 
    }

     
})
$main_form.Controls.Add($Button_Valide)

# Main Form - # Bouton pour quitter
$Button_Exit = New-Object System.Windows.Forms.Button
$Button_Exit.Location = New-Object System.Drawing.Point(50,100)
$Button_Exit.Size = New-Object System.Drawing.Size(60,20)
$Button_Exit.Text = "Quitter"
$Button_Exit.DialogResult = "Cancel"
$main_form.CancelButton = $Button_Exit
$Button_Exit.Add_Click({ $Button_Exit.Text; $main_form.Close() })
$main_form.Controls.Add($Button_Exit)

# Form Choix Modele IMP
$Form_ChoixImp = New-Object System.Windows.Forms.Form
$Form_ChoixImp.Text = "Selectionner la marque de l'imprimante"
$Form_ChoixImp.Width = 400
$Form_ChoixImp.Height = 180
$Form_ChoixImp.AutoSize = $true
$Form_ChoixImp.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

# Form Choix Modele IMP - Texte Form Choix Modele IMP
$Label_Form_ChoixImp = New-Object System.Windows.Forms.Label
$Text_Label_Form_ChoixImp = "Selectionner la marque de l'imprimante" 
$Text_Label_Form_ChoixImp = $Text_Label_Form_ChoixImp + "`r`n"
$Text_Label_Form_ChoixImp = $Text_Label_Form_ChoixImp + "Si elle n'est pas présente, contacter la hotline"
$Label_Form_ChoixImp.Text = $Text_Label_Form_ChoixImp
$Label_Form_ChoixImp.Location  = New-Object System.Drawing.Point(10,10)
$Label_Form_ChoixImp.AutoSize = $true
$Form_ChoixImp.Controls.Add($Label_Form_ChoixImp)

# Form Choix Modele IMP - Radio Button HP
$CheckBox_HP = New-Object System.Windows.Forms.RadioButton
$CheckBox_HP.Text = " Imprimante HP"
$CheckBox_HP.Location = New-Object System.Drawing.Point(10,40)
$CheckBox_HP.Size = New-Object System.Drawing.Size(200,20)
$CheckBox_HP.Checked = $false
$Form_ChoixImp.Controls.Add($CheckBox_HP)

# Form Choix Modele IMP - Radio Button CANON
$CheckBox_CAN = New-Object System.Windows.Forms.RadioButton
$CheckBox_CAN.Text = " Imprimante CANON"
$CheckBox_CAN.Location = New-Object System.Drawing.Point(10,60)
$CheckBox_CAN.Size = New-Object System.Drawing.Size(200,20)
$CheckBox_CAN.Checked = $false
$Form_ChoixImp.Controls.Add($CheckBox_CAN)

# Form Choix Modele IMP - Radio Button RICOH
$CheckBox_RIC = New-Object System.Windows.Forms.RadioButton
$CheckBox_RIC.Text = " Imprimante RICOH"
$CheckBox_RIC.Location = New-Object System.Drawing.Point(10,80)
$CheckBox_RIC.Size = New-Object System.Drawing.Size(200,20)
$CheckBox_RIC.Checked = $false
$Form_ChoixImp.Controls.Add($CheckBox_RIC)

# Form Choix Modele IMP - Bouton Validation pour installer 
$Button_Valid_ChoixImp = New-Object System.Windows.Forms.Button
$Button_Valid_ChoixImp.Location = New-Object System.Drawing.Point(170,110)
$Button_Valid_ChoixImp.Size = New-Object System.Drawing.Size(60,20)
$Button_Valid_ChoixImp.Text = "Installer"
$Button_Valid_ChoixImp.DialogResult = "OK"
$Form_ChoixImp.AcceptButton = $Button_Valid_ChoixImp
$Button_Valid_ChoixImp.Add_Click({

    if ( ($CheckBox_CAN.Checked -eq $true) -or ($CheckBox_RIC.Checked -eq $true) -or ($CheckBox_HP.Checked -eq $true) )
    {
        $Form_ChoixImp.Close()
    }
    else 
    {
        [void] [System.Windows.MessageBox]::Show( "Veuiller Choisir la marque de l'imprimante", "Erreur", "OK", "Information" )
    } 
    
     
})
$Form_ChoixImp.Controls.Add($Button_Valid_ChoixImp)

# Form Choix Modele IMP - Bouton pour quitter
$Button_Exit_Form_ChoixImp = New-Object System.Windows.Forms.Button
$Button_Exit_Form_ChoixImp.Location = New-Object System.Drawing.Point(10,110)
$Button_Exit_Form_ChoixImp.Size = New-Object System.Drawing.Size(60,20)
$Button_Exit_Form_ChoixImp.Text = "Quitter"
$Button_Exit_Form_ChoixImp.DialogResult = "Cancel"
$Form_ChoixImp.CancelButton = $Button_Exit_Form_ChoixImp
$Button_Exit_Form_ChoixImp.Add_Click({ $Button_Exit_Form_ChoixImp.Text; $Form_ChoixImp.close })
$Form_ChoixImp.Controls.Add($Button_Exit_Form_ChoixImp)

# Form instalation en cours
$form_WIP = [Windows.forms.form]::new()
$form_WIP.Text = "Installation en cours"
$form_WIP.Size = [Drawing.Size]::new(300, 150)

# label_WIP instalation en cours
$label_WIP = [Windows.forms.Label]::new()
$label_WIP.Text = "Installation en cours "
$label_WIP.AutoSize = $true
$label_WIP.Location = [Drawing.Point]::new(10, 10)
$form_WIP.Controls.Add($label_WIP)

# Creation runspace pour form installation en cours 
$runspace_WIP = [runspacefactory]::CreateRunspace()
$runspace_WIP.Open()

# ScriptBlock pour runspace installation en cours 
$scriptBlock_WIP = {
    param($form_WIP)
    #[Windows.forms.Application]::Run($form_WIP)
    $form_WIP.ShowDialog()
}

# Creation de l'instance Powershell pour le runspace afin de lancer le ScriptBlock installation en cours 
$runspace_WIP.SessionStateProxy.SetVariable("form", $form_WIP)
$runspace_WIP.SessionStateProxy.SetVariable("scriptBlock", $scriptBlock_WIP)
$PsScript_WIP = [powershell]::Create().AddScript($scriptBlock_WIP).AddArgument($form_WIP)
$PsScript_WIP.Runspace = $runspace_WIP

function Instal_Imp {
    param(
        [Parameter()]
        [string] $PrinterName
    )
    # On suprimme tout ce qui n'est pas lettre, chiffre
    $PrinterName = $PrinterName -replace('[^a-zA-Z0-9]+','')
    # On met en majuscules 
    $PrinterName = $PrinterName.ToUpper()

    switch -Regex  ($PrinterName) {
        '^CAN' {
            $DriverName = $DriverNameCanon
        }
        '^HP' {
            $DriverName = $DriverNameHP
        }
        '^RIC' {
            $DriverName = $DriverNameRicoh
        }
        default{
            $Result_Form_ChoixImp = $Form_ChoixImp.ShowDialog()

            if ($CheckBox_CAN.Checked -eq $true){ $DriverName = $DriverNameCanon }
            if ($CheckBox_RIC.Checked -eq $true){ $DriverName = $DriverNameRicoh }
            if ($CheckBox_HP.Checked -eq $true){ $DriverName = $DriverNameHP }

            if ($Result_Form_ChoixImp -eq [System.Windows.Forms.DialogResult]::Cancel) 
            { 
                $Result_Form_ChoixImp
                return $false
            }
        }
    }
    
        $PrinterExist = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
        $DriverExist = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue

        If ( $PrinterExist )
        {
            [void] [System.Windows.MessageBox]::Show( "Imprimante déja existante", "Erreur", "OK", "Information" )
            return $false
        }
        else {
            If ( $DriverExist)
            {
                $PsScript_WIP.BeginInvoke()
                $label_WIP.Text = $MsgInstal
                $label_WIP.Text = $label_WIP.Text + $PrinterName + "."

                $Printerping = Test-Connection -ComputerName $PrinterName -Count 3 -Quiet

                if ($Printerping) {
                    try {
                        $label_WIP.Text = $label_WIP.Text + "."
                        Add-PrinterPort -Name $PrinterName -PrinterHostAddress "$($PrinterName).samse.fr" -ErrorAction stop
                    }
                    catch [Microsoft.Management.Infrastructure.CimException]{
                        if ($_.Exception.ErrorData.error_Code -eq "2147942583" ) {
                            #[void] [System.Windows.MessageBox]::Show( "DEBUG - Le port existe dÃ©ja", "Erreur", "OK", "Information" )
                        }                     
                        else {
                            [void] [System.Windows.MessageBox]::Show( "Erreur ajout port imprimante", "Erreur", "OK", "Information" )
                            $form_WIP.Close()
                            $PsScript_WIP.Stop()
                            return $false
                        }
                    }
                    catch {
                        [void] [System.Windows.MessageBox]::Show( "Erreur ajout port imprimante", "Erreur", "OK", "Information" )
                        $form_WIP.Close()
                        $PsScript_WIP.Stop()
                        return $false
                    }

                    try {
                        $label_WIP.Text = $label_WIP.Text + "."
                        Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PrinterName -ErrorAction stop
                   
                    }
                    catch {
                        [void] [System.Windows.MessageBox]::Show( "Erreur ajout imprimante", "Erreur", "OK", "Information" )
                        $form_WIP.Close()
                        $PsScript_WIP.Stop()
                        return $false
                    }
                    
                    $label_WIP.Text = $label_WIP.Text + "."
                    $form_WIP.Close()
                    $PsScript_WIP.Stop()

                    [void] [System.Windows.MessageBox]::Show( "Imprimante installée", "Succés", "OK", "Information" )
                    return $true
                } else {
                    [void] [System.Windows.MessageBox]::Show( "Imprimante Injoignable `r Verifier si le nom est correct et si celle-ci est bien allumée", "Erreur", "OK", "Information" )
                    $form_WIP.Close()
                    $PsScript_WIP.Stop()
                    return $false           
                }
            }
            else {
                [void] [System.Windows.MessageBox]::Show( "Driver $($DriverName) n'est pas installé, a installer depuis le centre logiciel", "Erreur", "OK", "Information" )
                return $false
            }

        }
}

$FinTraitement = $false

while ($FinTraitement -ne $true) {
    Write-Host $FinTraitement

    # On affiche la MainForm
    $Result_MainForm = ""
    $Result_MainForm  = $main_form.ShowDialog()
    

    if ($Result_MainForm -eq [System.Windows.Forms.DialogResult]::Cancel) 
    { 
        $main_form.Close()

        $PsScript_WIP.EndInvoke($Result_WIP)
        $PsScript_WIP.Dispose()
        $runspace_WIP.Close()
        $runspace_WIP.Dispose()

        exit
    }

}
