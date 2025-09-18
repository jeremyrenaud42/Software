$global:xamlFileMain = "$global:appPathSource\InstallationMainWindow.xaml"
$global:xamlContentMain = Read-XamlFileContent $global:xamlFileMain
$global:formatedXamlFileMain = Format-XamlFile $global:xamlContentMain
$global:xamlDocMain = Convert-ToXmlDocument $global:formatedXamlFileMain
$global:XamlReaderMain = New-XamlReader $global:xamlDocMain
$global:windowMain = New-WPFWindowFromXaml $global:XamlReaderMain
$global:formControlsMain = Get-WPFControlsFromXaml $global:xamlDocMain $global:windowMain $sync

$global:formControlsMain.btnclose.Add_Click({
    $global:windowMain.Close()
    Exit
})
$global:formControlsMain.btnmin.Add_Click({
        $global:windowMain.WindowState = [System.Windows.WindowState]::Minimized
})
$global:formControlsMain.Titlebar.Add_MouseDown({
        $global:windowMain.DragMove()
})
 
$global:windowMain.add_Closed({
        Remove-Item -Path "$env:SystemDrive\_Tech\Applications\source\installation.lock" -Force 
        exit
})

$global:formControlsMain.richTxtBxOutput.add_Textchanged({
    $global:windowMain.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Background, [action]{}) #Refresh le Text
    $global:formControlsMain.richTxtBxOutput.ScrollToEnd() #scroll en bas
})

$appsInfo = Import-JsonFile "$global:appPathSource\InstallationApps.JSON"

function Update-InstallationStatus($softwareName) 
{
    if (Test-SoftwarePresence $appsInfo.$softwareName) 
    {
        Update-JsonFile "$global:appPathSource\InstallationApps.JSON" $softwareName "InstalledStatus" "1"
    }
}

function Install-SoftwareMenuApp($softwareName)
{
    if ($appNames -contains $softwareName) 
    {
        $appsInfo.$softwareName.path64 = $ExecutionContext.InvokeCommand.ExpandString($appsInfo.$softwareName.path64)
        $appsInfo.$softwareName.path32 = $ExecutionContext.InvokeCommand.ExpandString($appsInfo.$softwareName.path32)
        $appsInfo.$softwareName.pathAppData = $ExecutionContext.InvokeCommand.ExpandString($appsInfo.$softwareName.pathAppData)
        $appsInfo.$softwareName.RemoteName = $ExecutionContext.InvokeCommand.ExpandString($appsInfo.$softwareName.RemoteName)
    }
    Install-Software $appsInfo.$softwareName       
} 

function script:Update-CheckboxStatus
{
    $window.FindName("gridSettingsInstallationConfig").Children | 
    Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.Name -like "chkbox*" } | 
    ForEach-Object {
        $checkbox = $_
        $checkboxName = $checkbox.Name
        $status = if ($checkbox.IsChecked) { 1 } else { 0 }
        $global:jsonChkboxContent.$checkboxName.status = $status
    }

    $window.FindName("gridInstallationConfig").Children | 
    Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.Name -like "chkbox*" } | 
    ForEach-Object {
        $checkbox = $_
        $checkboxName = $checkbox.Name
        $status = if ($checkbox.IsChecked) { 1 } else { 0 }
        $global:jsonChkboxContent.$checkboxName.status = $status
    }

    $window.FindName("gridSettingsInstallationConfig").Children | 
    Where-Object { $_ -is [System.Windows.Controls.ComboBox] -and $_.Name -like "CbBox*" } | 
    ForEach-Object {
        $comboBox = $_
        $comboBoxName = $comboBox.Name
        $selectedValue = $comboBox.SelectedItem.Content

        if ($selectedValue) 
        {
            $global:jsonChkboxContent.$comboBoxName.Status = $selectedValue
        } 
    }

    $window.FindName("gridSettingsInstallationConfig").Children | 
    Where-Object { $_ -is [System.Windows.Controls.TextBox] -and $_.Name -like "TxtBx*" } | 
    ForEach-Object {
        $textBox = $_
        $textBoxName = $textBox.Name
        $textValue = $textBox.Text

        if ($textValue) {
            $global:jsonChkboxContent.$textBoxName.Status = $textValue
        }
    }
    Save-JsonFile $global:jsonChkboxContent $global:jsonSettingsFilePath 
}

function UpdateStepLabelForeGround
{
    foreach ($property in $global:jsonChkboxContent.PSObject.Properties) {
        $checkboxName = $property.Name
        $checkboxStatus = $property.Value.Status
        $labelName = "lbl_$checkboxName"
        $label = $global:windowMain.FindName($labelName)
    
        if ($null -ne $label) 
        {
            if ($checkboxStatus -eq "1") 
            {
                $label.Foreground = 'White'
            } 
        } 
    }
}

function Add-Text 
{
    param
    (
        [string]$Text,
        [string]$colorName,
        [switch]$SameLine = $false
    )
    # Create a Run element with the specified Text and color
    $run = New-Object System.Windows.Documents.Run
    $run.Text = $Text
    $run.Foreground = Get-BrushByName -brushName $colorName
    
    # Get the document and last paragraph
    $document = $global:sync["richTxtBxOutput"].Document
    $lastParagraph = $document.Blocks.LastBlock
    
    if ($lastParagraph -eq $null) 
    {
        # If no paragraph exists, create one
        $lastParagraph = New-Object System.Windows.Documents.Paragraph
        $document.Blocks.Add($lastParagraph)
    }
    
    # If not using SameLine, add a line break
    if (-not $SameLine) 
    {
        $lineBreak = New-Object System.Windows.Documents.LineBreak
        $lastParagraph.Inlines.Add($lineBreak)
    }
    
    # Add the colored run to the paragraph
    $lastParagraph.Inlines.Add($run)
    if ($formControls.rtbOutput_InstallationConfig -ne $null) 
    {
        $formControls.rtbOutput_InstallationConfig.AppendText("$Text`r")
        $formControls.rtbOutput_InstallationConfig.ScrollToEnd()
    }
}

function Get-Winget
{
    $global:formControlsMain.lblWinget.foreground = "DodgerBlue"
    $wingetStatus = Get-WingetStatus
    Add-Text -Text "Installation de Winget"
    if($wingetStatus -le '1.10')
    {
        Install-Winget
        $wingetStatus = Get-WingetStatus
        if($wingetStatus -ge '1.10')
        {
            Add-Log $global:logFileName " - Winget a été installé"
            Add-Text -Text " - Winget a été installé" -SameLine
            $global:formControlsMain.lblWinget.foreground = "MediumSeaGreen"
        }
        else 
        {
            Add-Log $global:logFileName " - Winget a échoué"
            Add-Text -Text " - Winget a échoué" -colorName "red" -SameLine
            $global:formControlsMain.lblWinget.foreground = "red"
        }
    }
    else 
    {
        Add-Log $global:logFileName " - Winget est déja installé"
        Add-Text -Text " - Winget est déja installé" -SameLine
        $global:formControlsMain.lblWinget.foreground = "MediumSeaGreen"
    }
}

function Get-Choco
{
    $global:formControlsMain.lblChoco.foreground = "DodgerBlue"
    $chocostatus = Get-ChocoStatus
    Add-Text -Text "Installation de Chocolatey"
    if($chocostatus -eq $false)
    {
        Install-Choco
        $chocostatus = Get-ChocoStatus
        if($chocostatus -eq $true)
        {
            Add-Log $global:logFileName " - Chocolatey a été installé"
            Add-Text -Text " - Chocolatey a été installé" -SameLine
            $global:formControlsMain.lblChoco.foreground = "MediumSeaGreen"
        }
        else 
        {
            Add-Log $global:logFileName " - Chocolatey a échoué"
            Add-Text -Text " - Chocolatey a échoué" -colorName "red" -SameLine
            $global:formControlsMain.lblChoco.foreground = "red"
        }
    }
    else 
    {
        Add-Log $global:logFileName " - Chocolatey  est déja installé"
        Add-Text -Text " - Chocolatey est déja installé" -SameLine 
        $global:formControlsMain.lblChoco.foreground = "MediumSeaGreen"
    }
}
function Get-Nuget
{
    $global:formControlsMain.lblNuget.foreground = "DodgerBlue"
    $nugetExist = Get-NugetStatus
    Add-Text -Text "Installation de NuGet"
    if($nugetExist -eq $false)
    {   
        Install-Nuget
        $nugetExist = Test-AppPresence "$env:SystemDrive\Program Files\WindowsPowerShell\Modules\NuGet" #permet de géré si lancé via autre user
        if($nugetExist -eq $true)
        {
            Add-Log $global:logFileName " - Nuget a été installé"
            Add-Text -Text " - Nuget a été installé" -SameLine
            $global:formControlsMain.lblNuget.foreground = "MediumSeaGreen"
        }
        else 
        {
            Add-Log $global:logFileName " - Nuget a échoué"
            Add-Text -Text " - Nuget a échoué" -colorName "red" -SameLine
            $global:formControlsMain.lblNuget.foreground = "red"
        }
    }
    else 
    {   
        Add-Log $global:logFileName " - Nuget est déja installé" 
        Add-Text -Text " - Nuget est déja installé" -SameLine
        $global:formControlsMain.lblNuget.foreground = "MediumSeaGreen"
    }
    Add-Text -Text "`n"
}

function Install-SoftwaresManager
{
    $global:formControlsMain.lblProgress.content = "Préparation"
    New-Item -Path "$applicationPath\source\Installation.lock" -ItemType 'File' -Force
    Add-Log $global:logFileName "Installation de $global:windowsVersion $global:OSUpdate le $global:actualDate"
    Clear-RichTextBox $global:sync["richTxtBxOutput"]
    Get-Winget
    Get-Choco
    Get-Nuget
}   

function Update-MsStore 
{
    $global:formControlsMain.lblProgress.Content = "Mises à jour du Microsoft Store"
    $global:formControlsMain.lbl_chkboxMSStore.foreground = "DodgerBlue"
    #pour gérer les vieilles versions qui ne vont pas s'updaté
    $storeVersion = (Get-AppxPackage Microsoft.WindowsStore).version
    if ($storeVersion -le 22110)
    {
        Add-Text -Text "Mettre le Store à jour manuellement"
        Start-Process "ms-windows-store:"
        return
    }
    Add-Text -Text "Lancement des updates du Microsoft Store"
    $namespaceName = "root\cimv2\mdm\dmmap"
    $className = "MDM_EnterpriseModernAppManagement_AppManagement01"
    $result = Get-CimInstance -Namespace $namespaceName -ClassName $className | Invoke-CimMethod -MethodName UpdateScanMethod
    if ($result.ReturnValue -eq 0) 
    {
        Add-Text -Text " - Mises à jour du Microsoft Store lancées" -SameLine
        $global:formControlsMain.lbl_chkboxMSStore.foreground = "MediumSeaGreen"
        Add-Log $global:logFileName "Mises à jour de Microsoft Store lancées" 
    } 
    else 
    {
        Add-Log $global:logFileName " - Échec des mises à jour du Microsoft Store" 
        Add-Text -Text " - Échec des mises à jour du Microsoft Store" -colorName "red" -SameLine
        $global:formControlsMain.lbl_chkboxMSStore.foreground = "red"
    }
}

Function Rename-SystemDrive
{
    <#
    .SYNOPSIS
        Renomme le lecteur C:
    .DESCRIPTION
        Renomme par OS par défaut au lieu du nom actuel (souvent disque local)
        Vérifie si ca a fonctionné
    .PARAMETER NewDiskName
        Le nouveau nom du disque
    .EXAMPLE
        Rename-SystemDrive -NewDiskName "OS"
        Renomme le disque OS
    #>
    
    
    [CmdletBinding()]
    param
    (
        [string]$NewDiskName = "OS"
    )
    $global:formControlsMain.lblProgress.Content = "Renommage du disque"
    $global:formControlsMain.lbl_chkboxDisque.foreground = "DodgerBlue"
    $systemDriverLetter = $env:SystemDrive.TrimEnd(':') #Retourne la lettre seulement sans le :
    $diskName = (Get-Volume -DriveLetter $systemDriverLetter).FileSystemLabel
    
    if($diskName -eq $NewDiskName)
    {
        $global:formControlsMain.lbl_chkboxDisque.foreground = "MediumSeaGreen"
        Add-Log $global:logFileName "Le disque est déja nommé $NewDiskName"
        Add-Text -Text "Le disque est déja nommé $NewDiskName"
    }
    else
    {
        Set-Volume -DriveLetter $systemDriverLetter -NewFileSystemLabel $NewDiskName
        $diskName = (Get-Volume -DriveLetter $systemDriverLetter).FileSystemLabel

        if($diskName -eq $NewDiskName)
        {
            $global:formControlsMain.lbl_chkboxDisque.foreground = "MediumSeaGreen"
            Add-Text -Text "Le disque $env:SystemDrive a été renommé $NewDiskName" 
            Add-Log $global:logFileName "Le disque $env:SystemDrive a été renommé $NewDiskName"
        }
        else
        {
            Add-Text -Text "Échec du renommage de disque" -colorName "red"
            Add-Log $global:logFileName "Échec du renommage de disque"
            $global:formControlsMain.lbl_chkboxDisque.foreground = "red"
        }
    } 
}

Function Set-ExplorerDisplay
{
    $global:formControlsMain.lblProgress.Content = "Configuration des paramètres de l'explorateur de fichiers"
    $global:formControlsMain.lbl_chkboxExplorer.foreground = "DodgerBlue"

    $explorerLaunchWindow = Get-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -EntryName 'LaunchTo'
    $providerNotifications = Get-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -EntryName 'ShowSyncProviderNotifications'

    if(($explorerLaunchWindow -eq '1') -and ($providerNotifications -eq '0'))
    {
        $global:formControlsMain.lbl_chkboxExplorer.foreground = "MediumSeaGreen" 
        Add-Text -Text "Ce PC remplace déja l'accès rapide"
        Add-Log $global:logFileName "Ce PC remplace déja l'accès rapide"
        Add-Text -Text "Le fournisseur de synchronisation est déjà décoché"
        Add-Log $global:logFileName "Le fournisseur de synchronisation est déjà décoché"
        return 
    }

    if (Set-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -EntryName 'LaunchTo' -Type 'DWord' -Value '1') 
    {
        Add-Log $global:logFileName "L'accès rapide a été remplacé par Ce PC"
        Add-Text -Text "L'accès rapide a été remplacé par Ce PC"
    }
    else 
    {
        Add-Text -Text "L'accès rapide n'a pas été remplacé par Ce PC" -colorName "red"
        Add-Log $global:logFileName "L'accès rapide n'a pas été remplacé par Ce PC"
    }

    if(Set-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -EntryName 'ShowSyncProviderNotifications' -Type 'DWord' -Value '0')
    {
        Add-Log $global:logFileName "Le fournisseur de synchronisation a été decoché"
        Add-Text -Text "Le fournisseur de synchronisation a été decoché" 
    }
    else
    {
        Add-Text -Text "Le fournisseur de synchronisation n'a pas été decoché" -colorName "red"
        Add-Log $global:logFileName "Le fournisseur de synchronisation n'a pas été decoché"
    }

    $explorerLaunchWindow = Get-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -EntryName 'LaunchTo'
    $providerNotifications = Get-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -EntryName 'ShowSyncProviderNotifications'
    if(($explorerLaunchWindow -eq '1') -and ($providerNotifications -eq '0'))
    {
        $global:formControlsMain.lbl_chkboxExplorer.foreground = "MediumSeaGreen"   
    }
    else 
    {
        $global:formControlsMain.lbl_chkboxExplorer.foreground = "red"     
    }
}

Function Disable-Bitlocker
{
    $global:formControlsMain.lblProgress.Content = "Désactivation du bitlocker"
    $global:formControlsMain.lbl_chkboxBitlocker.foreground = "DodgerBlue"
    $bitlockerStatus = Get-BitLockerVolume -MountPoint $env:SystemDrive | Select-Object -expand VolumeStatus
    if ($bitlockerStatus -eq 'FullyEncrypted' -or $bitlockerStatus -eq 'EncryptionInProgress')
    {
        manage-bde $env:systemdrive -off
        $bitlockerStatus = Get-BitLockerVolume -MountPoint $env:SystemDrive | Select-Object -expand VolumeStatus
        if ($bitlockerStatus -eq 'DecryptionInProgress')
        {
            $global:formControlsMain.lbl_chkboxBitlocker.foreground = "MediumSeaGreen"
            Add-Text -Text "Bitlocker a été désactivé"
            Add-Log $global:logFileName "Bitlocker a été désactivé"
        }
        else 
        {
            $global:formControlsMain.lbl_chkboxBitlocker.foreground = "red"
            Add-Text -Text "Bitlocker a échoué" -colorName "red"
            Add-Log $global:logFileName "Bitlocker a échoué"
        }
    }
    elseif ($bitlockerStatus -eq 'FullyDecrypted')
    {
        $global:formControlsMain.lbl_chkboxBitlocker.foreground = "MediumSeaGreen"
        Add-Text -Text "Bitlocker est déja désactivé"
        Add-Log $global:logFileName "Bitlocker est déja désactivé"
    }
    elseif ($bitlockerStatus -eq 'DecryptionInProgress')
    {
        $global:formControlsMain.lbl_chkboxBitlocker.foreground = "MediumSeaGreen"
        Add-Text -Text "Bitlocker est déja en cours de déchiffrement" 
        Add-Log $global:logFileName "Bitlocker est déja en cours de déchiffrement"
    }
    else 
    {
        $global:formControlsMain.lbl_chkboxBitlocker.foreground = "red"
        Add-Text -Text "Bitlocker a échoué" -colorName "red"
        Add-Log $global:logFileName "Bitlocker a échoué"
    }
}

Function Disable-FastBoot
{
    $global:formControlsMain.lblProgress.Content = "Desactivation du demarrage rapide"    
    $global:formControlsMain.lbl_chkboxStartup.foreground = "DodgerBlue"
    
    $power = Get-RegistryValue -KeyPath "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -EntryName 'HiberbootEnabled'
    if($power -eq '0')
    {  
        Add-Text -Text "Le démarrage rapide est déjà désactivé"
        Add-Log $global:logFileName "Le démarrage rapide est déjà désactivé"
        $global:formControlsMain.lbl_chkboxStartup.foreground = "MediumSeaGreen"
        return
    }
  
    if (Set-RegistryValue -KeyPath "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -EntryName 'HiberbootEnabled' -Type 'DWord' -Value '0')
    {
        Add-Text -Text "Le démarrage rapide a été désactivé"
        Add-Log $global:logFileName "Le démarrage rapide a été désactivé"
        $global:formControlsMain.lbl_chkboxStartup.foreground = "MediumSeaGreen"
    } 
    else  
    {
        Add-Text -Text "Le démarrage rapide n'a pas été désactivé" -colorName "red"
        Add-Log $global:logFileName "Le démarrage rapide n'a pas été désactivé"
        $global:formControlsMain.lbl_chkboxStartup.foreground = "red"
    }
}

Function Remove-EngKeyboard($selectedLanguage)
{
    $global:formControlsMain.lblProgress.Content = "Suppression du clavier $selectedLanguage"   
    $global:formControlsMain.lbl_chkboxClavier.foreground = "DodgerBlue"
    $langList = Get-WinUserLanguageList #Gets the language list for the current user account
    $filteredUserLangList = $langList | Where-Object LanguageTag -eq $selectedLanguage #sélectionne le clavier anglais canada de la liste
    if(($filteredUserLangList).LanguageTag -eq $selectedLanguage)
    {
        $langList.Remove($filteredUserLangList) #supprimer la clavier sélectionner
        Set-WinUserLanguageList $langList -Force #applique le changement
        $filteredUserLangList = $langList | Where-Object LanguageTag -eq $selectedLanguage #sélectionne le clavier anglais canada de la liste
        if(($filteredUserLangList).LanguageTag -eq $selectedLanguage)
        {
            Add-Text -Text "Le clavier $selectedLanguage n'a pas été supprimé" -colorName "red"
            Add-Log $global:logFileName "Le clavier $selectedLanguage n'a pas été supprimé"
            $global:formControlsMain.lbl_chkboxClavier.foreground = "red"
        }
        else
        {
            Add-Text -Text "Le clavier $selectedLanguage a été supprimé"
            $global:formControlsMain.lbl_chkboxClavier.foreground = "MediumSeaGreen"
            Add-Log $global:logFileName "Le clavier $selectedLanguage a été supprimé"
        }
    }
    else 
    {
        Add-Text -Text "Le clavier $selectedLanguage est déja supprimé"
        Add-Log $global:logFileName "Le clavier $selectedLanguage est déja supprimé"
        $global:formControlsMain.lbl_chkboxClavier.foreground = "MediumSeaGreen"
    }   
}

Function Set-Privacy
{
    $global:formControlsMain.lblProgress.Content = "Paramètres de confidentialité"
    $global:formControlsMain.lbl_chkboxConfi.foreground = "DodgerBlue"

    #Active ou désactive les contenus suggérés dans l'application Paramètres.
    $338393 = Get-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -EntryName 'SubscribedContent-338393Enabled'
    #Active ou désactive les contenus suggérés dans le menu Démarrer.
    $353694 = Get-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -EntryName 'SubscribedContent-353694Enabled'
    #Active ou désactive les contenus suggérés dans la barre des tâches.
    $353696 = Get-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -EntryName 'SubscribedContent-353696Enabled'
    #Active ou désactive le suivi des programmes récemment utilisés dans le menu Démarrer.
    $Start_TrackProgs = Get-RegistryValue -KeyPath "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -EntryName 'Start_TrackProgs'
    
    if (($338393 -eq 0) -and ($353694 -eq 0) -and ($353696 -eq 0) -and ($Start_TrackProgs -eq 0))
    {
        Add-Text -Text "Les options de confidentialité sont déjà configurées"
        Add-Log $global:logFileName "Les options de confidentialité sont déjà configurées"
        $global:formControlsMain.lbl_chkboxConfi.foreground = "MediumSeaGreen"
    }
    else 
    {
        $338393 = Set-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -EntryName 'SubscribedContent-338393Enabled' -Type 'DWord' -Value '0'
        $353694 = Set-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -EntryName 'SubscribedContent-353694Enabled' -Type 'DWord' -Value '0'
        $353696 = Set-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -EntryName 'SubscribedContent-353696Enabled' -Type 'DWord' -Value '0'
        $Start_TrackProgs = Set-RegistryValue -KeyPath "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -EntryName 'Start_TrackProgs' -Type 'DWord' -Value '0'
        
        if ($338393 -and $353694 -and $353696 -and $Start_TrackProgs)
        { 
            Add-Text -Text "Les options de confidentialité ont été configurées"
            Add-Log $global:logFileName "Les options de confidentialité ont été configurées"
            $global:formControlsMain.lbl_chkboxConfi.foreground = "MediumSeaGreen" 
        }
        else 
        {
            $global:formControlsMain.lbl_chkboxConfi.foreground = "red" 
            Add-Text -Text "Les options de confidentialité n'ont pas été configurées" -colorName "red"
            Add-Log $global:logFileName "Les options de confidentialité n'ont pas été configurées"
        } 
    }    
}

Function Enable-DesktopIcon
{
    $global:formControlsMain.lblProgress.Content = "Installation des icones systèmes sur le bureau"  
    $global:formControlsMain.lbl_chkboxIcone.foreground = "DodgerBlue"
    $keypath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
    New-RegistryKey $keypath

    $configPanel = Get-RegistryValue -KeyPath $keypath -EntryName '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}'
    $myPC = Get-RegistryValue -KeyPath $keypath -EntryName '{20D04FE0-3AEA-1069-A2D8-08002B30309D}'
    $userFolder = Get-RegistryValue -KeyPath $keypath -EntryName '{59031a47-3f72-44a7-89c5-5595fe6b30ee}'
    
    if (($configPanel -eq 0) -and ($myPC -eq 0) -and ($userFolder -eq 0))
    {
        Add-Text -Text "Les icones systèmes sont déjà installés sur le bureau"
        Add-Log $global:logFileName "Les icones systèmes sont déjà installés sur le bureau"
        $global:formControlsMain.lbl_chkboxIcone.foreground = "MediumSeaGreen"
    }
    else 
    {
        $configPanel = Set-RegistryValue -KeyPath $keypath -EntryName '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}' -Type 'DWord' -Value '0'
        $myPC = Set-RegistryValue -KeyPath $keypath -EntryName '{20D04FE0-3AEA-1069-A2D8-08002B30309D}' -Type 'DWord' -Value '0'
        $userFolder = Set-RegistryValue -KeyPath $keypath -EntryName '{59031a47-3f72-44a7-89c5-5595fe6b30ee}' -Type 'DWord' -Value '0'
        
        if ($configPanel -and $myPC -and $userFolder)
        {
            Add-Text -Text "Les icones systèmes ont été installés sur le bureau"
            Add-Log $global:logFileName "Les icones systèmes ont été installés sur le bureau"
            $global:formControlsMain.lbl_chkboxIcone.foreground = "MediumSeaGreen"  
        }
        else 
        {
            Add-Text -Text "Les icones systèmes n'ont pas été installés sur le bureau" -colorName "red"
            Add-Log $global:logFileName "Les icones systèmes n'ont pas été installés sur le bureau"
            $global:formControlsMain.lbl_chkboxIcone.foreground = "red"
        }
    }  
}

#Install les logiciels cochés
function Get-CheckBoxStatus 
{
    $Global:failStatus = $false
    $global:formControlsMain.lblProgress.Content = "Installation des logiciels"
    $global:formControlsMain.lblSoftware.foreground = "DodgerBlue"
    
    $checkboxes = $formControls.gridInstallationConfig.Children |
    Where-Object { 
        $_ -is [System.Windows.Controls.CheckBox] -and 
        $_.Name -like "chkbox*" 
    }
    $checkboxes | ForEach-Object {
        $checkbox = $_
        $checkboxName = $checkbox.Name
        
        $global:jsonChkboxContent = Import-JsonFile $global:jsonSettingsFilePath
        $checkboxStatus = $global:jsonChkboxContent.$checkboxName.status
    
        if ($checkboxStatus -eq 1) 
        {
            $softwareName = "$($checkbox.Content)"
            Install-Software $appsInfo.$softwareName
        } 
    }

    if($Global:failStatus -eq $true)
    {
        $global:formControlsMain.lblSoftware.foreground = "red"
    }
    elseif ($Global:failStatus -eq $false) 
    {
        $global:formControlsMain.lblSoftware.foreground = "MediumSeaGreen"
    }
    else
    {
        $global:formControlsMain.lblSoftware.foreground = "red"
    }
}

function Test-SoftwarePresence($appInfo)
{
   $softwareInstallationStatus= $false
   if (($appInfo.path64 -AND (Test-Path $appInfo.path64)) -OR 
   ($appInfo.path32 -AND (Test-Path $appInfo.path32)) -OR 
   ($appInfo.pathAppData -AND (Test-Path $appInfo.pathAppData)))
   {
     $softwareInstallationStatus = $true
   }
   return $softwareInstallationStatus
}

function Install-Software($appInfo)
{
    Add-Text -Text "Installation de $softwareName en cours"
    $window.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Background, [action]{}) #Refresh le Text
    Add-Log $global:logFileName "Installation de $softwareName"
    $softwareInstallationStatus = Test-SoftwarePresence $appInfo
        if($softwareInstallationStatus)
        {
            Add-Text -Text "- $softwareName est déja installé" -SameLine
            Add-Log $global:logFileName "- $softwareName est déja installé"
        }
        elseif($softwareInstallationStatus -eq $false)
        {  
            Install-SoftwareWithWinget $appInfo
        }
}

function Install-SoftwareWithWinget($appInfo)
{
    if($appInfo.WingetName)
    {
        winget install -e --id $appInfo.wingetname --accept-package-agreements --accept-source-agreements --silent
    } 
    $softwareInstallationStatus = Test-SoftwarePresence $appInfo
        if($softwareInstallationStatus)
        {
            Add-Text -Text " - $softwareName installé avec succès" -SameLine
            Add-Log $global:logFileName " - $softwareName installé avec succès"
            Update-InstallationStatus $softwareName
        } 
        else
        {
            Install-SoftwareWithChoco $appInfo
        }     
}

function Install-SoftwareWithChoco($appInfo)
{
    if($appInfo.ChocoName)
    {
        choco install $appInfo.ChocoName -y
    }
    $softwareInstallationStatus = Test-SoftwarePresence $appInfo
    if($softwareInstallationStatus)
    {     
        Add-Text -Text " - $softwareName installé avec succès" -SameLine
        Add-Log $global:logFileName " - $softwareName installé avec succès"
        Update-InstallationStatus $softwareName
    }
    else
    {
        Install-SoftwareWithNinite $appInfo
    } 
}

function Install-SoftwareWithNinite($appInfo)
{
    if($appInfo.RemoteName)
    {
        Invoke-WebRequest $appInfo.RemoteLink -OutFile $appInfo.RemoteName
        Start-Process $appInfo.RemoteName -Verb runAs -wait
    }
    $softwareInstallationStatus = Test-SoftwarePresence $appInfo
    if($softwareInstallationStatus)
    {     
        Add-Text -Text " - $softwareName installé avec succès" -SameLine
        Add-Log $global:logFileName " - $softwareName installé avec succès"
        Update-InstallationStatus $softwareName
    }
    else
    {
        Add-Text -Text " - $softwareName a échoué" -colorName "red" -SameLine
        Add-Log $global:logFileName " - $softwareName a échoué"
        $Global:failStatus = $true
    } 
}

function Get-ActivationStatus
{
    $global:formControlsMain.lblActivation.foreground = "DodgerBlue"
    $activated = Get-CIMInstance -query "select LicenseStatus from SoftwareLicensingProduct where LicenseStatus=1" | Select-Object -ExpandProperty LicenseStatus 
    Add-Text -Text "`n"
    if($activated -eq "1")
    {
        Add-Text -Text "$global:windowsVersion est activé sur cet ordinateur"
        Add-Log $global:logFileName "$global:windowsVersion est activé sur cet ordinateur"
        $global:formControlsMain.lblActivation.foreground = "MediumSeaGreen"      
    }
    else 
    {  
        Add-Text -Text "Windows n'est pas activé" -colorName "red"
        Add-Log $global:logFileName "Windows n'est pas activé"
        [System.Windows.MessageBox]::Show("Windows n'est pas activé","Installation Windows",0,64)   
        $global:formControlsMain.lblActivation.foreground = "red"
    }  
}

function Initialize-WindowsUpdate
{
    Install-Module PSWindowsUpdate -Force #install le module pour les Update de Windows
    $pathPSWindowsUpdateExist = test-path "$env:SystemDrive\Program Files\WindowsPowerShell\Modules\PSWindowsUpdate" 
    if($pathPSWindowsUpdateExist -eq $false) #si le module n'est pas là (Plan B)
    {
        choco install pswindowsupdate -y
    }
    Import-Module PSWindowsUpdate 
}

function Get-WindowsUpdateReboot
{
    $restartComputer = $false
    $rebootStatus = get-wurebootstatus -Silent #vérifie si ordi doit reboot à cause de windows update (PSwindowsupdate)
    if($rebootStatus)
    {
        Add-Text -Text "`n"
        Add-Text -Text "L'ordinateur devra redémarrer pour finaliser l'Installation des mises à jour"
        $messageBox = [System.Windows.MessageBox]::Show("L'ordinateur devra redémarrer pour finaliser l'Installation des mises à jour.`nVoulez-vous redémarrer ?","Installation Windows",4,64)
        if($messageBox -eq '6')
        {
            return $restartComputer = $true
        }
        else
        {
            return $restartComputer = $false
        }     
    } 
} 

function Install-WindowsUpdate
{
    <#
    .SYNOPSIS
        Installer les mises a jour de Windows
    .DESCRIPTION
        Liste les updates de Windows qui sont disponibles
        Si il ya 0 update dispo, va afficher que c'est deja bon
        si il y a des updates de trouvé, il va les faire une par une et afficher ce qu'il fait en temps reel
    .PARAMETER UpdateSize
        La Taille maximum des mises a jour de Windows
    .EXAMPLE
        Install-WindowsUpdate -UpdateSize "250" 
        Install chaque update qui est de moins de 250mb
    #>
    
    
    [CmdletBinding()]
    param
    (
        [int]$UpdateSize = 250
    )


    $global:formControlsMain.lblProgress.Content = "Mises à jour de Windows"
    $global:formControlsMain.lbl_chkboxWindowsUpdate.foreground = "DodgerBlue"
    Add-Text -Text "Vérification des mises à jour de Windows"
    Initialize-WindowsUpdate 
    $maxSizeBytes = $UpdateSize * 1MB #sans ca ca marchera pas
    $updates = Get-WUList -MaxSize $maxSizeBytes
    $totalUpdates = $updates.Count
        if($totalUpdates -eq 0)
        {
            Add-Text -Text " - Toutes les mises à jour sont deja installées" -SameLine 
            Add-Log $global:logFileName " - Toutes les mises à jour sont deja installées"
            $global:formControlsMain.lbl_chkboxWindowsUpdate.foreground = "MediumSeaGreen"   
        }
        elseif($totalUpdates -gt 0)
        {
            Add-Text -Text " - $totalUpdates mises à jour de disponibles" -SameLine 
            Add-Log $global:logFileName " - $totalUpdates mises à jour de disponibles"
            $currentUpdate = 0
                foreach($update in $updates)
                { 
                    $currentUpdate++ 
                    $kb = $update.KB
                    Add-Text -Text "Mise à jour $($currentUpdate) sur $($totalUpdates): $($update.Title)"
                    Add-Log $global:logFileName "Mise à jour $($currentUpdate) sur $($totalUpdates): $($update.Title)"
                    Get-WindowsUpdate -KBArticleID $kb -MaxSize $maxSizeBytes -Install -AcceptAll -IgnoreReboot     
                }
                $global:formControlsMain.lbl_chkboxWindowsUpdate.foreground = "MediumSeaGreen"
        }  
        else
        {
            Add-Text -Text " - Échec de la vérification des mise a jours de Windows" -colorName "red" -SameLine
            Add-Log $global:logFileName " - Échec de la vérification des mise a jours de Windows"
            $global:formControlsMain.lbl_chkboxWindowsUpdate.foreground = "red"
        } 
}

function Set-DefaultBrowser
{
    $currentHttpAssocation = Get-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\Shell\Associations\URLAssociations\http\UserChoice" -EntryName 'ProgId'    
    $currentHttpsAssocation = Get-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\Shell\Associations\URLAssociations\https\UserChoice" -EntryName 'ProgId'
    if(($currentHttpAssocation -notlike "ChromeHTML*") -and ($currentHttpsAssocation -notlike "ChromeHTML*"))
    {
        Start-Process ms-settings:defaultapps
        [System.Windows.MessageBox]::Show("Mettre Google Chrome par défaut","Installation Windows",0,64)   
    }
}
   
function Set-DefaultPDFViewer
{
    $currentDefaultPdfViewer = Get-RegistryValue -KeyPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\UserChoice" -EntryName 'ProgId'
    if($currentDefaultPdfViewer -notlike "*.Document.DC")
    {
        [System.Windows.MessageBox]::Show("Mettre Adobe Reader par défaut","Installation Windows",0,64)   
    }
}
    
function Set-GooglePinnedTaskbar
{
    $taskbardir = "$env:SystemDrive\Users\$env:username\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    $chromeTaskbarStatus= Test-Path "$taskbardir\*Google*Chrome*"
    if($chromeTaskbarStatus-eq $false)
    {
        [System.Windows.MessageBox]::Show("Épingler Google Chrome dans la barre des tâches","Installation Windows",0,64)   
    } 
}

function Complete-Installation
{
    $global:formControlsMain.lblManualComplete.foreground = "DodgerBlue"
    Add-Log $global:logFileName "Installation de Windows effectué avec Succès"
    Copy-Log $global:logFileName "$env:SystemDrive\Temp"
    Send-FTPLogs $global:appPathSource\$global:logFileName
    [Audio]::Volume = 0.25
    [console]::beep(1000,666)
    Start-Sleep -s 1
    [Audio]::Volume = 0.75
    Get-voice -Verb runAs
    Send-VoiceMessage "Vous avez terminer la configuration du Windows."
    Add-Text -Text "`n"
    Add-Text -Text "Vous avez terminer la configuration du Windows."
    net accounts /maxpwage:unlimited #mot de passe expire à 0
    $getNiniteProcess = get-process -name "ninite" -ErrorAction Ignore
    foreach ($Process in $getNiniteProcess)
    {
        Stop-Process $Process -Force -erroraction ignore
    }
    if ($global:jsonChkboxContent.chkboxGoogleChrome.status -eq 1)
    {
        Set-DefaultBrowser
        Set-GooglePinnedTaskbar
    }
    if ($global:jsonChkboxContent.chkboxAdobe.status -eq 1)
    {
        Set-DefaultPDFViewer
    }
    start-Process -FilePath "$global:appPathSource\caffeine64.exe" -ArgumentList "-appexit"
    $global:formControlsMain.lblManualComplete.foreground = "MediumSeaGreen"
    if ($global:jsonChkboxContent.chkboxWindowsUpdate.status -eq 1)
    {
        $wuRestart = Get-WindowsUpdateReboot
        if($wuRestart -eq $true)
        {
            $restartTime = $global:jsonChkboxContent.CbBoxRestartTimer.status
            shutdown /r /t $restartTime
        }  
    } 
    if ($global:jsonChkboxContent.chkboxRemove.status -eq 1)
    { 
        Invoke-Task -TaskName 'delete _tech' -ExecutedScript "$env:SystemDrive\Temp\Stoolbox\Remove.ps1"
    }
    $global:windowMain.Close()
    exit
}

function Main
{
    UpdateStepLabelForeGround
    Install-SoftwaresManager
    if ($global:jsonChkboxContent.chkboxMSStore.status -eq 1)
    { 
        Update-MsStore
    }
    if ($global:jsonChkboxContent.chkboxDisque.status -eq 1)
    { 
        Rename-SystemDrive -NewDiskName $global:jsonChkboxContent.TxtBxDiskName.status
    }
    if ($global:jsonChkboxContent.chkboxExplorer.status -eq 1)
    { 
        Set-ExplorerDisplay
    }
    if ($global:jsonChkboxContent.chkboxBitlocker.status -eq 1)
    { 
        Disable-Bitlocker
    }
    if ($global:jsonChkboxContent.chkboxStartup.status -eq 1)
    { 
        Disable-FastBoot
    }
    if ($global:jsonChkboxContent.chkboxClavier.status -eq 1)
    { 
        Remove-EngKeyboard 'en-CA'
    }
    if ($global:jsonChkboxContent.chkboxConfi.status -eq 1)
    { 
        Set-Privacy
    }
    if ($global:jsonChkboxContent.chkboxIcone.status -eq 1)
    { 
        Enable-DesktopIcon  
    }
    Add-Text -Text "`n"
    Get-CheckBoxStatus
    Get-ActivationStatus
    if ($global:jsonChkboxContent.chkboxWindowsUpdate.status -eq 1)
    { 
        Install-WindowsUpdate -UpdateSize $global:jsonChkboxContent.CbBoxSize.status
    }
    Complete-Installation
}