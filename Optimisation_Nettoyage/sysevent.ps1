Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName PresentationFramework

# Bring a window to front
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@

function Show-InputBoxFront {
    param(
        [string]$Prompt,
        [string]$Title = "",
        [string]$Default = ""
    )

    # Bring current process window to front before showing InputBox
    $hwnd = (Get-Process -Id $PID).MainWindowHandle
    [Win32]::SetForegroundWindow($hwnd) | Out-Null

    return [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title, $Default)
}

function Get-MyApplicationEvent {
    param (
        [string]$LogName = 'System'
    )

    $days = $null

    do {
        $inputDays = Show-InputBoxFront "Nombre de jours à remonter ? (Laissez vide ou entrez 'q' pour annuler)" "Entrée requise"

        if ([string]::IsNullOrWhiteSpace($inputDays) -or $inputDays -match '^(q|quit)$') {
            [System.Windows.MessageBox]::Show("Opération annulée par l'utilisateur.", "Annulation", 'OK', 'Information') | Out-Null
            return
        }

        if ([int]::TryParse($inputDays, [ref]$days) -and $days -ge 1) {
            break
        }
        else {
            [System.Windows.MessageBox]::Show("Veuillez entrer un nombre entier supérieur ou égal à 1.", "Entrée invalide", 'OK', 'Warning') | Out-Null
        }
    } while ($true)

    $date = (Get-Date).AddDays(-$days)
    [System.Windows.MessageBox]::Show("Obtention des évènements systèmes depuis le $date ...", "Information", 'OK', 'Information') | Out-Null

    $EventLog = Get-WinEvent -FilterHashtable @{LogName=$LogName; StartTime=$date} |
        Where-Object { $_.Level -in 1, 2 } |
        Where-Object { $_.ProviderName -notin @(
            'Microsoft-Windows-DistributedCOM',
            'DCOM',
            'Microsoft-Windows-Dhcp-Client',
            'DHCP'
        )}

    if (-not $EventLog) {
        [System.Windows.MessageBox]::Show("Aucun évènement trouvé.", "Résultat", 'OK', 'Warning') | Out-Null
        return
    }

    $EntryCount = 0
    $SysEventLogEntries = @()

    foreach ($event in $EventLog) {
        $EntryCount++
        $SysEventLogEntries += [PSCustomObject]@{
            'Date et heure'        = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
            'ID'                   = $event.Id
            'ID name (Source)'     = $event.ProviderName
            'Niveau'               = $event.LevelDisplayName
            'Description'          = $event.Message
            '# Emplacement'        = $event.RecordId
        }

        Write-Progress -Activity "Évènements total trouvés: $($EventLog.Count)" `
                       -PercentComplete (($EntryCount / $EventLog.Count) * 100) `
                       -Status "Progrès: $EntryCount"
    }

    Write-Progress -Activity "Fini" -Completed

    $SysEventLogEntries | Out-GridView -Title "Observateur d'évènements" -Wait
}

# Appel
Get-MyApplicationEvent
