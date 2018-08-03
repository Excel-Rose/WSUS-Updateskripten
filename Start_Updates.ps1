<#

Das Skript startet alle Rechner der CSV-Datei per WOL, startet die Aufgabeplanung zur Update-Suche und Installation. Das dazugehörige Skript wurde per GPO verteilt.
Dazu wurde eine GPO erstellt, die den Dienst startet und eine Firewall-Ausnahme hinzufügt, die den Remote Zugriff erlaubt.

Mithilfe des ForEach -parallel Switches wird der Prozess gleichzeitig auf vier Clients gestartet.

Nach dem ersten Durchgang wird der Status der Rechner mithilfe der WSUS berechnet und alle Rechner, die unter 100% liegen, erneut gestartet.
Nach Beenden des zweiten Durchgang erfolgt eine neuerliche Abfrage und die Logfiles werden per Telegram an den Admin versandt.

Die Dateipfade wurde entfernt und müssen vor der Verwendung natürlich angepasst werden!

#>

### Skriptweite ErrorAction

$ErrorActionPreference = 'SilentlyContinue'

# Das erste Skript untersucht das Verzeichnis der Logfiles auf Dateireste und löscht diese

Get-ChildItem "D:\irectory\of\the\LogFiles" | Remove-Item -Force

### Logfile anlegen

Write-Host "Lege Logfile für Updateprozess an." -ForegroundColor Green
start-transcript "P:\ath\to\Update-Log_EDV1.txt"

### Funktion zum Senden von Messages via Telegram Bot

Write-Host "Lade Zugangsdaten zu Telegram Bot für den Versand der Logfiles." -ForegroundColor Green
$BotKey = "Your BotKey here"
$ChatID = "You chat ID here"

###Tagesdatum als Variable speichern

Write-Host "Speichere verschiedene Systemwerte in Variablen." -ForegroundColor Green
$Tagesdatum = get-date -format d	

### CSV-Datei mit den NetBios-Namen und MAC-Adressen der zu startenden Rechner einlesen.

Write-Host "Lese CSV-Datei mit NetBios-Namen und MAC-Adressen ein und starte den Updateprozess der Clients." -ForegroundColor Green

Workflow Updateroutine_EDV_1 {

$csv = Import-Csv "P:\ath\to.csv"

foreach -parallel -throttlelimit 4 ($line in $csv) {

$NetBios = $line.NetBios
$MAC = $line.MAC   

# Test, ob der Rechner erreichbar ist und gegebenenfalls neustarten.

$erreichbar = Test-Connection -ComputerName $NetBios -Quiet -Count 1
if ($erreichbar -eq 'True') {
restart-computer -PSComputerName $NetBios -Force
sleep 120
}
 
# WOL-Paket schicken

Inlinescript {

    $MacByteArray = $Using:line.MAC -split "[:-]" | ForEach-Object { [Byte] "0x$_"}
    [Byte[]] $MagicPacket = (,0xFF * 6) + ($MacByteArray  * 16)
    $UdpClient = New-Object System.Net.Sockets.UdpClient
    $UdpClient.Connect(([System.Net.IPAddress]::Broadcast),7)
    $UdpClient.Send($MagicPacket,$MagicPacket.Length)
    $UdpClient.Close()

}

# Auf erfolgreichen Ping warten

Inlinescript {

$timer = [Diagnostics.Stopwatch]::StartNew()

# Loopen, bis der PC erreichbar ist.

while (-not (Test-Connection -ComputerName $Using:NetBios -Quiet -Count 1))
{  
    if ($timer.Elapsed.TotalSeconds -ge 300)
    {
       break
    }
        Start-Sleep -Seconds 15
}

# Timer nach den Beenden stoppen.

$timer.Stop()

}

# Wenn der Host erreichbar ist, Updates installieren, neustarten und WSUS-Report senden, ansonsten nächster Client.

if ((Test-Connection -ComputerName $NetBios -quiet) -eq 'Success') {

# Credentials erzeugen

# Read-Host -Prompt "Passwort fuer Benutzer administrator@XXX eingeben." -AsSecureString | ConvertFrom-SecureString | Out-File "\\Path\to\Admin_PW.txt"

$AdminName = "administrator@xxx" 
$Pass = Get-Content "\\Path\to\Admin_PW.txt" | ConvertTo-SecureString
$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass

sleep 120 # Batchskript wird beim Hochfahren parallel ausgeführt!

Inlinescript {Invoke-Command $Using:NetBios -ScriptBlock {schtasks /run /tn "WUpdate"} -Credential $Using:Cred}

sleep 180

# Per Log-File abfragen, ob der Prozess erfolgreich beendet wurde

$Logpath = "P:\ath\to\LogFiles\$NetBios.txt"

Inlinescript {

$timer = [Diagnostics.Stopwatch]::StartNew()

# Loopen, bis String mit Erfolgsmeldung im LogFile austaucht

while (-not (Select-String -Path $Using:Logpath -pattern "Ende der Windows PowerShell-Aufzeichnung")) # Pattern may differ depending on your OS language
{
    if ($timer.Elapsed.TotalSeconds -ge 10000)
    {
       break
    }
       Start-Sleep -Seconds 10
    }

$timer.Stop()

}

# Logfile löschen

Remove-Item $Logpath -Force

# Rechner neustarten

restart-computer -PSComputerName $NetBios -force
sleep 30

Inlinescript {

$timer = [Diagnostics.Stopwatch]::StartNew()

# Loopen, bis der PC erreichbar ist.

while (-not (Test-Connection -ComputerName $Using:NetBios -Quiet -Count 1))
{  
    if ($timer.Elapsed.TotalSeconds -ge 300)
    {
       break
    }
        Start-Sleep -Seconds 15
}

# Timer nach den Beenden stoppen.

$timer.Stop()

}

sleep 90

Inlinescript {Invoke-Command -ComputerName $Using:NetBios -Credential $Using:Cred -ScriptBlock {wuauclt.exe /reportnow}}

sleep 180

# Rechner herunterfahren.

stop-computer -PSComputerName $NetBios -force

}
}
}

# Workflow ausführen

Updateroutine_EDV_1

### Abfrage des UpdateStatus der WSUS Clients und erneuter Updateversuch

Write-Host "UpdateStatus von EDV-1 wird zum dritten Mal abgefragt und alle Rechner, deren Updatestatus unter 100% liegt, erneut gestartet." -ForeGroundColor Green

# Variable für Pfad des Logfiles anlegen

$LogFile = "P:\ath\to\LogFiles\Updatestatistik_EDV-1.csv"

# Test, ob Logfile bereits vorhanden ist und ggf. löschen

if (Test-Path -Path $LogFile) {Remove-Item $LogFile -Force}

# Verbindungsdaten zu WSUS API, Abfrage des Status der Rechner in EDV-1, Ausgabe in CSV-Datei

$Computername = 'Name of your WSUS Server'
$UseSSL = $False
$Port = 8530

[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
$WSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($Computername,$UseSSL,$Port)

$WOL_WSUS_Gruppe = $WSUS.GetComputerTargetGroups() | ? {$_.Name -eq "EDV-1"}
$Computerscope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope  ;
[void]$ComputerScope.ComputerTargetGroups.Add($WOL_WSUS_Gruppe)
$Updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope ;
$WSUS.GetSummariesPerComputerTarget($Updatescope,$ComputerScope) | Select-Object @{L='NetBios';E={($wsus.GetComputerTarget([guid]$_.ComputerTargetId)).FullDomainName}}, InstalledCount, DownloadedCount, FailedCount, UnKnownCount, NotInstalledCount, @{L= "InstalledOrNotApplicablePercentage";E={(($_.NotApplicableCount + $_.InstalledCount) / ($_.NotApplicableCount + $_.InstalledCount + $_.NotInstalledCount + $_.FailedCount + $_.UnknownCount))*100}} | Export-Csv -Path $LogFile -NoTypeInformation -Append -Encoding ASCII

### CSV-Datei mit nicht vollständig upgedateten Clients laden

$Updatestatistik = Import-CSV $LogFile -Delimiter ","

# CSV-Datei mit NetBios-Namen und MAC-Adressen alles Kabelclients laden

$Alle_Kabel_Clients = Import-CSV "P:\ath\to\CSV\containing\all\clients\Alle_Kabel_Macs.csv" -Delimiter "," 

# Füge der CSV eine Tabelle mit MAC hinzu

$Updatestatistik | ForEach {$_ | Add-Member 'MAC' $Null}

# CSV-Dateien vergleichen und neue CSV zur weiteren Verarbeitung erzeugen

ForEach ($Client in $Updatestatistik) {$Client.MAC = $Alle_Kabel_Clients | Where {$_.NetBios -eq $Client.NetBios} | Select -Expand 'MAC'}

$Updatestatistik | Export-Csv -Path $LogFile -NoTypeInformation -Encoding ASCII

# Worflow erneute Updates aller Clients, die unter 100% liegen

Workflow Erneute_Updateroutine_EDV_1 {

$csv = Import-Csv "P:\ath\to\Updatestatistik_EDV-1.csv"

foreach -parallel -throttlelimit 4 ($line in $csv) {

$NetBios = $line.NetBios
$MAC = $line.MAC   
$NetBios_ganz = $line.Netbios
$NetBios = $NetBios_ganz.Substring(0,$NetBios_ganz.Length - 12)

if ($($line.InstalledOrNotApplicablePercentage) -as [int] -lt 100) {

# Test, ob der Rechner erreichbar ist und gegebenenfalls neustarten.

$erreichbar = Test-Connection -ComputerName $NetBios -Quiet -Count 1
if ($erreichbar -eq 'True') {
restart-computer -PSComputerName $NetBios -Force
sleep 120
}
 
# WOL-Paket schicken

Inlinescript {

    $MacByteArray = $Using:line.MAC -split "[:-]" | ForEach-Object { [Byte] "0x$_"}
    [Byte[]] $MagicPacket = (,0xFF * 6) + ($MacByteArray  * 16)
    $UdpClient = New-Object System.Net.Sockets.UdpClient
    $UdpClient.Connect(([System.Net.IPAddress]::Broadcast),7)
    $UdpClient.Send($MagicPacket,$MagicPacket.Length)
    $UdpClient.Close()

}

# Auf erfolgreichen Ping warten

Inlinescript {

$timer = [Diagnostics.Stopwatch]::StartNew()

# Loopen, bis der PC erreichbar ist.

while (-not (Test-Connection -ComputerName $Using:NetBios -Quiet -Count 1))
{  
    if ($timer.Elapsed.TotalSeconds -ge 300)
    {
       break
    }
        Start-Sleep -Seconds 15
}

# Timer nach den Beenden stoppen.

$timer.Stop()

}

# Wenn der Host erreichbar ist, Updates installieren, neustarten und WSUS-Report senden, ansonsten nächster Client.

if ((Test-Connection -ComputerName $NetBios -quiet) -eq 'Success') {

# Credentials erzeugen

# Read-Host -Prompt "Passwort fuer Benutzer administrator@xxx eingeben." -AsSecureString | ConvertFrom-SecureString | Out-File "\\Path\to\Admin_PW.txt"

$AdminName = "administrator@xxx" 
$Pass = Get-Content "\\Path\to\Admin_PW.txt" | ConvertTo-SecureString
$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass

sleep 120 # Batchskript wird beim Hochfahren parallel ausgeführt!

Inlinescript {Invoke-Command $Using:NetBios -ScriptBlock {schtasks /run /tn "WUpdate"} -Credential $Using:Cred}

sleep 180

# Per Log-File abfragen, ob der Prozess erfolgreich beendet wurde

$Logpath = "P:\ath\to\LogFiles\$NetBios.txt"

Inlinescript {

$timer = [Diagnostics.Stopwatch]::StartNew()

# Loopen, bis String mit Erfolgsmeldung im LogFile austaucht

while (-not (Select-String -Path $Using:Logpath -pattern "Ende der Windows PowerShell-Aufzeichnung"))
{
    if ($timer.Elapsed.TotalSeconds -ge 6000)
    {
       break
    }
       Start-Sleep -Seconds 10
    }

$timer.Stop()

}

# Logfile löschen

Remove-Item $Logpath -Force

# Rechner neustarten

restart-computer -PSComputerName $NetBios -force
sleep 30

Inlinescript {

$timer = [Diagnostics.Stopwatch]::StartNew()

# Loopen, bis der PC erreichbar ist.

while (-not (Test-Connection -ComputerName $Using:NetBios -Quiet -Count 1))
{  
    if ($timer.Elapsed.TotalSeconds -ge 300)
    {
       break
    }
        Start-Sleep -Seconds 15
}

# Timer nach den Beenden stoppen.

$timer.Stop()

}

sleep 90

Inlinescript {Invoke-Command -ComputerName $Using:NetBios -Credential $Using:Cred -ScriptBlock {wuauclt.exe /reportnow}}

sleep 180

# Rechner herunterfahren.

stop-computer -PSComputerName $NetBios -force

}
}
}
}

Erneute_Updateroutine_EDV_1

### Erneutes WSUS Query nach nicht vollständig upgedateten Rechnern 

$WSUS.GetSummariesPerComputerTarget($Updatescope,$ComputerScope) | Select-Object @{L='NetBios';E={($wsus.GetComputerTarget([guid]$_.ComputerTargetId)).FullDomainName}}, InstalledCount, DownloadedCount, FailedCount, UnknownCount, NotInstalledCount, InstalledPendingRebootCount, @{L= "InstalledOrNotApplicablePercentage";E={(($_.NotApplicableCount + $_.InstalledCount) / ($_.NotApplicableCount + $_.InstalledCount + $_.NotInstalledCount + $_.FailedCount + $_.UnknownCount + $_.InstalledPendingRebootCount))*100}} | Export-Csv -Path "P:\ath\to\LogFiles\WUpdatestatistik_2_EDV-1.csv" -NoTypeInformation -Append -Encoding UTF8
$Failed_Updates = Import-CSV -Path "P:\ath\to\LogFiles\WUpdatestatistik_2_EDV-1.csv" -Delimiter "," | Where-Object {($_.InstalledOrNotApplicablePercentage) -as [int] -lt 100} | Group-Object -Property NetBios

# Liste der nicht erfolgreich upgedateten Rechner für den Versand mit Telegram

$Telegram_Liste = ($Failed_Updates).Name

if ($Telegram_Liste) {

# Logfile per Telegram senden löschen.

Write-Host "LogFiles werden per Telegram an den Administrator versandt." -ForeGroundColor Green

$Uhrzeit = Get-Date -Format T
$Nachricht_CleanUp = "Der Update-Vorgang vom $Tagesdatum wurde um $Uhrzeit erfolgreich beendet. Folgende Rechner konnten nicht vollstaendig gepatcht werden: 

$Telegram_Liste"

curl -s -X POST https://api.telegram.org/bot$BotKey/sendMessage -d chat_id=$ChatID -d text="$Nachricht_CleanUp"

# Logfile löschen.

Write-Host "Lösche das Logfile und beende das Skript." -ForegroundColor Green
remove-item "P:\ath\to\LogFiles\WUpdatestatistik_2_EDV-1.csv" -force

} 

ELSE {

$Erfolgsmeldung = "Alle Rechner wurden vollstaendig gepatcht."
curl -s -X POST https://api.telegram.org/bot$BotKey/sendMessage -d chat_id=$ChatID -d text="$Erfolgsmeldung"
Remove-item "P:\ath\to\LogFiles\WUpdatestatistik_2_EDV-1.csv"

}

# Aufzeichnung des Logfiles stoppen.

stop-transcript

Remove-Item "P:\ath\to\LogFiles\Update-Log_EDV1.txt" -Force
Remove-Item $Logfile -Force

Exit
