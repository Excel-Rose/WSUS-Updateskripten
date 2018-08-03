<# 

Das Skript erledigt lokal den Updatevorgang per WSUS und diverse CleanUps.

- WSUS Query, Download und Installation
- Errorhandling des Clients
- Reporting per Telegram
- Logfile wird zentral auf Share gespeichert und per Telegram an Admin gesandt

#>

### Hostname auslesen und als Variable speichern

$Client = (Get-Childitem env:computername).Value

### Logfile auf Share anlegen

Start-Transcript -Path "\\Path\to\LogFiles\$Client.txt"

### Skriptweite Erroraction festlegen

$ErrorActionPreference = "silentlycontinue"

### CleanUp

Write-Host "Führe verschiedene Cleanup-Aufgaben durch." -ForegroundColor Green

# Dateien im Download-Folder löschen, die älter als zwei Tage sind

$limit = (Get-Date).AddDays(-2)
$Downloadfolder = "C:\Windows\SoftwareDistribution\Download"

Get-ChildItem -Path $Downloadfolder -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force
Get-ChildItem -Path $Downloadfolder -Recurse -Force | Where-Object { $_.PSIsContainer -and (Get-ChildItem -Path $_.FullName -Recurse -Force | Where-Object { !$_.PSIsContainer }) -eq $null } | Remove-Item -Force -Recurse

# Cleanmanager bereinigt C:, wenn Speicherplatz unter 2GB

if (((Get-WMIObject Win32_LogicalDisk -filter "name='c:'").freespace / 1GB) -le 5) {

$Freier_Speicherplatz = ((Get-WMIObject Win32_LogicalDisk -filter "name='c:'").freespace / 1GB)

Write-Host "Der freie Speicherplatz auf C: beträgt nur mehr $Freier_Speicherplatz GB, Bereinigungsmaßnahmen werden durchgeführt."

# SAGERUN-Keys erzeugen

$volumeCaches = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"

foreach($key in $volumeCaches)
{
    New-ItemProperty -Path "$($key.PSPath)" -Name StateFlags0099 -Value 2 -Type DWORD -Force | Out-Null
}

# DiskCleanup  ausführen

Start-Process -Wait "$env:SystemRoot\System32\cleanmgr.exe" -ArgumentList "/sagerun:99"

# SAGERUN-Keys entfernen

$volumeCaches = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
foreach($key in $volumeCaches)
{
    Remove-ItemProperty -Path "$($key.PSPath)" -Name StateFlags0099 -Force | Out-Null
}
}

### Funktion zum Senden von Messages via Telegram Bot

Write-Host "Telegram Zugangsdaten werden geladen und diverse Werte in Variablen gespeichert." -ForegroundColor Green

$BotKey = "..."
$ChatID = "..."

### Tagesdatum als Variable speichern

$Tagesdatum = get-date -format d	

### Updateprozess und Reporting

Write-Host "Der Updateprozess wird gestartet." -ForegroundColor Green

# Updatekriterien definieren

Write-Host "Suchkriterien werden definiert und ein Query an den definierten WSUS-Server geschickt." -ForegroundColor Yellow

$Criteria = "IsInstalled=0 and Type='Software'and IsHidden=0"

# Relevante Updates suchen

$UpdateSession = New-Object -ComObject 'Microsoft.Update.Session'
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
$SearchResult = $UpdateSearcher.Search($Criteria).Updates

# Ergebnis zählen und bei vollständig upgedatetem Client beenden

if ($SearchResult.Count -lt "1") {

Write-Host "Der Client ist vollständig upgedatet, der Prozess wird beendet." -ForegroundColor Yellow

# wuauclt /reportnow

wuauclt.exe /reportnow

Sleep -Seconds 500

Stop-Transcript

Exit}

# Updates herunterladen

$AnzahlUpdates = $SearchResult.Count

if ($SearchResult.Count -ge "1") {

Write-Host "Es wurden $AnzahlUpdates Updates gefunden. Es wird versucht, diese vom WSUS-Server herunterzuladen." -ForegroundColor Yellow

$Session = New-Object -ComObject Microsoft.Update.Session
$Downloader = $Session.CreateUpdateDownloader()
$Downloader.Updates = $SearchResult
$Downloadresult = $Downloader.Download()

}

# Bei Fehlern folgende Schritte ausführen

if (($Downloadresult.ResultCode) -eq "4") {

Write-Host "Der Download schlug fehl. Reparaturmaßnahmen werden ausgeführt." -ForegroundColor Red

# Dienste Windows Update und BITS anhalten

Stop-Service wuauserv -Force
Stop-Service bits -Force

# Windows Update Verzeichnis löschen

Remove-Item  “$env:windir\softwaredistribution” -Force -Recurse
Sleep -Seconds 30

# Dienste Windows Update und BITS starten, das Verzeichnis wird neu angelegt 

Start-Service wuauserv
Start-Service bits

sleep -Second 60

wuauclt /resetauthorization /detectnow

Sleep -Seconds 600

# erneute Suche nach Updates

$Criteria = "IsInstalled=0 and Type='Software'and IsHidden=0"
$Searcher = New-Object -ComObject Microsoft.Update.Searcher
$SearchResult = $Searcher.Search($Criteria).Updates
$Session = New-Object -ComObject Microsoft.Update.Session
$Downloader = $Session.CreateUpdateDownloader()
$Downloader.Updates = $SearchResult
$Downloadresult = $Downloader.Download()

Sleep -Seconds 600

# ResultCode 4: Download Failed

if (($Downloadresult.ResultCode) -eq "4") {

# $Nachricht = "Die Updates auf Client $Client konnten am $Tagesdatum nicht erfolgreich beendet werden."
# curl -s -X POST https://api.telegram.org/bot$BotKey/sendMessage -d chat_id=$ChatID -d text="$Nachricht_WUpdateClient"

Stop-Transcript

Exit

}}

# ResultCode 2: Download SuccessUpdates installieren

if (($Downloadresult.ResultCode) -eq "2") {Write-Host "Die Updates wurden erfolgreich heruntergeladen." -ForegroundColor Yellow}

# ResultCode 3: Download SucceededWithError

if (($Downloadresult.ResultCode) -eq "3") {

sleep -Seconds 300
$Criteria = "IsInstalled=0 and Type='Software'"
$Searcher = New-Object -ComObject Microsoft.Update.Searcher
$SearchResult = $Searcher.Search($Criteria).Updates
$Session = New-Object -ComObject Microsoft.Update.Session
$Downloader = $Session.CreateUpdateDownloader()
$Downloader.Updates = $SearchResult
$Downloadresult = $Downloader.Download()

if (($Downloadresult.ResultCode) -eq "3","4") {

$Nachricht = "Die Updates auf Client $Client konnten am $Tagesdatum nicht erfolgreich beendet werden."
curl -s -X POST https://api.telegram.org/bot$BotKey/sendMessage -d chat_id=$ChatID -d text="$Nachricht_WUpdateClient"
Invoke-WebRequest -Uri "https://api.telegram.org/bot$BotKey/sendMessage" -Method Post -ContentType "application/json;charset=utf-8" -Body (ConvertTo-Json -Compress -InputObject @{chat_id=$ChatID;text="$Nachricht"})

Stop-Transcript

Exit

}}

### Installation der Updates

Write-Host "Es wird versucht, die heruntergeladenen Updates zu installieren." -ForegroundColor Yellow
$Installer = New-Object -ComObject Microsoft.Update.Installer
$Installer.Updates = $SearchResult
$Result = $Installer.Install()

# ResultCode

# ResultCode 2: Install Successful

if (($Result.ResultCode) -eq "2") {

Write-Host "Die Installation war erfolgreich." -ForegroundColor Yellow

If (($Result.RebootRequired) -eq "True") {Write-Host "Der Updatevorgang war erfolgreich, der Rechner muss zum Abschluss neu gestartet werden" -ForegroundColor Yellow}
ELSE {Write-Host "Der Updatevorgang  war erfolgreich, Neustart ist nicht notwendig." -ForegroundColor Yellow}

# wuauclt /reportnow

wuauclt /reportnow

<# $Reportnow = New-Object -ComObject "Microsoft.Update.AutoUpdate"
$Reportnow.DetectNow() 
#>

Sleep -Seconds 300

Stop-Transcript
Exit

}

Stop-Transcript
Exit
