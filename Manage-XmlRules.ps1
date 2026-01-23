<#
.SYNOPSIS
    Estos Location Rules Synchronizer - Automatisiertes Standort-Routing für ProCall 8.

.DESCRIPTION
    Dieses Skript synchronisiert Teilnehmerdaten aus einer OpenScape 4000 (via api2hipath.exe) 
    mit der Konfigurationsdatei 'locations.xml' des Estos UCServers.
    
    HAUPTFUNKTIONEN:
    1. Cross-Site Logic: Teilnehmer werden nur in den Standorten als interne Regel hinterlegt, 
       denen sie NICHT angehören (verhindert Routing-Schleifen).
    2. RegEx-Kompression: Fasst Tisch- und Softphone-Nummern (Präfix 7) zu einer Regel 
       vom Typ '(^7?Durchwahl$)' zusammen.
    3. GUI-Schutz: Erkennt manuelle Regeln in der ProCall GUI anhand des Musters und 
       schützt diese vor dem Löschen.
    4. Integrität: Erstellt automatische Backups und erzwingt UTF-8-BOM Kodierung.

.PARAMETER Action
    Definiert die Hauptaufgabe:
    - Sync: Vollständiger Abgleich (Backup -> API -> Transformation -> XML-Update).
    - List: Tabellarische Anzeige aller aktuell aktiven Regeln.
    - Add: Manuelles Hinzufügen einer einzelnen Durchwahl.
    - Remove: Gezieltes Löschen einer Regel anhand der Durchwahl.

.PARAMETER OnlyApi
    Test-Modus. Führt nur den API-Export über Port 2013 TCP aus. Es werden keine 
    Änderungen an der ProCall Konfiguration vorgenommen.

.PARAMETER Value
    Die interne Durchwahl/Extension (z.B. 123). Erforderlich für -Action Add oder Remove.

.PARAMETER Replace
    Das Ziel-Routing-Format. Standardmäßig '+492150916\1'. Benötigt für -Action Add.

.PARAMETER CityCode
    Filtert Aktionen auf einen bestimmten Standort (z. B. 2132 für Büderich).

.PARAMETER SoftphonePrefix
    Die Ziffer zur Identifizierung von Softphones. Standard ist '7'.

.PARAMETER Path
    Absoluter Pfad zur 'locations.xml'. Standard ist der ProCall 8 Installationspfad.

.PARAMETER CSVPath
    Dateiname oder Pfad für den API-Export (Standard: 'PORT.csv').

.PARAMETER LogPath
    Pfad für das detaillierte Protokoll (Standard: 'sync_log.txt').

.PARAMETER Delimiter
    Trennzeichen der CSV-Datei (Standard: ';').

.PARAMETER ApiExe
    Vollständiger Pfad zur 'api2hipath.exe'.

.PARAMETER ApiHost
    Hostname oder IP-Adresse des OpenScape 4000 Assistant.

.PARAMETER ApiUser
    API-Benutzername (Standard: 'engr').

.PARAMETER ApiPass
    API-Passwort (Standard: '13370*').

.PARAMETER SkipApi
    Überspringt den API-Download und nutzt die lokal vorhandene PORT.csv für den Sync.

.EXAMPLE
    .\Manage-XmlRules6.ps1 -Action Sync
    Führt einen vollständigen nächtlichen Abgleich durch.

.EXAMPLE
    Get-Help .\Manage-XmlRules6.ps1 -Full
    Zeigt die vollständige Dokumentation des Skripts an.

.NOTES
    VERSION: 1.0 (Meilenstein 8)
    AUTHOR: Boris Hürtgen
    LICENSE: MIT License
    COLLABORATION: Created with the assistance of Google Gemini.
    PORT: Benötigt Port 2013 TCP zum OS4K Assistant.
    WARNUNG: Während schreibender Zugriffe muss der Dienst 'eucsrv' beendet sein!

.LINK
    https://github.com/bhuertgen/Estos-Location-Rules-Synchronizer/wiki
#>

# --- COPYRIGHT & LIZENZ ---
# (c) 2026 Boris Hürtgen
# Dieses Werk ist lizenziert unter der MIT-Lizenz.
# Erstellt in Zusammenarbeit mit Google Gemini.

param (
    [Parameter(Mandatory=$false)]
    [ValidateSet("Add", "Remove", "List", "Sync")]
    [string]$Action,

    [string]$Value,
    [string]$Replace,
    [string]$CityCode,
    
    [string]$SoftphonePrefix = "7",
    [string]$Path = "C:\Program Files\estos\UCServer\config\locations.xml",
    [string]$CSVPath = "PORT.csv", 
    [string]$LogPath = "sync_log.txt",

    # API Parameter (OpenScape 4000)
    [string]$ApiExe = "C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe",
    [string]$ApiUser = "<user>",
    [string]$ApiPass = "<password>*",
    [string]$ApiHost = "<ip>",
    
    [string]$Delimiter = ";", 

    [switch]$SkipApi,  
    [switch]$OnlyApi   
)

# --- LOGGING FUNKTION ---
function Write-Log {
    param (
        [Parameter(Mandatory=$true)] [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")] [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    $logMessage | Out-File -FilePath $absoluteLogPath -Append -Encoding UTF8
    
    $color = switch($Level) {
        "ERROR"   { "Red" }
        "WARNING" { "Yellow" }
        "SUCCESS" { "Green" }
        default   { "Gray" }
    }
    Write-Host $logMessage -ForegroundColor $color
}

# 1. INITIALISIERUNG
$absoluteLogPath = if ([System.IO.Path]::IsPathRooted($LogPath)) { $LogPath } else { Join-Path $PWD $LogPath }
$xmlItem = Get-Item $Path -ErrorAction SilentlyContinue

if (-not $Action -and -not $OnlyApi) {
    Write-Host "Fehler: -Action (Sync, List, Add, Remove) oder -OnlyApi erforderlich." -ForegroundColor Red
    return
}

Write-Log "================================================================"
Write-Log "START: $(if($OnlyApi){'API-Test'}else{$Action})"
Write-Log "Ziel-Datei: $Path"

# 2. API-ABFRAGE & BACKUP
if (($Action -eq "Sync" -and -not $SkipApi) -or $OnlyApi) {
    if (Test-Path $CSVPath) {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmm"
        Copy-Item $CSVPath "$($CSVPath)_$timestamp.bak" -Force
        Write-Log "Backup der PORT.csv erstellt." "SUCCESS"
    }

    Write-Log "Abruf von $ApiHost (Port 2013) wird gestartet..."
    $sqlClause = "1=1 ORDER BY e164_num"
    $apiArgs = @("-l", $ApiUser, "-p", $ApiPass, "-h", $ApiHost, "-o", "PORT", "-s", "e164_num,extension", "-c", $Delimiter, "-z", "-w", $sqlClause, $CSVPath)

    if (Test-Path $ApiExe) {
        try {
            $process = Start-Process -FilePath $ApiExe -ArgumentList $apiArgs -Wait -NoNewWindow -PassThru
            if ($process.ExitCode -eq 0) {
                Write-Log "API-Export erfolgreich abgeschlossen." "SUCCESS"
            } else {
                Write-Log "API-Fehler: Prozess beendet mit Code $($process.ExitCode)" "ERROR"
                if ($OnlyApi) { return }
            }
        } catch {
            Write-Log "Kritischer Fehler beim API-Aufruf: $_" "ERROR"
            if ($OnlyApi) { return }
        }
    } else {
        Write-Log "api2hipath.exe nicht gefunden: $ApiExe" "ERROR"
        if ($OnlyApi) { return }
    }
    if ($OnlyApi) { return }
}

# 3. XML VERARBEITUNG
if (-not $xmlItem) { Write-Log "XML-Datei fehlt unter $Path" "ERROR"; return }

if ($Action -in @("Add", "Remove", "Sync")) {
    $timestamp = Get-Date -Format "yyyyMMdd-HHmm"
    $xmlBackup = Join-Path $xmlItem.DirectoryName "$($xmlItem.BaseName)-$timestamp$($xmlItem.Extension)"
    Copy-Item $xmlItem.FullName $xmlBackup -Force
    Write-Log "XML-Sicherung erstellt: $xmlBackup" "SUCCESS"
}

$xml = New-Object System.Xml.XmlDocument
$xml.Load($xmlItem.FullName)
$allLocations = $xml.locations.location | Where-Object { $_.CityCode }
$validCityCodes = $allLocations.CityCode

# Helfer-Logik
function Get-PrefixFromE164($e164, $ext) { $prefix = $e164 -replace "$($ext)$", ""; return "+$prefix\1" }
function Get-BaseExtension($ext, $prefix) { if ($ext.StartsWith($prefix) -and $ext.Length -gt $prefix.Length) { return $ext.Substring($prefix.Length) }; return $ext }

# --- SYNC AKTION ---
if ($Action -eq "Sync") {
    if (-not (Test-Path $CSVPath)) { Write-Log "CSV fehlt!" "ERROR"; return }
    $csvData = Import-Csv -Path $CSVPath -Delimiter $Delimiter | Where-Object { $_.extension -and $_.e164_num }
    Write-Log "Starte Synchronisation (GUI-Schutz aktiv)..."

    $globalAdded = 0; $globalDeleted = 0; $statsTable = @()
    $escapedPrefix = [regex]::Escape($SoftphonePrefix)
    $toolPattern = "^\s*\(\^($escapedPrefix)\?\d+\$\)\s*$"

    foreach ($loc in $allLocations) {
        $currentCC = $loc.CityCode
        $locAdded = 0; $locDeleted = 0
        $internalRules = $loc.InternalRules
        if ($null -eq $internalRules) { continue }

        $allCurrentElements = @($internalRules.Element)
        $manualRulesCount = 0
        foreach ($el in $allCurrentElements) {
            if ($null -ne $el.Search -and $el.Search -notmatch $toolPattern) { $manualRulesCount++ }
        }

        $locData = $csvData | Where-Object { 
            $sourceCC = $_.e164_num.Substring(2, 4)
            ($sourceCC -ne $currentCC) -and ($validCityCodes -contains $sourceCC)
        }
        $grouped = $locData | Group-Object { Get-BaseExtension $_.extension $SoftphonePrefix }
        $desiredRules = foreach ($group in $grouped) {
            $baseExt = $group.Name; $first = $group.Group[0]
            [PSCustomObject]@{ Search = "(^$($SoftphonePrefix)?$baseExt$)"; Replace = Get-PrefixFromE164 $first.e164_num $first.extension }
        }

        foreach ($el in $allCurrentElements) {
            if ($null -eq $el) { continue }
            if ($el.Search -match $toolPattern) {
                if (-not ($desiredRules | Where-Object { $_.Search -eq $el.Search -and $_.Replace -eq $el.Replace })) {
                    $internalRules.RemoveChild($el) | Out-Null; $locDeleted++
                }
            }
        }

        foreach ($rule in $desiredRules) {
            if (-not ($internalRules.Element | Where-Object { $_.Search -eq $rule.Search -and $_.Replace -eq $rule.Replace })) {
                $newEl = $xml.CreateElement("Element"); $newEl.SetAttribute("Search", $rule.Search); $newEl.SetAttribute("Replace", $rule.Replace); $newEl.SetAttribute("MatchReplace", "0")
                $internalRules.AppendChild($newEl) | Out-Null; $locAdded++
            }
        }
        
        $totalRules = @($internalRules.Element).Count
        $statsTable += [PSCustomObject]@{ 
            Standort = $loc.name; CC = $currentCC; 'Neu' = [int]$locAdded; 'Geloescht' = [int]$locDeleted; 'Manuell' = [int]$manualRulesCount; 'Gesamt' = [int]$totalRules 
        }
        $globalAdded += $locAdded; $globalDeleted += $locDeleted
    }

    $utf8WithBOM = New-Object System.Text.UTF8Encoding($true)
    $writer = New-Object System.IO.StreamWriter($xmlItem.FullName, $false, $utf8WithBOM)
    $xml.Save($writer); $writer.Close()
    Write-Log "STATISTIK:`n$($statsTable | Format-Table -AutoSize | Out-String)" "SUCCESS"
}

# --- LISTE ---
elseif ($Action -eq "List") {
    $allLocations | ForEach-Object { $l = $_; $_.InternalRules.Element | Select-Object @{n="Standort";e={$l.name}}, Search, Replace } | Format-Table -AutoSize
}

# --- MANUELLE AKTIONEN ---
elseif ($Action -eq "Add") {
    $target = $allLocations | Where-Object { $_.CityCode -eq $CityCode }
    if ($target) {
        $newEl = $xml.CreateElement("Element"); $newEl.SetAttribute("Search", "(^$($SoftphonePrefix)?$Value$)"); $newEl.SetAttribute("Replace", $Replace); $newEl.SetAttribute("MatchReplace", "0")
        $target.InternalRules.AppendChild($newEl) | Out-Null
        $xml.Save($xmlItem.FullName); Write-Log "Manuelle Regel hinzugefügt." "SUCCESS"
    }
}
elseif ($Action -eq "Remove") {
    $target = $allLocations | Where-Object { $_.CityCode -eq $CityCode }
    if ($target) {
        $toRem = $target.InternalRules.Element | Where-Object { $_.Search -like "*$Value*" }
        foreach($r in $toRem){ $target.InternalRules.RemoveChild($r) | Out-Null }
        $xml.Save($xmlItem.FullName); Write-Log "Regel(n) entfernt." "SUCCESS"
    }
}

Write-Log "SCRIPT BEENDET.`n"
