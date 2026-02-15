# Estos Location Rules Synchronizer

Automatisiertes Standort-Routing für **Estos ProCall 8 Enterprise** — synchronisiert Teilnehmerdaten aus einer **Unify OpenScape 4000** mit den `InternalRules` der `locations.xml` des UCServers.

Das Skript ermöglicht **No-VNR-Routing**: Interne Durchwahlen können standortübergreifend ohne Verkehrsausscheidungsziffern gewählt werden.

---

## Inhaltsverzeichnis

- [Hauptfunktionen](#hauptfunktionen)
- [Systemvoraussetzungen](#systemvoraussetzungen)
- [Installation](#installation)
- [Parameter-Referenz](#parameter-referenz)
- [Aktionen im Detail](#aktionen-im-detail)
  - [Sync — Vollständiger Abgleich](#sync--vollständiger-abgleich)
  - [List — Regeln anzeigen](#list--regeln-anzeigen)
  - [Add — Manuelle Regel hinzufügen](#add--manuelle-regel-hinzufügen)
  - [Remove — Regel entfernen](#remove--regel-entfernen)
  - [OnlyApi — API-Testmodus](#onlyapi--api-testmodus)
- [api2hipath.exe — Aufruf im Detail](#api2hipathexe--aufruf-im-detail)
- [Interne Hilfsfunktionen](#interne-hilfsfunktionen)
- [CityCode-Extraktion](#citycode-extraktion)
- [Funktionsweise im Detail](#funktionsweise-im-detail)
  - [Cross-Site Logic](#cross-site-logic)
  - [RegEx-Kompression (Softphone-Präfix)](#regex-kompression-softphone-präfix)
  - [GUI-Schutz](#gui-schutz)
  - [Backup-Strategie](#backup-strategie)
  - [UTF-8-BOM Kodierung](#utf-8-bom-kodierung)
- [CSV-Format (PORT.csv)](#csv-format-portcsv)
- [XML-Struktur (locations.xml)](#xml-struktur-locationsxml)
- [Logging](#logging)
- [Automatisierung (Task Scheduler)](#automatisierung-task-scheduler)
- [Anwendungsbeispiele](#anwendungsbeispiele)
- [Fehlerbehebung](#fehlerbehebung)
- [Lizenz](#lizenz)

---

## Hauptfunktionen

| Funktion | Beschreibung |
|---|---|
| **Cross-Site Logic** | Teilnehmer werden nur in Standorten als interne Regel hinterlegt, denen sie **nicht** angehören — verhindert Routing-Schleifen |
| **RegEx-Kompression** | Fasst Tischtelefon- und Softphone-Nummern (Präfix `7`) zu einer einzelnen Regel `(^7?Durchwahl$)` zusammen |
| **GUI-Schutz** | Erkennt manuell in der ProCall-GUI erstellte Regeln anhand ihres Musters und schützt sie vor dem Löschen |
| **Integrität** | Erstellt automatische Backups vor jeder Änderung und erzwingt UTF-8-BOM Kodierung |
| **Reporting** | Detaillierte Statistiken pro Standort über hinzugefügte, gelöschte und geschützte Regeln |
| **Service Management** | Automatisierter Stop/Start des UCServer-Dienstes (`eucsrv`) im Rahmen der nächtlichen Synchronisation |

---

## Systemvoraussetzungen

- **Estos ProCall Enterprise 8.x** mit konfigurierter `locations.xml`
- **Unify OpenScape 4000** mit installiertem **api2hipath.exe** (Export Table Tool)
  - Download über den OpenScape 4000 Assistant
- **Netzwerk:** TCP-Port **2013** zum OpenScape 4000 Assistant
- **PowerShell 5.1+** mit **Administratorrechten**
- **Dienst `eucsrv`** muss während schreibender Zugriffe gestoppt sein

---

## Installation

1. Repository klonen oder herunterladen:
   ```powershell
   git clone https://github.com/bhuertgen/Estos-Location-Rules-Synchronizer.git
   ```

2. `api2hipath.exe` installieren (aus dem OpenScape 4000 Assistant herunterladen).

3. Parameter im Skript anpassen oder beim Aufruf übergeben (siehe [Parameter-Referenz](#parameter-referenz)).

---

## Parameter-Referenz

### Aktions-Parameter

| Parameter | Typ | Pflicht | Beschreibung |
|---|---|---|---|
| `-Action` | String | Ja* | Hauptaktion: `Sync`, `List`, `Add` oder `Remove` |
| `-OnlyApi` | Switch | Nein | API-Testmodus — führt nur den Export aus, keine XML-Änderungen |
| `-SkipApi` | Switch | Nein | Überspringt den API-Download und nutzt eine vorhandene `PORT.csv` |

\* Pflicht, wenn `-OnlyApi` nicht gesetzt ist.

### Werte-Parameter

| Parameter | Typ | Standard | Beschreibung |
|---|---|---|---|
| `-Value` | String | — | Interne Durchwahl (z.B. `123`). Pflicht bei `Add` und `Remove` |
| `-Replace` | String | — | Ziel-Routing-Format (z.B. `+492150916\1`). Pflicht bei `Add` |
| `-CityCode` | String | — | Filtert Aktionen auf einen bestimmten Standort (z.B. `2132`) |
| `-SoftphonePrefix` | String | `7` | Ziffer zur Identifizierung von Softphones |

### Pfad-Parameter

| Parameter | Typ | Standard | Beschreibung |
|---|---|---|---|
| `-Path` | String | `C:\Program Files\estos\UCServer\config\locations.xml` | Absoluter Pfad zur `locations.xml` |
| `-CSVPath` | String | `PORT.csv` | Dateiname oder Pfad für den API-Export |
| `-LogPath` | String | `sync_log.txt` | Pfad für das Protokoll |

### API-Parameter (OpenScape 4000)

| Parameter | Typ | Standard | Beschreibung |
|---|---|---|---|
| `-ApiExe` | String | `C:\Program Files (x86)\Unify\OpenScape 4000 Export Table\api2hipath.exe` | Pfad zur `api2hipath.exe` |
| `-ApiHost` | String | — | Hostname oder IP-Adresse des OpenScape 4000 Assistant |
| `-ApiUser` | String | — | API-Benutzername |
| `-ApiPass` | String | — | API-Passwort |

### Sonstige

| Parameter | Typ | Standard | Beschreibung |
|---|---|---|---|
| `-Delimiter` | String | `;` | CSV-Trennzeichen |

---

## Aktionen im Detail

### Sync — Vollständiger Abgleich

```powershell
.\Manage-XmlRules.ps1 -Action Sync
```

Ablauf eines vollständigen Sync-Zyklus:

1. **Backup** der vorhandenen `PORT.csv` (Zeitstempel-Suffix `.bak`)
2. **API-Export** — Abruf aller Teilnehmerdaten vom OpenScape 4000 via `api2hipath.exe` über Port 2013
3. **CSV-Validierung** — Prüfung, ob jede `extension` ein Suffix der zugehörigen `e164_num` ist
4. **XML-Backup** — Sicherungskopie der aktuellen `locations.xml`
5. **Cross-Site-Zuordnung** — Für jeden Standort werden nur Teilnehmer anderer Standorte verarbeitet
6. **RegEx-Kompression** — Gruppierung von Tisch-/Softphone-Nummern
7. **Regelabgleich:**
   - Veraltete tool-generierte Regeln werden gelöscht
   - Fehlende Regeln werden hinzugefügt
   - Manuelle (GUI-) Regeln bleiben unberührt
8. **XML speichern** — UTF-8-BOM-kodiert
9. **Statistik-Ausgabe** — Tabellarische Übersicht pro Standort

### List — Regeln anzeigen

```powershell
.\Manage-XmlRules.ps1 -Action List
.\Manage-XmlRules.ps1 -Action List -CityCode "2132"
```

Zeigt alle aktuell in der `locations.xml` hinterlegten `InternalRules` als Tabelle mit den Spalten **Standort**, **Search** und **Replace** an.

### Add — Manuelle Regel hinzufügen

```powershell
.\Manage-XmlRules.ps1 -Action Add -Value "123" -Replace "+492150916\1" -CityCode "2132"
```

Fügt eine einzelne interne Regel für die angegebene Durchwahl in den durch `-CityCode` spezifizierten Standort ein. Das Suchmuster wird automatisch im Format `(^7?123$)` generiert (mit dem konfigurierten Softphone-Präfix).

### Remove — Regel entfernen

```powershell
.\Manage-XmlRules.ps1 -Action Remove -Value "123" -CityCode "2132"
```

Entfernt alle Regeln, deren `Search`-Attribut die angegebene Durchwahl enthält (Wildcard-Suche `*123*`), aus dem spezifizierten Standort.

### OnlyApi — API-Testmodus

```powershell
.\Manage-XmlRules.ps1 -OnlyApi -ApiHost "10.1.2.3"
```

Führt ausschließlich den API-Export über Port 2013 TCP durch. Die `PORT.csv` wird erzeugt/aktualisiert, aber es werden **keine Änderungen** an der ProCall-Konfiguration vorgenommen. Ideal zum Testen der Netzwerkverbindung und API-Zugangsdaten.

---

## api2hipath.exe — Aufruf im Detail

Das Skript ruft die `api2hipath.exe` als externen Prozess auf, um die Teilnehmerdaten aus der OpenScape 4000 als CSV zu exportieren. Der vollständige Aufruf wird intern wie folgt zusammengebaut:

```
api2hipath.exe -l <ApiUser> -p <ApiPass> -h <ApiHost> -o PORT -s e164_num,extension -c ; -z -w "1=1 ORDER BY e164_num" PORT.csv
```

### Kommandozeilen-Parameter von api2hipath.exe

| Flag | Skript-Parameter | Beschreibung |
|---|---|---|
| `-l` | `-ApiUser` | Login-Benutzername für die OpenScape 4000 API |
| `-p` | `-ApiPass` | Passwort für die API-Authentifizierung |
| `-h` | `-ApiHost` | Hostname oder IP-Adresse des OpenScape 4000 Assistant |
| `-o` | *(fest: `PORT`)* | Name der abzufragenden Datenbanktabelle |
| `-s` | *(fest: `e164_num,extension`)* | Spaltenauswahl (SELECT-Klausel) |
| `-c` | `-Delimiter` | CSV-Trennzeichen (Standard: `;`) |
| `-z` | *(fest)* | Kopfzeile in CSV einfügen |
| `-w` | *(fest: `1=1 ORDER BY e164_num`)* | WHERE-Klausel — `1=1` selektiert alle Einträge, sortiert nach E.164-Nummer |
| *(positional)* | `-CSVPath` | Ausgabedatei (Standard: `PORT.csv`) |

### Ablauf des API-Aufrufs

1. Falls eine bestehende `PORT.csv` vorhanden ist, wird sie mit Zeitstempel gesichert
2. `api2hipath.exe` wird synchron via `Start-Process -Wait` gestartet
3. Der **Exit-Code** wird ausgewertet:
   - `0` = Erfolg
   - Alles andere = Fehler (wird als `ERROR` geloggt)
4. Im `-OnlyApi`-Modus bricht das Skript nach dem API-Aufruf ab — unabhängig vom Ergebnis
5. Im `-Action Sync`-Modus wird bei API-Fehler **nicht** abgebrochen, sondern mit der vorhandenen CSV weitergearbeitet

### Netzwerk-Anforderungen

- **Protokoll:** TCP
- **Port:** 2013
- **Richtung:** Skript-Host → OpenScape 4000 Assistant
- Firewalls zwischen den Systemen müssen diesen Port freigeben

---

## Interne Hilfsfunktionen

Das Skript definiert drei Hilfsfunktionen für die Datenverarbeitung:

### `Get-PrefixFromE164($e164, $ext)`

Leitet das Replace-Format aus der E.164-Nummer ab, indem die Extension am Ende entfernt und durch die Rückreferenz `\1` ersetzt wird.

```
Eingabe:  e164 = "492150916123", ext = "123"
Logik:    "492150916123" -replace "123$" → "492150916"
Ausgabe:  "+492150916\1"
```

### `Get-BaseExtension($ext, $prefix)`

Entfernt den Softphone-Präfix von einer Durchwahl, um die Basis-Extension für die RegEx-Gruppierung zu ermitteln.

```
Eingabe:  ext = "7123", prefix = "7"
Prüfung:  Beginnt mit "7" UND ist länger als 1 Zeichen
Ausgabe:  "123"

Eingabe:  ext = "123", prefix = "7"
Prüfung:  Beginnt NICHT mit "7"
Ausgabe:  "123" (unverändert)
```

### `Validate-CSV($e164, $ext)`

Prüft die Datenintegrität: Die Extension muss ein Suffix der E.164-Nummer sein. Bei Abweichung wird eine `WARNING` ins Log geschrieben.

```
OK:       e164 = "492150916123", ext = "123"  →  "123" ist Suffix ✓
Warnung:  e164 = "492150916999", ext = "123"  →  "123" ist kein Suffix ✗
```

---

## CityCode-Extraktion

Die Zuordnung eines Teilnehmers zu einem Standort erfolgt über den **CityCode**, der aus der E.164-Nummer extrahiert wird:

```powershell
$sourceCC = $_.e164_num.Substring(2, 4)
```

| Position | Bedeutung | Beispiel (`492150916123`) |
|---|---|---|
| `[0..1]` | Landesvorwahl (49 = Deutschland) | `49` |
| `[2..5]` | **CityCode** (4 Stellen) | `2150` |
| `[6..]` | Teilnehmeranschluss + Durchwahl | `916123` |

Der extrahierte CityCode wird gegen die Liste aller in der `locations.xml` definierten Standorte validiert (`$validCityCodes`). Nur Teilnehmer mit bekanntem CityCode werden verarbeitet.

---

## Funktionsweise im Detail

### Cross-Site Logic

Das Kernprinzip des Skripts verhindert **Routing-Schleifen** in Multi-Standort-Umgebungen:

```
Standort A (CityCode 2150)  ←  Teilnehmer von B und C
Standort B (CityCode 2132)  ←  Teilnehmer von A und C
Standort C (CityCode 2151)  ←  Teilnehmer von A und B
```

Für jeden Standort werden **nur Teilnehmer anderer Standorte** als interne Regeln eingetragen. Die Zuordnung erfolgt anhand des **CityCode**, der aus den ersten 4 Ziffern der `e164_num` (nach dem `+`) extrahiert wird:

```
e164_num: 492150916123
              ^^^^
          CityCode = 2150
```

Zusätzlich wird validiert, dass der extrahierte CityCode tatsächlich als Standort in der `locations.xml` existiert (`$validCityCodes`). Teilnehmer mit unbekanntem CityCode werden ignoriert.

### RegEx-Kompression (Softphone-Präfix)

In OpenScape 4000-Umgebungen haben Softphones typischerweise die gleiche Durchwahl wie das zugehörige Tischtelefon, aber mit einem vorangestellten Präfix (Standard: `7`).

**Beispiel:** Tischtelefon `123`, Softphone `7123`

Das Skript gruppiert beide Einträge zu **einer einzigen RegEx-Regel:**

```
Suchmuster:  (^7?123$)
```

Diese Regel matcht sowohl `123` als auch `7123`. Die Gruppierung erfolgt über die Hilfsfunktion `Get-BaseExtension`, die den Softphone-Präfix entfernt, um die Basis-Durchwahl zu ermitteln.

### GUI-Schutz

Administratoren können in der ProCall-GUI manuell Regeln anlegen. Diese sollen beim Sync nicht gelöscht werden.

Die Erkennung erfolgt über ein **Muster-Matching**: Nur Regeln, deren `Search`-Attribut dem Tool-Pattern entspricht, werden vom Skript verwaltet:

```regex
^\s*\(\^7?\d+\$\)\s*$
```

Dieses Pattern matcht exakt die vom Tool generierten Regeln im Format `(^7?Durchwahl$)`. Alle Regeln, die **nicht** diesem Muster entsprechen, gelten als manuell erstellt und werden in der Statistik separat als **„Manuell"** ausgewiesen.

### Backup-Strategie

Das Skript erstellt vor jeder schreibenden Operation automatische Backups:

| Datei | Backup-Format | Beispiel |
|---|---|---|
| `PORT.csv` | `PORT.csv_YYYYMMDD-HHmm.bak` | `PORT.csv_20260215-0200.bak` |
| `locations.xml` | `locations-YYYYMMDD-HHmm.xml` | `locations-20260215-0200.xml` |

Backups werden bei den Aktionen `Sync`, `Add` und `Remove` erstellt.

### UTF-8-BOM Kodierung

Die `locations.xml` wird explizit mit **UTF-8 BOM** (Byte Order Mark) gespeichert. Dies ist erforderlich, da der Estos UCServer diese Kodierung für die korrekte Verarbeitung von Sonderzeichen (Umlaute in Standortnamen) erwartet. Das Skript verwendet dazu `System.IO.StreamWriter` mit `System.Text.UTF8Encoding($true)`.

---

## CSV-Format (PORT.csv)

Die von `api2hipath.exe` exportierte CSV-Datei enthält zwei Spalten, getrennt durch Semikolon (`;`):

| Spalte | Beschreibung | Beispiel |
|---|---|---|
| `e164_num` | Vollständige E.164-Rufnummer (ohne `+`) | `492150916123` |
| `extension` | Interne Durchwahl/Nebenstelle | `123` |

**Beispiel-Inhalt:**
```csv
e164_num;extension
492150916123;123
4921509167123;7123
492132916456;456
```

**Validierung:** Das Skript prüft, ob jede `extension` ein Suffix der zugehörigen `e164_num` ist. Bei Inkonsistenz wird eine Warnung ins Log geschrieben.

---

## XML-Struktur (locations.xml)

Das Skript arbeitet auf der `locations.xml` des Estos UCServers. Relevanter Ausschnitt der Struktur:

```xml
<?xml version="1.0" encoding="utf-8"?>
<locations>
  <location name="Standort A" CityCode="2150" ...>
    <InternalRules>
      <Element Search="(^7?456$)" Replace="+492132916\1" MatchReplace="0" />
      <Element Search="(^7?789$)" Replace="+492151916\1" MatchReplace="0" />
    </InternalRules>
  </location>
  <location name="Standort B" CityCode="2132" ...>
    <InternalRules>
      <Element Search="(^7?123$)" Replace="+492150916\1" MatchReplace="0" />
    </InternalRules>
  </location>
</locations>
```

**Attribute je `Element`:**

| Attribut | Beschreibung |
|---|---|
| `Search` | RegEx-Suchmuster, z.B. `(^7?123$)` |
| `Replace` | Ersetzungsformat mit Rückreferenz, z.B. `+492150916\1` |
| `MatchReplace` | Immer `0` (Standardwert des UCServers) |

Das **Replace-Format** wird automatisch aus der `e164_num` abgeleitet: Die Extension wird am Ende entfernt und durch `\1` (Rückreferenz auf die Capture Group) ersetzt, mit vorangestelltem `+`.

---

## Logging

Alle Aktionen werden in eine Log-Datei geschrieben (Standard: `sync_log.txt`). Gleichzeitig erfolgt eine farbcodierte Konsolenausgabe.

**Log-Level:**

| Level | Farbe | Verwendung |
|---|---|---|
| `INFO` | Grau | Allgemeine Statusmeldungen |
| `SUCCESS` | Grün | Erfolgreiche Operationen (Backup, Export, Sync) |
| `WARNING` | Gelb | Validierungswarnungen (z.B. CSV-Inkonsistenz) |
| `ERROR` | Rot | Fehler (fehlende Dateien, API-Fehler) |

**Log-Format:**
```
[2026-02-15 02:00:00] [INFO] START: Sync
[2026-02-15 02:00:00] [INFO] Ziel-Datei: C:\Program Files\estos\UCServer\config\locations.xml
[2026-02-15 02:00:01] [SUCCESS] Backup der PORT.csv erstellt.
[2026-02-15 02:00:05] [SUCCESS] API-Export erfolgreich abgeschlossen.
[2026-02-15 02:00:05] [SUCCESS] XML-Sicherung erstellt: locations-20260215-0200.xml
[2026-02-15 02:00:06] [SUCCESS] STATISTIK:
Standort     CC   Neu Geloescht Manuell Gesamt
--------     --   --- --------- ------- ------
Standort A   2150   3         1       2     15
Standort B   2132   0         0       1     10
```

---

## Automatisierung (Task Scheduler)

Für den nächtlichen Betrieb kann das Skript über den **Windows Aufgabenplaner** automatisiert werden. Da der UCServer-Dienst (`eucsrv`) während schreibender Zugriffe gestoppt sein muss, empfiehlt sich ein Batch-Wrapper:

**Beispiel `sync-nightly.bat`:**
```batch
@echo off
REM UCServer-Dienst stoppen
net stop eucsrv

REM Synchronisation durchführen
powershell.exe -ExecutionPolicy Bypass -File "C:\Scripts\Manage-XmlRules.ps1" -Action Sync -ApiHost "10.1.2.3" -ApiUser "apiuser" -ApiPass "apipass"

REM UCServer-Dienst starten
net start eucsrv
```

**Empfohlene Aufgabenplaner-Einstellungen:**
- **Trigger:** Täglich, z.B. 02:00 Uhr
- **Aktion:** `sync-nightly.bat` ausführen
- **Ausführen als:** Lokaler Administrator
- **Ausführen, ob Benutzer angemeldet ist oder nicht:** Ja

---

## Anwendungsbeispiele

```powershell
# Vollständiger Sync (Standard-Workflow)
.\Manage-XmlRules.ps1 -Action Sync

# Sync mit expliziten API-Zugangsdaten
.\Manage-XmlRules.ps1 -Action Sync -ApiHost "10.1.2.3" -ApiUser "admin" -ApiPass "geheim"

# Sync ohne API-Abruf (vorhandene PORT.csv verwenden)
.\Manage-XmlRules.ps1 -Action Sync -SkipApi

# Nur API-Verbindung testen
.\Manage-XmlRules.ps1 -OnlyApi -ApiHost "10.1.2.3"

# Alle Regeln aller Standorte anzeigen
.\Manage-XmlRules.ps1 -Action List

# Regeln eines bestimmten Standorts anzeigen
.\Manage-XmlRules.ps1 -Action List -CityCode "2132"

# Einzelne Durchwahl manuell hinzufügen
.\Manage-XmlRules.ps1 -Action Add -Value "999" -Replace "+492150916\1" -CityCode "2132"

# Regel für eine Durchwahl entfernen
.\Manage-XmlRules.ps1 -Action Remove -Value "999" -CityCode "2132"

# Alternative XML-Datei verwenden
.\Manage-XmlRules.ps1 -Action List -Path "D:\Backup\locations.xml"

# Anderes Softphone-Präfix (z.B. 8 statt 7)
.\Manage-XmlRules.ps1 -Action Sync -SoftphonePrefix "8"
```

---

## Fehlerbehebung

| Problem | Ursache | Lösung |
|---|---|---|
| `XML-Datei fehlt unter ...` | `locations.xml` nicht am Standard-Pfad | `-Path` Parameter mit korrektem Pfad angeben |
| `api2hipath.exe nicht gefunden` | Export-Tool nicht installiert | Aus dem OpenScape 4000 Assistant herunterladen und installieren |
| `API-Fehler: Prozess beendet mit Code X` | Verbindungsproblem oder falsche Credentials | Firewall prüfen (Port 2013 TCP), Zugangsdaten prüfen, mit `-OnlyApi` testen |
| `CSV fehlt!` | Kein API-Export durchgeführt | Zuerst `-Action Sync` oder `-OnlyApi` ausführen; oder `-SkipApi` entfernen |
| `Nebenstelle X passt nicht zu Rufnummer Y` | Inkonsistente Daten in OpenScape 4000 | Teilnehmerdaten in der TK-Anlage prüfen |
| XML-Datei enthält Zeichensalat | Falsche Kodierung | Skript erzwingt UTF-8-BOM; prüfen, ob andere Tools die Datei überschrieben haben |
| Änderungen nicht sichtbar in ProCall | UCServer liest Konfiguration beim Start | Dienst `eucsrv` nach der Sync neu starten |

---

## Lizenz

MIT License — (c) 2026 Boris Hürtgen

Erstellt in Zusammenarbeit mit Google Gemini und aktualisiert auf GitHub via Claude Code.

[Wiki](https://github.com/bhuertgen/Estos-Location-Rules-Synchronizer/wiki) · [Issues](https://github.com/bhuertgen/Estos-Location-Rules-Synchronizer/issues) · [Lizenz](LICENSE)
