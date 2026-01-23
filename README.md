ÜBERSICHT
    Estos Location Rules Synchronizer - Automatisiertes Standort-Routing für ProCall 8.


SYNTAX

    Manage-XmlRules.ps1 [[-Action] <String>]
    [[-Value] <String>] [[-Replace] <String>] [[-CityCode] <String>] [[-SoftphonePrefix] <String>] [[-Path] <String>]
    [[-CSVPath] <String>] [[-LogPath] <String>] [[-ApiExe] <String>] [[-ApiUser] <String>] [[-ApiPass] <String>]
    [[-ApiHost] <String>] [[-Delimiter] <String>] [-SkipApi] [-OnlyApi] [<CommonParameters>]


BESCHREIBUNG
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


VERWANDTE LINKS
    https://github.com/bhuertgen/Estos-Location-Rules-Synchronizer/wiki
