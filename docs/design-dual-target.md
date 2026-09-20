# Entwurf: zwei Ziele — `locations.xml` oder Verwaltungsschnittstelle

**Stand 2026-09-20. Noch nicht umgesetzt.** Diese Seite hält den Plan fest, damit daran
weitergearbeitet werden kann, ohne die Vorarbeit zu wiederholen. Die technische Grundlage
steht in [`ucserver-mgmt-api.md`](ucserver-mgmt-api.md).

## Die Leitplanke

**Der Dateiweg bleibt die Voreinstellung.** Wer `Manage-XmlRules.ps1` heute aufruft, darf
nichts merken. Bestehende Installationen laufen unverändert weiter; die Schnittstelle ist eine
*zusätzliche* Option, bis ein Kunde umstellen will.

Jede Änderung, die das bricht, ist falsch — auch wenn sie technisch schöner wäre.

## Die zweite Achse

Das Skript hat heute bereits eine Umschaltung der **Datenquelle**:

```
-OnlyApi     nur aus der OpenScape-4000-API
-SkipApi     nur aus der CSV
(keines)     beides
```

Dazu kommt eine Umschaltung des **Ziels**:

```
-Ziel File   locations.xml bearbeiten          (Vorgabe, heutiges Verhalten)
-Ziel Api    über Port 7221 schreiben          (neu)
```

Vier Kombinationen, alle sinnvoll:

| Quelle | Ziel | Fall |
|---|---|---|
| OSV-API | Datei | heute |
| CSV | Datei | heute |
| OSV-API | UCServer | Vollautomatik ohne Dateizugriff |
| CSV | UCServer | Einmalimport, Migration |

Die vier Aktionen `Add`, `Remove`, `List`, `Sync` bleiben unverändert und funktionieren gegen
beide Ziele. Was sich ändert, ist allein die unterste Schicht: statt `$xml.Save(...)` ein
`asnMgmtSetLocationInformationMgmt`.

## Neue Parameter

Die vorhandenen `-Api*`-Parameter meinen die **OpenScape-4000**-Seite. Damit es keine
Verwechslung gibt, heissen die neuen anders:

```powershell
-Ziel        <File|Api>    Vorgabe: File
-UcHost      <name>        UCServer
-UcPort      <int>         Vorgabe: 7221
-UcUser      <name>        Konto der UCServer-Verwaltung
-UcPass      <string>      besser: -UcCredential [pscredential]
-UcUseTls    [switch]      asnMgmtStartTLS aushandeln (setzt TLS am Server voraus)
-WhatIf                    Trockenlauf
```

**`-WhatIf` ist Pflicht, nicht Zierrat.** Die Operation überschreibt Rufnummernregeln eines
Produktivservers. Der Trockenlauf soll zeigen, welche Regeln hinzukämen, wegfielen und
unverändert blieben — und das muss er auch gegen das Ziel `Api` können, ohne zu schreiben
(lesen über 2634 genügt dafür).

## Aufbau

Drei Schichten, damit das Stück, das beide Werkzeuge brauchen, nur einmal existiert:

```
UcServerMgmt.psm1        Verbindung, Anmeldung, BER-Kodierung, Invoke, Antwort zerlegen
   ↑                     — wiederverwendbar, auch für andere Werkzeuge
LocationRules.psm1       lesen (2633/2634), schreiben (2635), Abbildung auf das Regelmodell
   ↑
Manage-XmlRules.ps1      Aktionen, Quellen, Protokollierung — im Wesentlichen wie heute
```

Die oberste Schicht sollte am Ende nicht mehr wissen, ob sie in eine Datei oder über das Netz
schreibt. Ein gemeinsames Zwischenmodell — eine Liste aus `Search`/`Replace` je Standort —
trennt beides sauber.

## Reihenfolge

1. **`UcServerMgmt.psm1`** bauen und gegen eine Testinstallation prüfen: verbinden, anmelden,
   `asnMgmtGetServerVersion` (2547) lesen. Wenn das durchläuft, steht die Werkzeugkette.
2. **Lesen** über 2633/2634, Ergebnis auf dasselbe Zwischenmodell abbilden, das der XML-Pfad
   erzeugt. Dann lässt sich `-Action List` gegen beide Ziele vergleichen — die beste
   Gegenprobe, die es gibt.
3. **`bIsReadOnlyFromFile` messen** (siehe offene Punkte in der Schnittstellenbeschreibung).
   Erst wenn klar ist, dass Datei- und API-Weg sich nicht gegenseitig blockieren, darf
   geschrieben werden.
4. **Schreiben** über 2635, zuerst nur mit `-WhatIf`, dann gegen eine Testumgebung.
5. Optional: **Ereignis 2707** abonnieren, um Änderungen von Hand zu erkennen.

## Was zum Prüfen fehlt

- Eine Testumgebung mit **mehreren Standorten**. Ohne die lässt sich das Zusammenspiel mit den
  automatisch erzeugten Regeln (`bAutoCreatedRule`) nicht beurteilen.
- Ein Verwaltungskonto auf dieser Testumgebung.
