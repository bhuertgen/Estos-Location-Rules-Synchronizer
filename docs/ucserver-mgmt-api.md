# Die Verwaltungsschnittstelle des estos UCServer (Port 7221)

Diese Seite beschreibt, wie sich Standort- und Rufnummernregeln eines estos UCServer
**programmatisch** lesen und schreiben lassen — als Alternative zum direkten Bearbeiten von
`config\locations.xml`, wie es `Manage-XmlRules.ps1` heute tut.

> **Herkunft und Gültigkeit.** Diese Beschreibung stammt nicht von estos. Sie entstand am
> 2026-09-20 aus zwei Quellen: den **öffentlichen, MIT-lizenzierten ASN.1-Definitionen** von
> estos unter [`github.com/ESTOS/UCServerWebSDK2.0`](https://github.com/ESTOS/UCServerWebSDK2.0)
> (Verzeichnis `interface/asn1/`, Datei `ENetUC_Mgmt.asn1`) und aus eigenen Messungen gegen
> einen laufenden **UCServer 8.10.2.12087**.
>
> Die ASN.1-Dateien dort tragen den Stand **8.0.20240916** und sind damit rund 21 Monate
> älter als der geprüfte Server. Sie werden zudem gefiltert veröffentlicht
> (`esnacc -nodeprecated -noprivate`), enthalten also nicht alles. Wo unten „gemessen" steht,
> ist es an einer echten Installation nachvollzogen; wo „laut Definition" steht, stammt es
> allein aus der ASN.1-Datei.
>
> **Verbindlich ist das Produktverhalten, nicht diese Seite.**

---

## Warum überhaupt

`Manage-XmlRules.ps1` schreibt heute direkt in `locations.xml`. Das funktioniert, hat aber drei
bekannte Nachteile:

- Es braucht **Dateizugriff** auf den Server, also lokale Ausführung oder eine Freigabe.
- Der Server liest die Datei nur beim Start — eine Änderung wird erst nach einem
  **Dienstneustart** wirksam.
- Wer gleichzeitig die **UCServer-Verwaltung** geöffnet hat, überschreibt beim Speichern die
  Änderungen aus der Datei, und umgekehrt.

Über die Verwaltungsschnittstelle entfällt alles drei: Die Änderung geht an den laufenden
Server, der sie selbst persistiert und alle verbundenen Verwaltungsclients darüber
benachrichtigt.

---

## Verbindung

| | |
|---|---|
| Port | **7221** (Standard; konfigurierbar in `config\ListeningInterfaces.xml`) |
| Transport | TCP, ROSE über BER |
| Anmeldung | Konto der UCServer-Verwaltung |

**Zur Verschlüsselung.** Die Schnittstelle kennt eine TLS-Aushandlung über die Operation
`asnMgmtStartTLS` (2546) als erste Nachricht der Sitzung. Sie setzt voraus, dass TLS am
UCServer **aktiviert** und ein **gültiges Zertifikat** hinterlegt ist; ist das nicht der Fall,
antwortet der Server mit `ReturnError` und die Sitzung läuft unverschlüsselt weiter.
**Für den Produktivbetrieb sollte TLS eingerichtet und die Aushandlung verwendet werden** —
insbesondere, wenn der Port über ein Netz erreichbar ist, das nicht vollständig
vertrauenswürdig ist.

---

## Rahmung und Kodierung (gemessen)

Es gibt **keine eigene Rahmung**: Jede Nachricht ist ein einzelnes BER-TLV. Die Längen
summieren sich exakt auf den Bytestrom, ein Längenpräfix oder Trennzeichen existiert nicht.

```
A1 <len>   ROSE Invoke            A2 <len>   ROSE ReturnResult
   02 <len> <invokeID>               02 <len> <invokeID>
   02 <len> <operationID>            30 <len> { <resultValue> <result> }
   <argument>

A3 <len>   ROSE ReturnError       A4 <len>   ROSE Reject
```

Beispiel — die kürzestmögliche Nachricht, ein Invoke der Operation 2546 mit leerem Argument:

```
a1 09 02 01 01 02 02 09 f2 30 00
│     │        │           └─ argument: SEQUENCE {} (leer)
│     │        └───────────── operationID 0x09F2 = 2546
│     └────────────────────── invokeID 1
└──────────────────────────── ROSE Invoke
```

**Ereignisse** sendet der Server als Invokes mit der festen `invokeID` **99999**; sie werden
nicht beantwortet. Sie kommen teilweise **vor** dem `ReturnResult` des auslösenden Aufrufs —
ein Client muss also nach dem Senden weiterlesen, bis das Ergebnis mit der passenden
`invokeID` eintrifft.

Alle Module sind `IMPLICIT TAGS`. Kontext-Tags ab `[31]` brauchen die mehrbytige Tag-Form
(`9F 36` für `[54]`, `BF 39` für `[57]`) — ein häufiger Fehler bei eigenen Kodierern.

---

## Anmeldung

```
asnMgmtLogin ::= 2629

AsnMgmtLoginArgument ::= SEQUENCE {
    u8sUsername    UTF8String,   -- Konto der UCServer-Verwaltung
    u8sPassword    UTF8String,
    iVersion       INTEGER,      -- 20 (beobachtet)
    iDuplicate     INTEGER,      -- 1  (beobachtet)
    iSIPAVLines    INTEGER,      -- 1  (beobachtet)
    ...
}

AsnMgmtLoginResult ::= SEQUENCE { iOpenState INTEGER, ... }
```

Bei Misserfolg kommt ein `ReturnError` mit `AsnMgmtLoginError`, unter anderem
`eNETLOGININVALUSERNAME(101)`, `eNETLOGININVALIDPASS(102)`,
`eNETLOGINUSERNOTENABLED(107)`.

> **Der UCServer hat einen Brute-Force-Schutz.** Gemessen: drei Fehlversuche sperren für
> 60 Sekunden, zehn für 3600. Ein Werkzeug sollte eine Verbindung offen halten und mehrere
> Aufrufe darüber schicken, statt sich wiederholt neu anzumelden.

---

## Die Standort-Operationen

| Kennung | Operation | Zweck |
|--------:|-----------|-------|
| 2633 | `asnMgmtGetLocationIDs` | alle Standorte auflisten |
| 2634 | `asnMgmtGetLocationInformationMgmt` | einen Standort vollständig lesen |
| 2635 | `asnMgmtSetLocationInformationMgmt` | einen Standort schreiben |
| 2577 | `asnMgmtDeleteLocation` | einen Standort löschen |
| 2707 | `asnMgmtLocationChanged` | **Ereignis**: ein Standort wurde geändert |

```
AsnMgmtGetLocationInformationMgmtArgument ::= SEQUENCE { u8sLocationID UTF8String, ... }
AsnMgmtGetLocationInformationMgmtResult   ::= SEQUENCE { asnLocationData AsnLocationInformationMgmt, ... }
```

Das Ereignis 2707 ist der eigentliche Gewinn gegenüber dem Dateiweg: Ein Synchronisationsdienst
erfährt damit, wenn jemand die Regeln von Hand ändert — statt sie beim nächsten Lauf stumm zu
überschreiben.

---

## Die Regeln: Abbildung XML ↔ API

Das ist der Kern für `Manage-XmlRules.ps1`. Die Struktur ist deckungsgleich.

| `locations.xml` | ASN.1 |
|---|---|
| `<Location name="…">` | ein Eintrag aus `asnMgmtGetLocationIDs` |
| `<InternalRules>` | `numberFormatInternal` — `AsnNumberFormatRegExpressionsMgmt` |
| `<Element>` | ein Eintrag in `formatRules` |
| `<Search>` | `u8sSearch` |
| `<Replace>` | `u8sReplace` |

```
AsnNumberFormatRegExpressionsMgmt ::= SEQUENCE {
    bIsReadOnlyFromFile    BOOLEAN,
    formatRules            AsnNumberFormatRegExpressListMgmt,
    ...
}

AsnNumberFormatRegExpressionMgmt ::= SEQUENCE {
    u8sSearch              UTF8String,
    u8sReplace             UTF8String,
    bMatchReplace          BOOLEAN,
    bAutoCreatedRule       BOOLEAN,
    u8sAutoRuleforLocation UTF8String,
    ...
}
```

**Zwei Felder gibt es nur über die API**, und beide sind nützlich:

- **`bAutoCreatedRule`** — der Server erzeugt bei mehreren Standorten selbst Regeln
  (wenn `m_bMultiLocationEnabled` und `m_bAutoCreateOtherLocationRules` gesetzt sind).
  Über dieses Flag lassen sich selbst gepflegte von automatisch erzeugten Regeln
  unterscheiden. Über die XML ist das nicht zuverlässig möglich.
- **`u8sAutoRuleforLocation`** — bei automatisch erzeugten Regeln der Standort, für den sie
  gilt.

Neben `numberFormatInternal` führt der Standort weitere Regellisten, die dieselbe Struktur
verwenden: `numberFormatExternal`, `numberFormatRegExpressListIn`,
`numberFormatRegExpressListOut` sowie die beiden automatischen Listen
`numberFormatAutomaticOtherLocationsDDItoSC` und `…SCtoDDI`.

---

## Offene Punkte

Ehrlich benannt, statt geraten:

1. **`bIsReadOnlyFromFile`.** Der Name legt nahe, dass der Server Regeln als schreibgeschützt
   markiert, wenn sie aus `locations.xml` stammen. Falls das so ist, könnten sich Datei- und
   API-Weg gegenseitig blockieren. **Das muss gemessen werden, bevor beide Wege nebeneinander
   angeboten werden.**
2. **Zusammenspiel mit der Automatik.** Wie verhält sich `asnMgmtSetLocationInformationMgmt`
   zu automatisch erzeugten Regeln? Werden sie überschrieben, erhalten, neu berechnet?
   Erfordert eine Testumgebung mit **mehreren** Standorten.
3. **Teilschreiben.** Ob 2635 den gesamten Standort ersetzt oder einzelne Felder übernimmt,
   ist ungeprüft. Bis das geklärt ist: vorher mit 2634 lesen, ändern, vollständig
   zurückschreiben.
4. **Die Zuordnung Kennung → Name** stammt aus der ASN.1-Datei vom Stand 09/2024. Die
   Operationskennungen sind offensichtlich stabil (2546 liess sich an der laufenden 8.10.2
   bestätigen), neuere Operationen fehlen dort aber.
5. Es gibt Operationen in der veröffentlichten Datei, die ein konkreter Server **nicht**
   bereitstellt — er antwortet dann mit `ROSEReject { unrecognisedOperation }` und läuft
   normal weiter. Ein Client sollte das abfangen.

---

## Quellen

- [`github.com/ESTOS/UCServerWebSDK2.0`](https://github.com/ESTOS/UCServerWebSDK2.0) —
  ASN.1-Module, Protokolldokumentation, generierte Stubs. **MIT License, Copyright (c) 2024
  estos GmbH.** Dort ist `interface/asn1/ENetUC_Mgmt.asn1` die maßgebliche Datei.
- [`github.com/ESTOS/esnacc`](https://github.com/ESTOS/esnacc) — der ASN.1-Compiler, mit dem
  estos diese Definitionen erzeugt, samt `ROSE/SNACCROSE.asn1`. GPL-2.0.
- Eigene Messungen gegen UCServer 8.10.2.12087.

Die beschriebene Schnittstelle ist Eigentum der estos GmbH. Diese Seite ist keine
estos-Dokumentation und von estos weder geprüft noch freigegeben.
