# Office Scrubber v13

Ein automatisiertes Skript zum vollständigen Entfernen, Deinstallieren und Bereinigen von Microsoft Office (MSI und Click‑to‑Run).

Das Skript nutzt hauptsächlich die OffScrub‑VBS‑Module aus dem SaRA‑Tool (Microsoft Support and Recovery Assistant). 

Es unterstützt Office 2003 und neuer, unter Windows XP oder neuer.

Da moderne Office‑Versionen (2024, 2021, 2019, 2016, 365) denselben Installationspfad und dieselbe Lizenzierungsstruktur verwenden, kann es zu Konflikten oder doppelten Lizenzeinträgen kommen. 

Zusätzlich bleiben bei Office 2013 und neuer unter Windows 8 oder neuer Lizenzreste im SPP‑Token‑Speicher erhalten.

---

## Scrub ALL – Übersicht

- Entfernt standardmäßig nur erkannte Office‑Installationen.
    
- Unter Windows 7 und neuer werden Office C2R/2016 zusätzlich aktiviert, auch wenn sie nicht erkannt werden.
    
- Unter Windows Vista und älter wird Office 2010 zusätzlich aktiviert, unabhängig von der Erkennung.
    
- Produktschlüssel und Lizenzreste werden automatisch bereinigt.
    

Du kannst mit den Tasten **2–8** festlegen, ob bestimmte Office‑Versionen erzwungen bereinigt oder übersprungen werden.

**Empfehlung:** Nur erkannte Office‑Versionen bereinigen. Manuelle Komplettauswahl verlängert den Vorgang erheblich.

---

## Hauptoptionen

### Scrub ALL

Entfernt eine oder mehrere Office‑Versionen (siehe Verhalten oben).

### Scrub Office C2R

Entfernt Office Click‑to‑Run (365, LTSC, 2024, 2021, 2019, 2016, 2013).  
Wird immer ausgeführt – auch wenn die Installation nicht erkannt wird.

### Scrub Office 2016 (MSI)

Entfernt die MSI‑Version von Office 2016.  
Wird unabhängig von der Erkennung ausgeführt.

### Scrub Office 2013 (MSI)

Entfernt die MSI‑Version von Office 2013.  
Wird unabhängig von der Erkennung ausgeführt.

### Scrub Office 2010 (MSI oder C2R)

Entfernt Office 2010.  
Wird unabhängig von der Erkennung ausgeführt.

### Scrub Office 2007

Entfernt Office 2007.  
Wird unabhängig von der Erkennung ausgeführt.

### Scrub Office 2003

Entfernt Office 2003.  
Wird unabhängig von der Erkennung ausgeführt.

### Scrub Office UWP

Entfernt Office‑Store‑Apps unter Windows 10 und 11.  
Wird unabhängig von der Erkennung ausgeführt.

---

## Zusatzoptionen

### Clean vNext Licenses

Entfernt Office‑vNext‑Lizenzen (Abo und Lifetime), Token‑Daten und gespeicherte Identitäten.

### Remove all Licenses

Entfernt alle Office‑Lizenzen für Office 2013 und neuer (nützlich bei Lizenzkonflikten).

- Danach kannst du Office reparieren, um die Original‑Lizenzen wiederherzustellen.
    
- Alternativ kannst du C2R‑R2V verwenden, um passende C2R‑Lizenzen zu installieren.
    

### Reset C2R Licenses

Entfernt Office‑Lizenzen für Office 2016 und neuer und installiert anschließend die Original‑C2R‑Lizenzen neu.  
Ideal, wenn eine Office‑Reparatur die ursprünglichen Lizenzen nicht wiederherstellen konnte, oder um R2V‑Volumenlizenzen zurück auf Retail zu setzen.

### Uninstall all Keys

Entfernt Produktschlüssel für Office 2013 und neuer (z. B. bei Key‑Konflikten).

---

## Unattended‑Parameter für die Kommandozeile

|Parameter|Aktion|
|---|---|
|`/P`|Scrub Office UWP|
|`/C`|Scrub Office C2R|
|`/M6`|Scrub Office 2016 (MSI)|
|`/M5`|Scrub Office 2013 (MSI)|
|`/M4`|Scrub Office 2010|
|`/M2`|Scrub Office 2007|
|`/M1`|Scrub Office 2003|
|`/A`|Scrub ALL|

### Hinweis

Der Parameter **/A** entfernt nur erkannte und die standardmäßigen Zusatzversionen.  
Um weitere Versionen unabhängig von der Erkennung zu bereinigen, gib deren Parameter zusätzlich an.

**Beispiel:** Bereinigt erkannte Versionen, die Standardversionen sowie explizit Office 2013 und 2003:

```
OfficeScrubber.cmd /A /M2 /M5
```
