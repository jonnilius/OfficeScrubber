# Office Scrubber

* Ein automatisiertes Skript zum Deinstallieren, Entfernen und Bereinigen von Microsoft
  Office (MSI oder Click-to-Run)

* Führt hauptsächlich OffScrub-VBS-Skripte aus, die aus dem SaRA-Tool 
  (Microsoft Support and Recovery Assistant) stammen

* Unterstützt das Bereinigen von Office 2003 und neuer unter Windows XP oder neuer

* Office (2024, 2021, 2019, 2016, 365) verwendet denselben Installationspfad und dieselbe 
  Lizenzierungsstruktur, was zu Lizenzkonflikten oder Duplikaten führen kann

Zusätzlich bleiben beim Deinstallieren von Office 2013+ unter Windows 8 oder neuer die 
Lizenzen im SPP-Token-Speicher des Systems zurück

## Übersicht: Scrub ALL

* Standardmäßig entfernt dieser Vorgang nur erkannte Versionen
unter Windows 7 und neuer werden Office C2R/2016 unabhängig von der Erkennung aktiviert
unter Windows Vista und älter wird Office 2010 unabhängig von der Erkennung aktiviert
zusätzlich werden Produktschlüssel deinstalliert und Lizenzreste bereinigt

* Mit den Zahlen 2–8 kannst du die Statuswerte für die Menüoptionen umschalten
und erzwingen, dass eine Office-Version bereinigt oder übersprungen wird

* Es wird empfohlen, nur die erkannten Versionen zu bereinigen {*}
alles manuell auszuwählen ist nicht nötig und dauert sehr lange.

## Übersicht: Hauptoptionen

* Scrub ALL

deinstalliert und entfernt eine oder mehrere Office-Versionen

* Scrub Office C2R

deinstalliert und entfernt Office Click-to-Run (365, LTSC, 2024, 2021, 2019, 2016, 2013)
wird unabhängig davon ausgeführt, ob diese Office-Version erkannt wird oder nicht

* Scrub Office 2016

deinstalliert und entfernt die MSI-Version von Office 2016
wird unabhängig von der Erkennung ausgeführt

* Scrub Office 2013

deinstalliert und entfernt die MSI-Version von Office 2013
wird unabhängig von der Erkennung ausgeführt

* Scrub Office 2010

deinstalliert und entfernt Office 2010 (MSI oder C2R)
wird unabhängig von der Erkennung ausgeführt

* Scrub Office 2007

deinstalliert und entfernt Office 2007
wird unabhängig von der Erkennung ausgeführt

* Scrub Office 2003

deinstalliert und entfernt Office 2003
wird unabhängig von der Erkennung ausgeführt

* Scrub Office UWP

deinstalliert und entfernt Office-Store-Apps unter Windows 11/10
wird unabhängig von der Erkennung ausgeführt

## Übersicht: Zusatzoptionen

* Clean vNext Licenses

entfernt Office-vNext-Lizenzen (Abo oder Lifetime), Tokens und zwischengespeicherte Identitäten

* Remove all Licenses

entfernt Lizenzen für Office 2013 und neuer (bei Konflikten)
du kannst danach Office reparieren, um die ursprünglichen Lizenzen wiederherzustellen,
oder C2R-R2V verwenden, um Office-C2R-Lizenzen zu installieren

* Reset C2R Licenses

entfernt Lizenzen für Office 2016 und neuer und installiert dann die ursprünglichen Office-C2R-Lizenzen neu
du kannst das verwenden, wenn die Office-Reparatur die Original-Lizenzen nicht wiederherstellen konnte
oder um C2R-R2V-Volumenlizenzen zu entfernen und wieder Retail-Lizenzen zu aktivieren

* Uninstall all Keys

deinstalliert Produktschlüssel für Office 2013 und neuer (bei Konflikten)

## Unattended-Kommandozeilenparameter

/P  
Scrub Office UWP

/C  
Scrub Office C2R

/M6  
Scrub Office 2016 MSI

/M5  
Scrub Office 2013 MSI

/M4  
Scrub Office 2010

/M2  
Scrub Office 2007

/M1  
Scrub Office 2003

/A  
Scrub ALL

* Hinweis:

Der Scrub-ALL-Parameter entfernt nur erkannte Versionen und die Standardversionen (wie oben beschrieben).
Um weitere oder andere Versionen unabhängig von der Erkennung zu erzwingen, gib ihre Parameter zusätzlich an.

Beispiel: Dies bereinigt erkannte und Standardversionen sowie Office 2013 und 2003:

OfficeScrubber.cmd /A /M2 /M5

