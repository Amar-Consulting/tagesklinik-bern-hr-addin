# Testanleitung: tagesklinik | bern HR-Dokumente auf Mac

Diese Anleitung beschreibt, wie das Word Office Add-in testweise auf einem Mac geladen wird.

## Voraussetzung

- Microsoft Word für Mac
- Internetverbindung
- Datei manifest-github-pages.xml

## Schritt 1: Word komplett beenden

In Word:

Word → Word beenden

Nicht nur das Dokument schließen!

## Schritt 2: WEF-Ordner öffnen oder anlegen

Im Finder oben klicken:

Gehe zu → Gehe zum Ordner …

Dann diesen Pfad einfügen:

~/Library/Containers/com.microsoft.Word/Data/Documents/wef

Anschließend auf Öffnen klicken.

Falls der Ordner geöffnet wird, ist alles bereit.

Falls eine Meldung erscheint, dass der Ordner nicht gefunden wurde, muss der Ordner angelegt werden.

Dazu im Finder wieder klicken:

Gehe zu → Gehe zum Ordner …

Dann diesen Pfad einfügen:

~/Library/Containers/com.microsoft.Word/Data/Documents

Dort einen neuen Ordner erstellen:

1. Rechtsklick in den Ordner Documents
2. Neuer Ordner auswählen
3. Den neuen Ordner exakt so benennen:

wef

Wichtig: Der Ordnername muss klein geschrieben sein:

wef

Danach den Ordner wef öffnen.
## Schritt 3: Manifest kopieren

Die Datei:

manifest-github-pages.xml

in diesen Ordner kopieren:

~/Library/Containers/com.microsoft.Word/Data/Documents/wef

## Schritt 4: Word neu starten

Microsoft Word öffnen.

Dann im Menüband prüfen:

Start → Add-ins

oder nach dem Button suchen:

HR-Dokument erstellen

## Schritt 5: Add-in testen

Ein Word-Dokument mit Platzhaltern öffnen, zum Beispiel:

[DATUM]

[ANREDE] [TITEL][NACHNAME]

Wir freuen uns, [VORNAME] [NACHNAME] als [BELEGARZT] begrüßen zu dürfen.

Im Add-in Daten eingeben und auf:

Dokument erstellen

klicken.

## Unterstützte Platzhalter

[DATUM]
[ANREDE]
[TITEL]
[VORNAME]
[NACHNAME]
[BELEGARZT]

## Hinweis

Dies ist eine Pilot-/Testinstallation. Für den produktiven Einsatz sollte das Add-in später zentral über Microsoft 365 bereitgestellt werden.