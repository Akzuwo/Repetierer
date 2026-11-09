# DESIGN.md

## 1. Zielbild

Repetierer soll wie eine einzige, konsistente Desktop-App wirken: ruhig, klar, modern, hochwertig und funktional.

Der visuelle Wow-Effekt entsteht nicht durch möglichst viele Effekte, Karten oder Farben, sondern durch:

- klare Informationshierarchie
- saubere Abstände
- konsistente Komponenten
- gezielte Akzentfarbe
- subtile Tiefe
- kurze, saubere Animationen
- klaren Fokus auf die aktuelle Aufgabe

Neue Screens, Modals und Funktionen müssen sich in die bestehende Designsprache einfügen. Sie dürfen nicht wie Komponenten aus einer anderen App wirken.

---

## 2. Grundprinzipien

### 2.1 Eine visuelle Sprache

Alle Screens verwenden dieselben Regeln für:

- Header
- Typografie
- Buttons
- Cards und Panels
- Modals
- Alerts
- Eingabefelder
- Navigation
- Icons
- Abstände
- Overlays
- Animationen

Keine neue Variante einer bestehenden Komponente einführen, wenn bereits ein passendes Muster existiert.

### 2.2 Nur den aktuellen Zustand zeigen

Die UI zeigt nur Elemente, die im aktuellen Schritt relevant sind.

Beispiel Session-Flow:

1. Klasse wählen
2. Abwesenheiten erfassen
3. Session starten
4. Repetition / Mitarbeit
5. Session beenden

Nicht bereits die spätere Repetitionsoberfläche ausgegraut anzeigen, solange noch keine Klasse gewählt wurde.

### 2.3 Fokus vor Informationsmenge

Auf jedem Screen muss sofort klar sein:

- Wo bin ich?
- Was ist der aktuelle Zustand?
- Was ist die wichtigste Information?
- Was ist die primäre Aktion?

Sekundäre Aktionen dürfen nicht visuell mit der Hauptaktion konkurrieren.

### 2.4 Bestehende Funktionalität schützen

UI-Arbeiten dürfen nicht unnötig Logik verändern.

Bei Designänderungen:

- bestehende Funktionen erhalten
- bestehende Datenflüsse respektieren
- nur notwendige Komponenten anfassen
- keine ungefragten Refactorings
- keine neuen Abhängigkeiten nur für kleine optische Details

---

# 3. Globale App-Struktur

## 3.1 Header

Der globale Header bleibt auf allen Hauptscreens identisch.

Links:

```text
☰  Repetierer
```

Rechts:

```text
KI   Einstellungen   Fenstersteuerung
```

Regeln:

- Der Appname ist immer `Repetierer`.
- Funktionsnamen wie `Mitmachnote` ersetzen den Appnamen nie.
- Autoreninformationen gehören nicht in den Header.
- Versionsinformationen gehören nicht in den Header.
- Headerhöhe und Innenabstände sind überall identisch.
- Header-Icons verwenden dieselbe Grösse und Strichstärke.
- Der Header soll ruhig wirken und nicht mit dem Seiteninhalt konkurrieren.

Autoren- und Versionsinformationen gehören in einen `Über Repetierer`-Bereich oder dezent in den unteren Bereich des Drawers.

---

## 3.2 Seitenhierarchie

Normale Inhaltsseiten verwenden klare Titel ohne vorgeschaltete Eyebrows:

```text
Seitentitel
Kurze Beschreibung, falls nötig

Hauptinhalt

Aktionen
```

Beispiel:

```text
Beurteilungsrunden                     Klasse L23a

[Inhalt]
```

---

# 4. Typografie

Es gibt klar definierte Ebenen.

Die App verwendet die lokal eingebundene variable Geist-Schrift von Vercel: 28 px für die Klassenauswahl, 20 px für Ansichten/Dialoge, 14 px für Inhalt und 12–13 px für Navigation/Metadaten. Medium (500) und Semibold (600) tragen die Hierarchie. Zahlen verwenden tabellarische Ziffern. Der Schülername bleibt gross (bis 88 px), in normaler Schreibweise; lange Namen skalieren auf die tatsächlich verfügbare Breite. Die bestehende Splash-Typografie bleibt getrennt in Roboto.

## Display

Nur für zentrale Inhalte:

- aktuell gezogener Schüler
- besonders wichtige Sessioninformation

Sehr gross und sparsam verwenden.

## H1

Seitentitel.

Beispiele:

- Klasse wählen
- Einstellungen
- Beurteilungsrunde

## H2

Bereichstitel innerhalb eines Screens.

## Body

Normaler Beschreibungstext.

## Label

Für:

- Buttons
- Inputs
- kleine UI-Beschriftungen

## Muted

Für:

- Zeitstempel
- Versionsnummer
- Sekundärinformationen
- Hilfstexte

---

# 5. Farben

## 5.1 Dunkle Ebenen

Nicht jede Fläche soll exakt dieselbe Farbe besitzen.

Verwende subtile Tiefenstufen:

1. App Background
2. Surface / Panel
3. Elevated Surface
4. Modal Surface

Unterschiede sollen subtil bleiben.

## 5.2 Akzentblau

Blau wird verwendet für:

- aktive Zustände
- primäre Aktionen
- Fokus
- aktive Navigation
- wichtige Highlights

Blau nicht gleichzeitig auf zu vielen Elementen verwenden.

## 5.3 Statusfarben

Einheitlich:

- Info: Blau
- Success: Grün
- Warning: Gelb
- Error: Rot
- Danger: Rot

---

# 6. Spacing

Verwende vorzugsweise diese Abstufungen:

```text
4 px
8 px
12 px
16 px
24 px
32 px
48 px
64 px
```

Regeln:

- ähnliche Komponenten verwenden gleiche Abstände
- keine zufälligen Zwischenwerte ohne Grund
- Inhalte nicht unnötig an den oberen Rand kleben
- grosse leere Bereiche bewusst ausbalancieren
- Container benötigen konsistente Innenabstände

---

# 7. Buttons

Es gibt vier Button-Typen.

## Primary

Für die wichtigste Aktion.

Beispiele:

- Start
- Speichern
- Beurteilung öffnen
- Antworten aktualisieren

Darstellung:

- klar hervorgehoben
- Akzentblau
- eindeutiger Hover- und Focus-State

## Secondary

Für normale Alternativen.

Beispiele:

- Abbrechen
- Zurück
- Chancen

Darstellung:

- dunkle Fläche
- dezenter Rand
- weniger dominant als Primary

## Ghost

Für kleine Toolbar- und Nebenaktionen.

Beispiele:

- Hilfe
- Aktualisieren
- Icon-Aktionen

## Danger

Nur für destruktive Aktionen.

Beispiele:

- Runde löschen
- Daten löschen
- Session beenden

Danger-Aktionen dürfen nicht wie normale Navigation wirken.

## Grössen

Standard:

```text
Höhe: 36 px
Radius: 6 px
```

Small:

```text
Höhe: 28–32 px
```

Large nur für spezielle Hauptaktionen:

```text
Höhe: ca. 52–56 px
```

Keine Mischung aus extrem runden Pills, fast rechteckigen Buttons und unterschiedlichen Höhen auf demselben Screen.

---

# 8. Cards und Panels

Cards dienen der Gruppierung, nicht der Dekoration.

Regeln:

- subtile Abhebung vom Hintergrund
- konsistenter Radius
- konsistente Innenabstände
- keine unnötigen Rahmen
- keine übertriebenen Schatten
- Schatten nur für tatsächlich überlagerte Flächen wie Dialoge und Drawer

Nicht jede Textgruppe benötigt eine Card.

Klassen, Mitarbeit und Beurteilungsrunden werden als flache Zeilen dargestellt. Die Startansicht und die Hauptansichten haben keine umrahmende Card. Statistikgruppen werden durch Abstand und Trennlinien gegliedert.

---

# 9. Inputs

Alle Eingabefelder verwenden:

- gleiche Höhe
- gleichen Radius
- gleiche Hintergrundfarbe
- gleiche Label-Struktur
- konsistente Focus-Darstellung
- Akzentblau im Focus-State
- keinen übertriebenen Glow

---

# 10. Modals

## 10.1 Standardstruktur

Normale Modals verwenden:

```text
Titel
Kurze Beschreibung

Inhalt

--------------------------------

Abbrechen    Primäre Aktion
```

Regeln:

- normalerweise linksbündig
- die Dialogbox immer horizontal und vertikal im Overlay zentrieren; lange Inhalte innerhalb der Box scrollen
- konsistente Breite
- konsistente Innenabstände
- konsistenter Footer
- gleiche Button-Höhen
- keine wechselnden Modal-Sprachen

Grosse Warnsymbole nur, wenn wirklich notwendig.

## 10.2 Overlay

- Hintergrund abdunkeln
- moderater Blur
- darunterliegenden Kontext weiterhin erkennen lassen

Das Modal soll über dem aktuellen Screen liegen und nicht wie ein neuer Screen wirken.

## 10.3 Bestätigungsdialoge

Beispiel:

```text
SESSION

Session beenden?

Alle bisherigen Einträge bleiben gespeichert.
Du gelangst danach zurück zur Klassenauswahl.

Abbrechen    Session beenden
```

Ein normales Session-Ende nicht unnötig dramatisieren.

---

# 11. Drawer / Navigation

Der Drawer ist Navigation, kein Stapel grosser Aktionsbuttons.

Bereiche:

- Repetition
- Mitmachnote
- Datenverwaltung
- Protokoll
- Statistiken

Regeln:

- kompakte Zeilen
- aktive Seite mit Akzentblau hervorheben
- inaktive Einträge ruhig darstellen
- Icons aus derselben Familie
- keine übergrossen Pills
- Hilfe und Status dezent platzieren

Unten können stehen:

- Version
- Autoren
- App-Status

---

# 12. Startseite / Klassenauswahl

Die Klassenauswahl ist ein eigener Zustand.

Solange keine Klasse gewählt wurde, nicht anzeigen:

- Repetition
- Schülername
- Note
- deaktivierte Sessionbuttons
- inaktive Hauptaktionen

Zielstruktur:

```text
Klasse wählen

Wähle eine Klasse und erfasse die heutigen Abwesenheiten.

1 Klasse wählen   2 Abwesenheiten   3 Session starten

L23a                                      >
M24b                                      >
K25c                                      >

3 Klassen · Unterricht.xlsx    Datei wechseln
```

Der Screen soll ruhig, klar und bewusst leer wirken.

Die Liste ist auf 720 px begrenzt und scrollt bei vielen Klassen. Ohne Datei zeigt derselbe Aufbau die Aktionen zum Öffnen und Importieren. Keine verschachtelten Panels, dekorativen Icons oder Marketingtexte.

---

# 13. Session-Flow

Der Session-Flow soll visuell eindeutig sein.

## Zustand A: Klasse wählen

Nur Klassenauswahl.

## Zustand B: Abwesenheiten

Klare Auswahl der abwesenden Schüler.

Primäraktion:

```text
Session starten
```

## Zustand C: Aktive Session

Nur jetzt erscheinen:

- Repetitionssteuerung
- Mitarbeit
- Schülerauswahl
- Bewertung
- Sessionaktionen

## Zustand D: Session beenden

Ein standardisierter Bestätigungsdialog.

---

# 14. Aktive Repetition

Die aktive Repetition ist der wichtigste Screen der App.

Der Aufbau orientiert sich an der ursprünglichen Version von Jonas-Luetolf/Repetierer: Datei und Klassen als flache Navigation ohne Cards links, Repetition rechts. Die Kontextlabels sind klein und zurückhaltend. Klassenname neben dem Titel, Start/Freiwillig darunter, Schülername gross und linksbündig im Hauptbereich, Note unten links und Bewertungsaktionen unten rechts. Die bestehenden Theme-Farben bleiben erhalten. Neue Funktionen erhalten eine sekundäre visuelle Ebene.

Der gezogene Schüler ist der visuelle Fokus.

Ziel:

```text
REPETITION

ANNA MÜLLER

[Bewertung / Aktion]

Abbrechen    OK    Joker
```

Regeln:

- Schülername gross und prominent
- nicht nur als normales Label oben links
- sekundäre Controls zurückhaltend
- Note / Bewertung klar gruppieren
- Sessionstatus sichtbar, aber nicht dominant

---

# 15. Mitmachnote

Mitmachnote bleibt Teil von Repetierer.

Der globale Header bleibt:

```text
Repetierer
```

Im Inhalt:

```text
MITMACHNOTE

L23a · HS 2026
```

## Statistik-Karten

Geeignet für:

- Selbstbeurteilungen
- Lehrerbeurteilungen
- Versendet
- Synchronisation

Alle Karten gleich aufbauen.

## Aktionen

Primär:

- Antworten aktualisieren
- Beurteilungen öffnen

Sekundär:

- Einladung kopieren
- Klassensatz versenden

Destruktiv:

- Runde löschen

`Runde löschen` nicht gleichwertig zwischen normalen Aktionen platzieren.

Bevorzugt:

```text
⋯
Runde löschen
```

---

# 16. Alerts und Statusmeldungen

Eine gemeinsame Alert-Komponente verwenden.

Varianten:

- Info
- Success
- Warning
- Error

Alle verwenden:

- gleiche Struktur
- gleiche Abstände
- gleichen Radius
- konsistente Icons
- konsistente Statusfarben

Statusmeldungen im Drawer oder Footer sollen kleiner und ruhiger sein als Alerts im Hauptinhalt.

---

# 17. Einstellungen

Klare Sektionen:

```text
Einstellungen

Lokale KI
[...]

Design
[...]

Weitere Optionen
[...]
```

Footer:

```text
--------------------------------

Abbrechen    Speichern
```

Wenn die Einstellungen weiter wachsen, Kategorien verwenden:

- Allgemein
- Design
- KI
- Integrationen
- Daten
- Updates

---

# 18. KI-Assistent

Der Empty State muss bewusst gestaltet sein.

Nicht nur ein leeres Panel mit Eingabefeld anzeigen.

Beispiel:

```text
KI-Assistent

Frage etwas zur aktuellen Klasse,
Session oder Schülerauswahl.

[Session zusammenfassen]
[Wer wurde lange nicht aufgerufen?]
[Auffälligkeiten anzeigen]

Nachricht schreiben...
```

Prompt-Vorschläge dezent darstellen.

---

# 19. Icons

Icons:

- aus derselben Icon-Familie
- gleiche Strichstärke
- gleiche visuelle Grösse
- kein Mix aus Emoji, Filled und Outline Icons

Icons unterstützen die Orientierung, ersetzen aber keine klaren Labels bei wichtigen Aktionen.

---

# 20. Animationen

Animationen sind kurz und funktional.

Thinking Orbs verwenden die originale Canvas-Engine aus `thinking-orbs`, mit Zuständen passend zur tatsächlichen KI-Aktivität. GSAP steuert die App-Übergänge zentral: Dialoge und Drawer, Ansichtswechsel, Namensauswahl, Nachrichten und Fensteraktionen. Keine gestaffelten Einblendungen statischer Cards oder Listen. Ein Border Beam markiert den KI-Ladestatus. Diese Effekte bleiben lokal und pausieren bei Unsichtbarkeit beziehungsweise respektieren reduzierte Bewegung. Quellen und Lizenzen stehen in `docs/ui-sources.md`.

Verwenden für:

- Screenwechsel
- Modal öffnen/schliessen
- Drawer
- Schülerauswahl
- Statuswechsel
- Hover
- Focus

Keine langen oder verspielten Animationen im normalen Workflow.

Beispiel Schülerauswahl:

```text
Start
→ bestehende Controls treten leicht zurück
→ Schülername wird kurz mit GSAP eingeblendet
→ UI stabilisiert sich
```

---

# 21. Glow

Die normale Arbeitsoberfläche verwendet keine dekorativen Glows. Der Akzent entsteht durch aktive Tabs, ausgewählte Klassen und primäre Aktionen.

Erlaubt bei:

- Intro und die ausdrücklich aktivierten zusätzlichen visuellen Effekte

Nicht bei:

- jedem Button
- jedem Panel
- jedem Icon

---

# 22. Destruktive Aktionen

Destruktive Aktionen klar trennen.

Nicht mitten zwischen normalen Aktionen platzieren.

Beispiele:

- Runde löschen
- Session beenden
- Daten löschen

Bevorzugt:

- eigener Danger-Bereich
- Kontextmenü
- Danger-Button im Bestätigungsdialog

---

# 23. Bestehende gute Komponenten

Nicht alles neu gestalten.

Bestehende Komponenten, die bereits gut zur App passen, sollen primär vereinheitlicht werden.

Dazu gehören insbesondere:

- Participation Entry Modal
- Chancen / Wahrscheinlichkeiten
- Logs
- Update-News
- funktionale Editor-Screens

Dort hauptsächlich anpassen:

- Spacing
- Typografie
- Buttons
- Modal-Standard
- Farben
- Footer

---

# 24. Regeln für neue UI

Vor einem neuen Screen oder Modal:

1. Prüfen, ob bereits eine passende Komponente existiert.
2. Bestehende Komponenten wiederverwenden.
3. Diese DESIGN.md beachten.
4. Keine neue visuelle Sprache einführen.
5. Neue Komponenten nur erstellen, wenn wirklich nötig.
6. Wiederverwendbare Komponenten bevorzugen.
7. Globalen Header und bestehende Navigation nicht unnötig verändern.

---

# 25. Regeln für Codex

Bei UI-Aufgaben:

- zuerst diese `DESIGN.md` lesen
- nur den relevanten Screen und direkt abhängige Komponenten bearbeiten
- keine app-weiten Änderungen ohne expliziten Auftrag
- keine neue Komponentenfamilie einführen
- bestehende Styles und Komponenten bevorzugen
- Funktionalität nicht verändern, wenn nur Design verlangt ist
- vor neuen Komponenten prüfen, ob bereits etwas Passendes existiert
- nach der Änderung gegen diese Datei prüfen

---

# 26. Zielzustand

Repetierer soll sich anfühlen wie:

- eine zusammenhängende Desktop-App
- nicht wie eine Sammlung einzelner Features
- schnell verständlich
- ruhig und professionell
- modern, aber nicht verspielt
- visuell fokussiert
- konsistent

Grundsatz:

**Funktionalität vor Dekoration.  
Fokus vor Informationsmenge.  
Konsistenz vor Individualität einzelner Screens.**
