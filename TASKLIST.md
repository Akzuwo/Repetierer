# TASKLIST.md

## Zweck

Diese Datei enthält die noch anstehenden UI-/UX-Anpassungen für Repetierer.

Die Reihenfolge ist bewusst gewählt. Zuerst werden globale Grundlagen vereinheitlicht, danach zentrale Workflows und zuletzt Spezialdialoge.

Bei jeder Aufgabe gilt:

- `DESIGN.md` beachten
- Funktionalität nicht unnötig verändern
- keine unrelated Refactorings
- vorhandene Komponenten wiederverwenden
- Aufgabe möglichst lokal halten

---

# Phase 1 – Designsystem vereinheitlichen

## 1.1 Button-System

- [ ] Bestehende Buttons erfassen und auf vier Typen reduzieren:
  - Primary
  - Secondary
  - Ghost
  - Danger
- [ ] Einheitliche Höhen definieren
- [ ] Einheitliche Radien definieren
- [ ] Einheitliches horizontales Padding verwenden
- [ ] Hover-State vereinheitlichen
- [ ] Focus-State vereinheitlichen
- [ ] Disabled-State vereinheitlichen
- [ ] Danger-Buttons klar von normalen Aktionen unterscheiden
- [ ] Übermässige Pill-Formen entfernen
- [ ] Abweichende Button-Stile in alten und neuen Modals vereinheitlichen

## 1.2 Typografie

- [ ] Display-Stil definieren
- [ ] H1 definieren
- [ ] H2 definieren
- [ ] Body definieren
- [ ] Label definieren
- [ ] Muted definieren
- [ ] Eyebrow definieren
- [ ] Bestehende Screens auf konsistente Hierarchie prüfen
- [ ] Zu viele gleichzeitig konkurrierende Schriftgrössen reduzieren

## 1.3 Farben und Oberflächen

- [ ] Background-Ebene definieren
- [ ] Surface-Ebene definieren
- [ ] Elevated-Surface definieren
- [ ] Modal-Surface definieren
- [ ] Akzentblau zentralisieren
- [ ] Statusfarben vereinheitlichen
- [ ] Unnötige individuelle Farben aus Einzelkomponenten entfernen

## 1.4 Spacing

- [ ] Gemeinsames Spacing-System einführen
- [ ] Modal-Innenabstände vereinheitlichen
- [ ] Panel-Innenabstände vereinheitlichen
- [ ] Button-Abstände vereinheitlichen
- [ ] Header-Abstände vereinheitlichen
- [ ] Screens mit unruhigen oder zufälligen Abständen korrigieren

## 1.5 Icons

- [ ] Verwendete Icon-Sets prüfen
- [ ] Auf eine visuell konsistente Icon-Familie reduzieren
- [ ] Icon-Grössen vereinheitlichen
- [ ] Strichstärken vereinheitlichen
- [ ] Emoji oder stilistisch fremde Icons ersetzen

---

# Phase 2 – Globaler Header

## 2.1 Appname

- [ ] Auf allen Hauptscreens immer `Repetierer` als Appname verwenden
- [ ] `Mitmachnote` nicht mehr als alternativen Appnamen im Header verwenden

## 2.2 Autoren und Version

- [ ] Autoreninformationen aus dem Header entfernen
- [ ] Versionsinformationen aus dem Header entfernen
- [ ] Einen passenden `Über Repetierer`-Bereich festlegen
- [ ] Optional Version und Autoren dezent unten im Drawer anzeigen

## 2.3 Header-Konsistenz

- [ ] Headerhöhe überall angleichen
- [ ] Linke und rechte Innenabstände angleichen
- [ ] KI-Button/Icon vereinheitlichen
- [ ] Einstellungen-Button/Icon vereinheitlichen
- [ ] Fenstersteuerung visuell sauber ausrichten
- [ ] Header auf allen Screens gegen DESIGN.md prüfen

---

# Phase 3 – Startseite / Klassenauswahl

## 3.1 Leeren Zustand neu strukturieren

- [ ] Ausgegraute Repetitionsoberfläche entfernen, solange keine Klasse gewählt ist
- [ ] Klassenauswahl als alleinigen Fokus darstellen
- [ ] Eyebrow `UNTERRICHTS-SESSION` verwenden
- [ ] Titel `Klasse wählen` klar hervorheben
- [ ] Hilfstext ergänzen/vereinheitlichen
- [ ] Klassenoptionen als konsistente Auswahlkomponenten darstellen
- [ ] Leere Flächen ausgewogen nutzen
- [ ] Inhalt nicht unnötig am oberen Rand kleben lassen

## 3.2 Zuletzt verwendete Klasse

- [ ] Optional `Zuletzt verwendet` dezent anzeigen
- [ ] Sicherstellen, dass diese Information nicht mit der Hauptauswahl konkurriert

---

# Phase 4 – Session-Flow

## 4.1 Zustände klar trennen

- [ ] Klasse wählen
- [ ] Abwesenheiten erfassen
- [ ] Session starten
- [ ] Aktive Session
- [ ] Session beenden

Jeder Zustand soll visuell klar vom nächsten getrennt sein.

## 4.2 Abwesenheiten

- [ ] Abwesenheitsauswahl gegen neues Designsystem prüfen
- [ ] Auswahlzustände klar darstellen
- [ ] Primäraktion `Session starten` deutlich machen
- [ ] Sekundäraktionen zurückhaltender gestalten

## 4.3 Sessionstatus

- [ ] Aktive Klasse sichtbar, aber nicht dominant anzeigen
- [ ] Sessionstatus konsistent platzieren
- [ ] Unnötige technische Informationen aus dem Hauptfokus entfernen

---

# Phase 5 – Aktive Repetition

## 5.1 Schüler als Fokus

- [ ] Aktuell ausgewählten Schüler deutlich grösser darstellen
- [ ] Name aus der bisherigen eher beiläufigen Position lösen
- [ ] Schülername als zentrales Display-Element behandeln
- [ ] Genügend Abstand rund um den Namen schaffen

## 5.2 Bewertung

- [ ] Note / Bewertung klar unter dem Schüler gruppieren
- [ ] Eingabe- und Auswahlbuttons vereinheitlichen
- [ ] Primäraktion klar hervorheben
- [ ] Abbrechen / Joker / weitere Aktionen visuell unterordnen

## 5.3 Übergang

- [ ] Kurze Animation beim Ziehen eines neuen Schülers
- [ ] Optional sehr kurzer Akzent-Glow
- [ ] Keine übertriebene oder lange Animation

---

# Phase 6 – Drawer

## 6.1 Navigation

- [ ] Grosse Button-Pills durch kompaktere Navigation ersetzen
- [ ] Aktiven Bereich deutlich hervorheben
- [ ] Inaktive Einträge ruhiger darstellen
- [ ] Icons vereinheitlichen
- [ ] Titel `Dateien` prüfen und ggf. weniger dominant gestalten

## 6.2 Statusbereich

- [ ] Meldungen wie `Datei geladen...` als dezenten Statusbereich gestalten
- [ ] Hilfe-Button sauber integrieren
- [ ] Version / Autoren optional unten integrieren

---

# Phase 7 – Modals vereinheitlichen

## 7.1 Gemeinsamer Modal-Standard

- [ ] Gemeinsame Modal-Komponente bzw. gemeinsame Layout-Regeln verwenden
- [ ] Titelposition vereinheitlichen
- [ ] Beschreibungsabstand vereinheitlichen
- [ ] Footer vereinheitlichen
- [ ] Button-Reihenfolge vereinheitlichen
- [ ] Overlay und Blur vereinheitlichen

## 7.2 Session-Ende-Modal

- [ ] Bestehenden Stil vollständig an Standardmodal angleichen
- [ ] Grosses Warnsymbol entfernen oder deutlich reduzieren
- [ ] Titel linksbündig
- [ ] Erklärung kürzer und klarer formulieren
- [ ] `Weiter unterrichten` durch `Abbrechen` ersetzen
- [ ] `Session beenden` als Danger-Aktion darstellen
- [ ] Modal weniger dramatisch gestalten

## 7.3 Participation Entry Modal

- [ ] Bestehendes Grunddesign beibehalten
- [ ] Buttons ans neue System anpassen
- [ ] Spacing angleichen
- [ ] Typografie angleichen
- [ ] Footer angleichen

## 7.4 Chancen / Wahrscheinlichkeiten

- [ ] Nur Designsystem-Anpassungen
- [ ] Keine funktionale Neugestaltung
- [ ] Buttons, Spacing und Typografie prüfen

## 7.5 Logs

- [ ] Modal-Layout angleichen
- [ ] Header/Footer angleichen
- [ ] Controls vereinheitlichen

## 7.6 Update-News

- [ ] Modal-Layout angleichen
- [ ] Buttons vereinheitlichen
- [ ] Typografie und Spacing vereinheitlichen

---

# Phase 8 – Mitmachnote

## 8.1 Globale Einordnung

- [ ] Globalen Appheader `Repetierer` verwenden
- [ ] `MITMACHNOTE` als Eyebrow im Inhalt verwenden
- [ ] Klasse und Semester sauber als Kontext anzeigen

## 8.2 Statistik-Karten

- [ ] Karten gleich gross machen
- [ ] Innenabstände angleichen
- [ ] Zahlen / Statusinformationen gleich hierarchisieren
- [ ] Karten nicht unnötig gross gestalten

## 8.3 Aktionen

- [ ] `Antworten aktualisieren` als Primäraktion behandeln
- [ ] `Beurteilungen öffnen` als Primär-/Hauptaktion behandeln
- [ ] `Einladung kopieren` sekundär
- [ ] `Klassensatz versenden` sekundär
- [ ] `Runde löschen` aus normaler Aktionsgruppe entfernen
- [ ] `Runde löschen` in Kontextmenü oder klaren Danger-Bereich verschieben

## 8.4 Alerts

- [ ] Warnung `Umfrage noch nicht eingerichtet` an gemeinsame Alert-Komponente anpassen
- [ ] Success-/Error-/Info-Zustände derselben Komponente verwenden

---

# Phase 9 – KI-Assistent

## 9.1 Empty State

- [ ] Leeren Zustand bewusst gestalten
- [ ] Titel und kurze Erklärung ergänzen
- [ ] Sinnvolle Prompt-Vorschläge anzeigen
- [ ] Eingabefeld klar am unteren Bereich positionieren
- [ ] Leere Fläche besser nutzen

## 9.2 Prompt-Vorschläge

- [ ] Visuell zurückhaltender gestalten
- [ ] Bestehende Button-/Chip-Komponenten verwenden
- [ ] Keine zusätzliche neue Designsprache einführen

---

# Phase 10 – Einstellungen

## 10.1 Struktur

- [ ] Sektionen visuell vereinheitlichen
- [ ] `Lokale KI`, `Design`, weitere Optionen klar gruppieren
- [ ] Abstände zwischen Sektionen harmonisieren

## 10.2 Footer

- [ ] Fester Modal-/Screen-Footer
- [ ] `Abbrechen`
- [ ] `Speichern` als Primary

## 10.3 Skalierung für weitere Einstellungen

- [ ] Prüfen, ob Kategorienavigation bereits nötig ist
- [ ] Bei weiterem Wachstum vorbereiten:
  - Allgemein
  - Design
  - KI
  - Integrationen
  - Daten
  - Updates

---

# Phase 11 – Alerts und Statuskomponenten

- [ ] Eine gemeinsame Alert-Komponente verwenden
- [ ] Info-State
- [ ] Success-State
- [ ] Warning-State
- [ ] Error-State
- [ ] Icons vereinheitlichen
- [ ] Farben vereinheitlichen
- [ ] Innenabstände vereinheitlichen
- [ ] Alle bestehenden Alerts migrieren

---

# Phase 12 – Eingabefelder

- [ ] Input-Höhen vereinheitlichen
- [ ] Radien vereinheitlichen
- [ ] Background vereinheitlichen
- [ ] Focus-State vereinheitlichen
- [ ] Labels vereinheitlichen
- [ ] Placeholder-Styling vereinheitlichen
- [ ] Disabled-State vereinheitlichen

---

# Phase 13 – Spezial- und Editor-Screens

## 13.1 Daten-/Excel-Editoren

- [ ] Funktionale Dichte beibehalten
- [ ] Header und Toolbar angleichen
- [ ] Buttons ans gemeinsame System anpassen
- [ ] Inputs angleichen
- [ ] Tabellenaktionen visuell ordnen
- [ ] Keine unnötige Vereinfachung funktionaler Werkzeuge

## 13.2 Weitere Spezialdialoge

- [ ] Alle verbleibenden Modals gegen Standardmodal prüfen
- [ ] Alle verbleibenden Buttons gegen Button-System prüfen
- [ ] Alle verbleibenden Alerts migrieren

---

# Phase 14 – Motion und Feinschliff

## 14.1 Screenwechsel

- [ ] Einheitliche kurze Transition
- [ ] Keine langen Animationen

## 14.2 Modals

- [ ] Einheitliche Open-/Close-Animation

## 14.3 Drawer

- [ ] Saubere Open-/Close-Animation

## 14.4 Schülerauswahl

- [ ] Kurzer Übergang beim Wechsel
- [ ] Optional dezenter Glow

## 14.5 Hover und Focus

- [ ] Buttons
- [ ] Inputs
- [ ] Navigation
- [ ] Cards, falls interaktiv

---

# Phase 15 – Abschluss-Audit

Nach Umsetzung aller Punkte:

- [ ] Screenshot aller Hauptscreens erstellen
- [ ] Screenshot aller Modals erstellen
- [ ] Screens gegen `DESIGN.md` prüfen
- [ ] Header überall identisch?
- [ ] Typografie konsistent?
- [ ] Buttons konsistent?
- [ ] Modals konsistent?
- [ ] Alerts konsistent?
- [ ] Drawer konsistent?
- [ ] Danger-Aktionen sauber getrennt?
- [ ] Appname überall korrekt?
- [ ] Autoren nicht mehr im Header?
- [ ] Version sinnvoll platziert?
- [ ] Keine ausgegrauten zukünftigen Zustände sichtbar?
- [ ] Aktiver Schüler klarer Fokus?
- [ ] Keine neuen UI-Muster ohne Grund?
- [ ] Alle bestehenden Funktionen weiterhin nutzbar?

---

# Priorität

## P0 – zuerst

- [ ] Designsystem
- [ ] Globaler Header
- [ ] Startseite / Klassenauswahl
- [ ] Aktive Repetition
- [ ] Session-Ende-Modal

## P1

- [ ] Drawer
- [ ] restliche Modals
- [ ] Mitmachnote
- [ ] Alerts

## P2

- [ ] KI-Assistent
- [ ] Einstellungen
- [ ] Inputs
- [ ] Spezial-/Editor-Screens

## P3 – Feinschliff

- [ ] Animationen
- [ ] Hover-/Focus-Details
- [ ] Abschluss-Audit
