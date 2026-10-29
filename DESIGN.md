# Repetierer – Kristall

Die aktuelle Oberfläche folgt dem Kristall-Material aus [Design.md](docs/kristall-material.md), Version 1.3.0. Die frühere Designspezifikation bleibt in [docs/design-before-kristall.md](docs/design-before-kristall.md) erhalten.

## Projektvorgaben

- Kein Foto. Die lokale, abstrakte Salbei-Szene `public/kristall/assets/scene.svg` ist die gemeinsame Hintergrundtextur für CSS und WebGL.
- Ein optischer Canvas; echte semantische HTML-Controls liegen darüber. Repetition, Mitarbeit, Beurteilungen, Excel, lokale KI und ihre Datenflüsse bleiben erhalten.
- Systemschrift, weisse Schrift/Icons, Mint-Akzente, pillenförmige Controls und Panels mit Radius 30/32 px.
- Unveränderte Kristall-Shaderkonstanten: Brechung 1, Frost 0, Dichte .08. Ein Rand, ein Schatten; keine zusätzliche CSS-Glasfüllung im WebGL-Modus.
- Eine gemeinsame Federuhr für Marker, Toggles und Regler. Bestehende GSAP-Ansichtswechsel versorgen dieselbe optische Uhr für ihre begrenzte Dauer.
- Inaktive HTML-Inhalte unter Dialogen/Drawer ausblenden: Der Shader bricht nur die Szene und darf keine übereinander lesbaren Texte erzeugen.
- Canvas-Koordinaten an den tatsächlichen Bounds und den effektiven Pixelverhältnissen messen. Alle Overflow-Vorfahren beim Clippen berücksichtigen.
- Native Tastaturbedienung, Fokus, reduzierte Bewegung, Pointer-Abbruch, WebGL-Fallback und Lifecycle-Cleanup gehören zur Integration.

## Prüfung

`npm test`, `npm run test:ui`, `npm run test:kristall`. Die Materialprüfung erfasst Scrollbalken, DPR 1/1.25/2, Rasterzentren, unterbrochene Navigation, Regler, Schalter, kompakte/mobile Fenster und Kontextverlust. Screenshots und Messwerte liegen lokal in `.kristall-check/` bzw. `.ui-polish-smoke/`.

Alte Builds nie löschen. Neue Ausgaben in einem eigenen Unterordner unter `builds/` erstellen. Keine Veröffentlichung ohne Auftrag.

## Präzisierung: feste App-Chrome

Header als abgetrennter, fester Glasbereich; Drawer direkt darunter am linken Fensterrand, ohne schwebende Abstände. Header- und Drawer-Aktionen ohne Glaslinsen oder Pills: nur leichter Hover-Glow. Grösserer statischer App-Titel, Zahnrad und reines Drei-Sterne-Icon für die KI. Fenster mit 24-px-Glaskanten, kein Mauszeigerlicht. Der Drawer-Titel steht linksbündig mit den Kategorie-Icons.

## Chat und Session

Header unten gerade; weisse Hover-Glows und rote Fenster-/Session-Schliessen-Aktionen. Einstellungen drehen sich beim Hover. Drawer lässt die Hauptansicht sichtbar. Dialoge erhalten ein ungerahmtes X rechts oben, mit den bestehenden Abbrechen-Handlern.

KI-Chat ohne Vorschläge, Bereitschaftspunkt oder Fusszeile. Pfeil zum Senden im Eingabefeld, ohne optische Linse auf der Eingabe. Nachricht fliegt beim Senden in den Verlauf; Reduced Motion überspringt die Bewegung. Drei Punkte zeigen die laufende Anfrage.

Session-Navigation als festes schwebendes Glasdock 18 px über dem Fensterrand, ohne Unterstrich. Session beenden und Auswertung bleiben ungerahmte Textaktionen. Mitarbeit hat keine eigene Glasbox.

## Letzte Präzisierungen

Schülernamen bleiben Glas-Pills; Session beenden steht rechts oben. Session-Tabs sind breiter, Mitarbeit und Repetition verwenden grössere Titel. Auswahl und Verlauf-/Hilfeaktionen stehen in einer gleich hohen Toolbar. Der Notenbereich hat Innenabstand, der Noteninput nur eine Unterlinie. Hilfe-Fusszeilen entfallen. Die Speicherbestätigung verwendet das gemeinsame Kristallmaterial. Das Zahnrad dreht um seine optische Mitte.

Der Begrüssungs-Orb bleibt im leeren Chat erhalten. Ladezustand: drei Punkte. Sendepill vorab in Zielgrösse berechnen, beim Flug einblenden. Antworten sind ungerahmter Text und werden nach Eingang schrittweise ausgegeben; dies ist eine Darstellungsanimation, kein Netzwerk-Streaming.

## Aktuelle Detailkorrekturen und Build-Grenze

Keine Installer oder Windows-Builds automatisch erstellen; der Nutzer baut diese selbst. Nur Quellcode bearbeiten und lokal testen, sofern kein neuer ausdrücklicher Build-Auftrag vorliegt.

Zahnrad über Grid zentrieren; keine absolute SVG-Zentrierung oder Translation mit der Rotation kombinieren. Fenster mit 32-px-Glaskanten. Dock-Tabs etwas kompakter, Session beenden wieder als Pill und in der Repetitionsansicht oberhalb der Hauptbox. Repetitionstitel 32 px, näher am oberen Boxrand; Note 21 px, Fokus nur durch die Unterlinie.

Roulette zeigt ausschliesslich Vornamen; Personenidentität und Auswahl bleiben unverändert. Der Chat verbirgt die Zielnachricht über einen eigenen Flight-Zustand, bis der beschleunigte Flug endet. Allgemeine Einblendanimationen dürfen die Zielnachricht nicht vorzeitig zeigen. Nach dem Senden zum Verlaufende scrollen. Mitarbeit-Punkte als fettgedruckte Zahlen ohne Kreis, mit zurückhaltendem Hover-Glow; Beschreibungen in einem stabilen Grid daneben.

## Joker-Datumsspalten

Die Excel-Vorlage enthält ab Q sechs Spalten «Joker Datum 1–6»; weitere werden bei Bedarf ergänzt. Bestehende Dateien erhalten beim Öffnen eine optionale Migrationsvorschau mit Sicherungskopie. Eigene Spalten bleiben erhalten; bei belegtem Bereich wird der neue Datumsblock weiter rechts angelegt und die alte Sammelspalte als Archiv gekennzeichnet. Fehlende historische Daten nicht rekonstruieren oder erfinden. Migration, Joker-Speicherung, Undo/Redo, Warteschlange und Klassenlistenbearbeitung verwenden dieselbe Spaltenlogik. Keine Installer automatisch bauen.

## Offener Unterrichtsstart und Scrollbewegung

Klassenauswahl ohne eigene Karte: Glas-Pills in drei Spalten, ohne sichtbare Scrollleiste und ohne Datei-wechseln-Aktion. Abwesenheiten als eigener Schritt auf derselben Inhaltsfläche, kein Modal und keine Checkboxen. Native Buttons mit aria-pressed markieren Abwesenheiten durch leicht rote Glaslinsen; erneutes Aktivieren nimmt die Markierung zurück. Zurück führt zur Klassenauswahl, Session starten übernimmt die ausgewählten IDs. Während des Starts sind die Auswahl und die Zurück-Aktion gesperrt; bei Fehlern werden sie wieder freigegeben.

Mitarbeit ebenfalls mit drei Schüler-Pills pro Zeile; schmale Fenster reduzieren die Spaltenzahl. Native Scrollbereiche bleiben per Mausrad, Touchpad und Tastatur nutzbar. Ein gemeinsamer Scroll-Controller beobachtet Listen, Statistiken, Einstellungen, Editoren und Verlaufseinträge. Beim Wiedereintritt in den sichtbaren Bereich wird der dezente Fade/Versatz erneut abgespielt, in beiden Richtungen und beliebig oft. Reduced Motion überspringt die Bewegung. Glas und DOM werden über dieselbe optische Uhr aktualisiert. Keine dauerhafte Animationsschleife im Leerlauf.

## Korrektur des offenen Startablaufs

Keine Drei-Schritte-Anzeige auf Klassen- oder Abwesenheitsauswahl. Dateiinfo der Klassenauswahl dicht über dem unteren Fensterrand. Die intern gewählte Klasse erhält auf der Auswahlseite keine abweichende Darstellung; Tastaturfokus bleibt sichtbar.

Neue Inhaltsflächen und Abwesenheits-Pills müssen ausdrücklich `app-region: no-drag` verwenden. Sie dürfen den verschiebbaren Fensterbereich nicht erben. Hover, Maus-Klick, Mausrad und Tastatur sind direkt zu testen; programmatische `.click()`-Tests allein erkennen native Electron-Drag-Regionen nicht.

## Flüssige Seitenwechsel und Debug-Szenarien

Seitenwechsel bewegen den lebenden Hauptinhalt: Die alte Seite zieht sich in 140 ms leicht zusammen und bewegt sich abwärts hinaus. Die neue Seite kommt von oben mit 94 Prozent Grösse und federt über 460 ms auf ihre Endposition. Nur `#content` bewegt sich; Header und Szene bleiben stationär. Der gemeinsame Glas-Renderer folgt den tatsächlichen Bounds und der Deckkraft, einschliesslich skalierter Scroll-Clips. Schnelle Navigation löst die vorherige Animation ab; reduzierte Bewegung setzt den Zustand direkt. Bestehende Scroll-Effekte bleiben unabhängig nutzbar.

Einstellungen sind eine eigene Seite im Hauptinhalt, ohne Modal oder äussere Glaskarte. Titel, Zurück und Speichern bleiben sichtbar; die Optionen scrollen ohne sichtbare Leiste. Zurück/Verwerfen speichern keine Änderungen, Speichern nutzt dieselbe bestehende Validierung und Persistenz. Die vorherige Ansicht wird wiederhergestellt; aktive Sessions und Abwesenheitsauswahl bleiben erhalten. Hilfe und notwendige Bestätigungen liegen als Dialog über dieser Seite. Der Debug-Katalog führt Einstellungen unter Pages, nicht unter Modals.

Debug starten mit `npm run start:debug` oder `Repetierer.exe --debug`; `--debug-mode` und `--dev-mode` bleiben kompatibel. Im Drawer «Debug-Szenarien», rechts unten «Debug» oder Ctrl+Shift+D öffnet die durchsuchbare Übersicht. Sie enthält alle vorhandenen Modals und feste Beispielszenarien für Seiten, KI, Updates, Benachrichtigungen und Datenmigration. Jedes Szenario startet mit einem frischen Dokument. Während der Simulation werden alle ausgehenden IPC-Aktionen lokal abgefangen, ausser Fenster schliessen/minimieren. Keine Dateiänderungen, Downloads, Installationen oder Nachrichten. «Simulation beenden» verwirft das Szenario und lädt die normalen Daten neu. Ohne Debug-Flag gibt es keine Debug-Oberfläche und gespeicherte Szenarien werden nicht aktiviert. `npm run test:debug` prüft den echten App-Einstieg und die Szenarien in einem separaten temporären Benutzerverzeichnis.

## Aktuelle Bewegungs- und Einstellungsdetails

Session-Dock und Session beenden liegen als eigene Chrome neben dem bewegten Inhalt; ihre Position bleibt bei Seitenwechseln unverändert. Notenleiste und Aktionen poppen beim Start einer Auswahl hinein und beim Abschluss hinaus. Während des Roulettes bleibt die Notenleiste inert.

Modals öffnen mit einem kurzen weichen Pop; beim Schliessen verformen sie sich kurz und ziehen sich zusammen. Keine X-Buttons. Ein Klick neben den Inhalt schliesst über den bestehenden Abbrechen-Handler, damit der Dialogzustand aufgeräumt bleibt. Laufende Schreiboperationen werden durch das Ausblenden nicht abgebrochen. Tastatur und reduzierte Bewegung bleiben nutzbar.

Mitarbeitspunkte haben eine direkte Druckreaktion und eine kurze Bestätigung der ausgewählten Zahl. Mehrfachaktivierungen speichern einen Eintrag. Die Speicherbestätigung ist eine kompakte Glasplakette im unteren Repetitionsbereich. Sie poppt kurz auf, zeichnet das dezente Häkchen, zeigt «Gespeichert» und zieht sich zurück. Keine Teilchen und kein flächiges Overlay; kein dauerhafter Animations-Loop.

Scrollpositionen treiben ein deutliches Ein- und Herausfliegen an beiden Rändern, auch beim Umkehren der Richtung. An Listenanfang und -ende bleiben die ersten/letzten Einträge vollständig lesbar. Reduzierte Bewegung entfernt den Versatz und die Randblenden.

Zurück steht ohne Pfeil über Einstellungen. Keine Designauswahl und keine Spinnerpfeile an Zahlenfeldern. Ein gemeinsamer Import/Export über `.repetierer` überträgt SMTP und Cloudflare-Umfrage-API zusammen; Details in [configuration-package.md](docs/configuration-package.md). Exportierte Pakete enthalten portable Zugangsdaten, keine Schülerdaten. Die Quelle bleibt beim Paketimport erhalten.

Wahrscheinlichkeits-Faktor, Boost-Faktor und Wiggers-Dauer sind native Range-Regler mit dem vorhandenen Kristall-Griff und der gefüllten Glasspur. Ein Wert rechts neben der Beschriftung aktualisiert sich beim Ziehen und bei Tastatureingaben. Faktoren bleiben bei 1.1–10 in 0.1-Schritten, Minuten bei 1–480 in Einer-Schritten. Deaktivierte Optionen sperren auch den jeweiligen Regler. Speichern verwendet weiterhin dieselben numerischen Einstellungen. Keine zusätzlichen Karten oder Eingabefelder um die Werte.


## Mitmachnote und Chat-Reihenfolge

Die Mitmachnote ist eine offene Seite ohne äussere oder verschachtelte dunkle Karten. Übersicht, Runde und Schülerbeurteilung verwenden den gemeinsamen Seiten-Pop; bei einem Wechsel beginnt der Inhalt wieder oben. Der Header bleibt fest. Ein begrenzter, per Mausrad und Tastatur nutzbarer Inhaltsbereich verhindert abgeschnittene Beurteilungslisten. Mitarbeit-Auswertungen erhalten ebenfalls den Übergang.

Die KI-Anfrage startet direkt beim Senden. Die drei Ladepunkte bleiben während des Nachrichtenflugs verborgen und erscheinen erst nach der Landung. Auch schnelle Antworten warten auf diese Landung. Reduzierte Bewegung und Schliessen während des Flugs räumen den Flug auf und lassen keine wartende Anzeige hängen.

Scrollleisten sind überall in der App unsichtbar, inklusive Chat, Dialogen, Textfeldern und Auswertungen. Native Scrollbedienung bleibt erhalten. Gezielte Prüfungen: `node scripts/liquid-refinement-focused.cjs --assessment-polish` und `--ai-flight`.
