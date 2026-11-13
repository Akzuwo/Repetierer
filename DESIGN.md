# Repetierer – Kristall / Liquid Glass

Stand: 9. November 2026. Verbindliche Beschreibung der aktuellen App. Technische Umsetzung: [Architekture.md](Architekture.md), Arbeitsregeln: [AGENTS.md](AGENTS.md).

## Gestaltungsziel und Plattform

Eine ruhige Desktopoberfläche für Lehrpersonen. Eine dunkelgrüne, abstrakte Landschaft bildet die gemeinsame optische Szene. Durchsichtige Kristallflächen brechen diese Szene an weichen gerundeten Kanten. Klare Typografie, nachvollziehbare Zustände und wenige bewusst gesetzte Glasflächen tragen die Bedienung.

Die Windows-Electron-App hat ein festes Fenster von 1100 × 800 Pixeln. Desktopprüfungen dürfen zusätzlich 1440 × 900 verwenden. Kein mobiles App-Layout entwerfen oder prüfen. Die gesonderte Schülerumfrage ist responsiv; die Benachrichtigungskomponente besitzt ebenfalls schmale CSS-Abstände, ohne daraus einen mobilen Appmodus zu machen.

Keine Fotos, kein Mauszeigerleuchten, kein Titelglitzern. Es gibt ein sichtbares Design; keine Designauswahl in den Einstellungen. Historische Theme-Werte werden intern kompatibel gelesen.

## Material, Farbe und Schrift

Die Szene liegt in `public/kristall/assets/scene.svg`. Das gemeinsame WebGL-Material verwendet Brechung 1, Frost 0, Tönung 0.08, Helligkeit 1, Kontrast 0 und Fokus 0. Der Hintergrund bewegt sich nicht mit dem Mauszeiger. Transparente lokale Zeichenflächen für Dropdowns und Benachrichtigungen verwenden denselben Shader und dieselbe Szenenabbildung. Bei WebGL-Ausfall bleibt eine lesbare CSS-Glasdarstellung erhalten.

Die App verwendet die vorhandenen lokalen Schriftdateien und weisse Schrift mit abgestuften Deckkräften. Überschriften haben klare Grösse und Gewicht; Hilfstext bleibt zurückhaltend. Lichtkanten sind dünn und gerundet, Flächenschatten weich. Transparenz darf die Lesbarkeit nicht aufheben. Weiss ist die allgemeine Hoverfarbe; destruktive Aktionen verwenden Rot. Fokus bleibt erkennbar, ohne zusätzliche leuchtende Eingabeboxen.

Das Hauptfenster hat ausgeprägt runde, gläserne Aussenkanten mit 32 px Radius. Der abgetrennte Header besitzt unten eine gerade Kante. Die gerundete Fenstermaske bleibt davon unabhängig. Glas zeigt sich durch Brechung, Kantenlicht und Bewegung; zusätzliche kleine Boxen um jedes Icon sind nicht erwünscht.

## Fenster, Drawer und Navigation

Der Header bleibt bei Seitenwechseln stationär. Links stehen Menü und grösserer Appname, rechts Drei-Sterne-KI-Icon, Zahnrad und Fensteraktionen. Das Zahnrad dreht beim Hovern um seine eigene Mitte. Werkzeugicons besitzen keine Glasbubble; Hover zeigt Farbwechsel und weissen Glow. Schliessen der App wird beim Hovern rot mit rotem Glow.

Der Drawer ist am Fenster angebunden. Öffnen blendet den Hauptinhalt nicht aus. Titel gross, linksbündig auf der Iconachse; Hilfe daneben auf derselben Höhe und ausreichend gross. Kategorien sind keine Pills. Hover glüht nur leicht weiss. Dateiverwaltung ist ein Untermenü mit eigener weicher Übergangsbewegung; Zurück steht über dem Titel.

Während einer Session schweben Mitarbeit, bei aktivierter Klassenoption Hausaufgaben und Repetition gemeinsam im kompakten Dock kurz über dem unteren Fensterrand. Alle Ziele sind ausreichend breit, der Abstand klein. Kein Auswahlunterstrich. Session beenden ist eine Pill rechts oben, ausserhalb der Repetitionskarte, mit rotem Hover. Dock und Beenden werden wie der Header nicht von der Seitenanimation bewegt.

## Seiten und Eingaben

Klassenauswahl steht offen auf der Inhaltsfläche. Klassen erscheinen als etwa drei Glas-Pills pro Zeile; keine Klasse erhält einen dauerhaften Auswahlzustand. Dateiinfo steht nahe dem unteren Rand. Keine Drei-Schritte-Anzeige, kein Datei-wechseln-Button und keine zusätzliche äussere Karte.

Abwesenheiten sind eine eigene Seite ohne Modal und ohne Checkboxen. Schüler-Pills sind echte Buttons mit `aria-pressed`; abwesende Personen erhalten eine leichte rötliche Tönung. Scrollen und Rücknahme der Auswahl bleiben möglich. Mitarbeit zeigt ebenfalls etwa drei Schüler-Pills pro Zeile, ohne äussere eigene Box. Einträge bleiben in der Session nachvollziehbar. Beim Schliessen der App mit laufender Session erscheint derselbe Kristalldialog wie beim Sessionende: zurück zur App oder schliessen, ohne die Session zu beenden. Gespeicherte Daten und die fortsetzbare Session bleiben erhalten; ohne laufende Session schliesst die App direkt.

Repetition hat einen grossen Titel nahe dem oberen Kartenrand. Start, Freiwillig, Chancen, Undo, Redo und Hilfe bilden eine gleich hohe Werkzeugzeile. Im Roulette stehen nur Vornamen, ohne die Personenidentität zu ändern. Notenleiste und Aktionen poppen nach dem Start hinein; während des Roulettes ist die Eingabe gesperrt. Note ist deutlich lesbar. Der Noteninput hat nur eine Unterlinie, auch beim Fokus keine leuchtende Pill. Joker und Note haben genügend Abstand zum rechten Rand. Ein bestätigter Speichervorgang nutzt die gemeinsame Erfolgsbenachrichtigung; es gibt keine zusätzliche Checkmark-/Tropfen-/Receipt-Anzeige.

Hausaufgaben zeigt anwesende Schüler als dieselben offenen Auswahl-Pills wie Mitarbeit. Nach der Auswahl folgt dieselbe Repetitionskarte mit Werkzeugzeile, grossem Namen und der identischen Notenleiste mit Unterlinie, Abbrechen und Speichern. Nach bestätigtem Speichern kehrt die Ansicht zur Schülerauswahl zurück. Zurück erlaubt eine andere Person. Standard: aktiviert, Gewicht 1.2 und ein zusätzlicher Joker nach jeweils zwei benoteten Vorführungen. Die Klassenregler ändern Gewicht und Schwelle; Deaktivieren blendet den Dockeintrag aus und bewahrt vorhandene Noten. Beim erstmaligen Verwenden zeigt ein Dialog ohne Aufzählungspunkte die gemeinsame Excel-Erweiterung aller Klassen der Datei mit einer Sicherung. Bereits erweiterte Klassen und vorhandene Einträge bleiben erhalten; ein Klassenwechsel erzeugt keinen weiteren Erweiterungsdialog. Nach dem Speichern gelten dieselben Seiten- und Meldungsbewegungen wie bei Repetition.

Im Mitarbeit-Dialog stehen fette Zahlen ohne Kreise. Auswahl reagiert direkt mit Gewicht, kleinem Impuls und dezentem Glow. Eine Aktivierung speichert genau einen Eintrag, auch bei Doppelklick. Der Auswerten-Button besitzt keine zusätzliche Hülle.

Einstellungen sind eine eigene Seite. Ein unten mittig schwebendes Navdock wechselt zwischen „Global“ und „Klasse“ und verwendet denselben Kristallmarker wie das Session-Dock. Das Dock bleibt ausserhalb des animierten Inhalts. Der Wechsel Global/Klasse nutzt denselben gemeinsamen Seitenwechsel mit Aus- und Einfedern wie Mitarbeit/Repetition/Hausaufgaben; kein separater Fade. Global sind Darstellung, Benachrichtigungsdauer und Integrationen. Der Klassen-Dropdown steht oben fest über dem Scrollbereich. Beim Öffnen ist nur die Klasse einer aktiven Session vorausgewählt. Ohne aktive Session bleibt „Klasse auswählen“ stehen, auch wenn eine Datei oder Klasse geladen ist. Die Auswahl erschliesst die Regeln der geladenen Datei, ohne die aktive Session umzuschalten. Repetitions-Joker, Nachtragen, Wahrscheinlichkeitsfaktor, Boost, Wiggers-Regel und Hausaufgaben samt Parametern gelten nur für die ausgewählte Klasse. Die App speichert sie ausschliesslich lokal im Nutzerprofil; Excel enthält keine Klassenkonfiguration. Änderungen an beiden Bereichen werden gemeinsam gespeichert. Vor einem Klassenwechsel müssen Entwürfe gespeichert oder verworfen werden. Zurück steht ohne Pfeil über dem Titel. Wahrscheinlichkeit, Boost, Wiggers-Dauer und Benachrichtigungsdauer verwenden native Range-Slider mit Kristallregler und sofort sichtbarem Wert. Keine Zahlenfelder mit Spinbuttons. Import/Export der gemeinsamen `.repetierer`-Datei ist bei den Integrationen. Speichern und Verwerfen sind ausdrücklich getrennt.

Excel bearbeiten ist eine eigene Seite, auch ausserhalb einer Session. Im Kopf stehen Titel, Datei/Klassenanzahl und rechts die Klassenauswahl. Eine klare Tabelle verwendet Zeilennummer, Name, Joker, bestehende Noten und Löschen. Bearbeitbare Werte haben ruhige Linien, keine eigene Glaslinse pro Zelle. Neue Person und Speichern stehen in klar getrennten Aktionsbereichen. Ein Klassen- oder Seitenwechsel mit Entwurf verlangt eine sichtbare Entscheidung; fehlgeschlagenes Speichern bewahrt den Entwurf.

Statistiken zeigen geordnete Kennzahlen mit Typografie statt verschachtelten Karten. Durchschnittliche Repetitionsdauer erscheint nur aus tatsächlich gemessenen Einträgen; fehlende Altwerte sind keine Nullmessungen. Auswahl, Detailwerte und Scrollbereich folgen derselben Designsprache.

Mitmachnote nutzt eine eigene Seite mit Seitenanimation, klaren Kennzahlen, Rundenübersicht und funktionsfähigem Scrollbereich. Einladungslink kopieren übernimmt nur die URL in die Zwischenablage; Mailtext, Betreff und zusätzliche Kopierfelder entfallen. Die Klassenübersicht öffnet einen Schüler direkt. Noch nicht gespeicherte Lehrerbewertungen und Kommentare bleiben bei der Navigation als Entwurf erhalten. Selbst-/Lehrerbewertung verwendet einen nativen 1–5-Regler; ein unberührter Mittelwert gilt nicht als ausgefüllte Antwort. Kommentare pro Kriterium sind optional und standardmässig zugeklappt. Vorhandene Kommentare bleiben gespeichert.

Die Abgabefrist zeigt die zulässige Spanne (heute bis höchstens 90 Tage) direkt am Datumsfeld. Ungültige Fristen werden vor dem Anlegen einer Runde erklärt. Geschlossene Runden bieten eine neue Abgabefrist und «Runde wieder öffnen»; Umfragelink und Antworten bleiben erhalten. Noch nicht veröffentlichte Runden können ihre Frist korrigieren und denselben Entwurf erneut veröffentlichen. Die App synchronisiert veröffentlichte Runden der geladenen Exceldatei alle 30 Sekunden, auch ausserhalb der Umfrageansicht. Hintergrundaktualisierungen dürfen ungespeicherte Lehrerbewertungen, Eingaben und Tastaturfokus nicht ersetzen.

## Dropdowns

Native Selects bleiben Daten- und Formquelle, ein gemeinsamer Combobox-/Listbox-Controller liefert die Darstellung. Die geöffnete Liste ist rechtsbündig zum tatsächlich angeklickten Button, mit ausreichend Platz für den Pfeil. Beschriftungen oder leere Bereiche auf gleicher Höhe öffnen sie nicht.

Die Liste ist eine gerundete, klare Kristalllinse mit echter Szenenbrechung und weichem Auf-/Zupoppen. Das aktuell ausgewählte Element erhält weder einen andersfarbigen Hintergrund noch einen Tick. Tastaturfokus bleibt unterscheidbar. Escape, Pfeiltasten, Home/End, Eingabe und Suche per Tippen funktionieren. Behinderte Auswahlfelder lassen sich nicht öffnen.

## KI-Assistent

Neben «Lokale KI» in den Einstellungen steht ein freies Fragezeichen mit Tooltip: ohne Pill, Glasfläche oder Rahmen, mit leichtem weissem Glow beim Hover und erkennbarem Tastaturfokus. Sie öffnet jederzeit dieselbe Erklärung zu Einrichtung, lokaler Verarbeitung, Fähigkeiten und bestätigten Änderungen. Vor dem Einrichten öffnet sich diese Erklärung mit einer ausdrücklichen Downloadfreigabe. «KI-Integration entfernen» verlangt eine Bestätigung und deaktiviert den Zugriff in Repetierer dauerhaft bis zur erneuten Einrichtung; Ollama, Modelle und Fachdaten bleiben erhalten. Entfernen und Einrichten sind unmittelbare Integrationsaktionen, unabhängig vom Speichern anderer Einstellungen.

Im Header steht nur das Drei-Sterne-Icon. Der Begrüssungs-Orb bleibt im leeren Chat. Keine Chatvorschläge, kein grüner Bereit-Punkt, kein Hinweistext unter dem Composer und keine leuchtende Eingabeumrandung. Senden ist ein korrekt ausgerichteter Pfeil nach oben.

Die Nutzernachricht wird vor dem Flug in ihrer endgültigen Pillgrösse gemessen. Sie fadet während des beschleunigten Fluges ein. Am Ziel darf sie erst nach der Landung sichtbar sein. Nach dem Absenden scrollt der Verlauf ganz nach unten. Erst nach der Landung erscheinen drei Tipp-Punkte, falls die KI noch arbeitet.

KI-Antworten haben keine Pill und werden schrittweise als Text eingeblendet. Die Ausgabeanimation stellt bereits empfangenen Text dar. Änderungsbestätigungen bleiben konkrete Karten mit Aktionen; Fehler verwenden das gemeinsame Meldungssystem und bei Bedarf eine neutrale Erneut-versuchen-Aktion im Verlauf.

## Einheitliche Benachrichtigungen

Eine Meldung erscheint oben mittig unterhalb des Desktopheaders. Ein kurzer weicher Plopp bewegt sie von oben herein, streckt sie leicht und lässt sie sanft zur Ruhe kommen. Die gerundeten Kanten brechen dieselbe Szene wie die übrige App. Die Fläche ist gleichmässig getönt, ohne Farbverlauf pro Meldungstyp:

| Typ | Tönung | Inhalt |
| --- | --- | --- |
| Fehler | Rot | Fehlersymbol, kurzer Titel, konkreter Fehler und nächster sinnvoller Schritt |
| Warnung | Orange | Warnsymbol, Einschränkung oder noch ausstehende Aktion |
| Information | Grau | Infosymbol, sachlicher Hinweis |
| Erfolg | Grün | Häkchen, bestätigtes Ergebnis |

Das Meldungsicon sitzt unabhängig von der Anzahl Textzeilen vertikal in der Mitte der gesamten Meldungsbox.

Standarddauer 6 Sekunden, in den Einstellungen in ganzen Sekunden von 2 bis 8 veränderbar. Danach bewegt sich die Meldung sanft nach oben hinaus. Kein Schliessen-Button und keine zusätzlichen Aktionen. Meldungen fangen keine Klicks ab und nehmen keinen Fokus. Fehler/Warnung verwenden `role=alert`, Information/Erfolg `role=status`. Texte bleiben Plaintext.

Mehrere Meldungen stehen mit gleichmässigem Abstand untereinander. Höchstens drei sind sichtbar; neue Meldungen entfernen bei Überfüllung die ältesten. Ein gleicher fachlicher Schlüssel aktualisiert dieselbe Karte und startet den Timer neu, etwa für Downloadfortschritt. Eine gemeinsame transparente Zeichenfläche bedient den gesamten Stapel. Auf schmalen Bildschirmen reduziert nur die Komponente Breite, Rand und Zwischenraum.

Alle bisherigen transienten Meldungen laufen über `notify.error`, `notify.warning`, `notify.info` oder `notify.success`. Notwendige Entscheidungen wie Updatefreigabe, Migration und Löschbestätigung bleiben Dialoge; dauerhafte Ladeanzeigen, Feldhinweise und Statusdaten verschwinden nicht nach einem Meldungstimer. Eine ausstehende Exceländerung ist eine Warnung und keine erfolgreich geschriebene Note.

## Bewegung, Dialoge und Scrollen

Seitenwechsel bewegen nur `#content`: die alte Seite zieht sich leicht zusammen und sinkt hinaus; die neue kommt von oben verkleinert hinein und federt weich auf ihre Endposition. Header, Hintergrund und Session-Chrome bleiben fest. Schnelle Navigation ersetzt die laufende Bewegung sauber.

Dialoge poppen mit leichter weicher Dehnung auf und schliessen über einen kurzen Impuls und ein sauberes Zusammenziehen. Keine X-Buttons und keine zusätzliche Hilfe-Fusszeile mit Trennstrich. Klick neben den Dialog schliesst ihn, Escape bleibt bedienbar; notwendige laufende Schreib-/Migrationsaktionen können das Schliessen vorübergehend sperren. Fachliche Bestätigungsaktionen bleiben im Dialog.

Scrollbereiche zeigen keine sichtbare Scrollbar, behalten aber native Maus-, Touchpad- und Tastaturbedienung. Elemente fliegen positionsabhängig von unten hinein und oben hinaus; beim Hochscrollen umgekehrt. Die Bewegung wiederholt sich bei jedem Eintritt. Reduced Motion überspringt räumliche Bewegung und beendet bereits laufende Übergänge. Im Leerlauf gibt es keine permanente Animationsschleife.

## Appicon, Debug und Prüfung

Im Drawer stehen auch in der finalen App getrennt „Version 7.0.0“ und „Build 28.09112026“. Das Datum wird beim Bauen festgeschrieben. Hamburger und X verwenden dieselbe GSAP-Animation in beide Richtungen; das Kreuz wird aus den tatsächlichen Balkenmitten berechnet, auch bei schnellem Richtungswechsel und Reduced Motion.

Appicon ist das blaue R mit weissem Pfeil aus der gelieferten PNG. Die PNG bleibt mit 1254 × 1254 erhalten und besitzt transparente Ecken ausserhalb der blauen Rundung; das ICO enthält zehn 32-Bit-Grössen bis 256 × 256 mit echtem Alpha. Die vorherigen undurchsichtigen Originale liegen unter `design-reference/branding-before-transparency-2026-11-09/`. Alte Brandassets und ursprüngliche Designvorlagen liegen unverändert unter `design-reference/`.

Debug startet mit `--debug-mode` beziehungsweise `npm run start:debug` (ältere Flag-Aliase bleiben intern kompatibel). Bei der gepackten EXE darf `--debug` nicht direkt verwendet werden: Electron fängt es bereits vor dem App-Start als veraltete Node-Option ab. Das Debug-Menü gliedert sich in Meldungen (je ein Button für Nachricht, Warnung und Fehler), Updates und Testmail. Meldungen verändern weder Seite noch Daten. „Fake Update“ öffnet das Update-Modal als Vorschau. „Update mit Neustart“ simuliert Fortschritt und Installation und startet die App nach ausdrücklichem Klick wirklich neu; ein echter Download oder Installer wird nicht ausgeführt. Das Testmail-Formular sendet an genau eine eingegebene Adresse über die eingerichtete SMTP-Verbindung. Layout und Absender entsprechen der normalen Mitmachnoten-Mail, alle Inhalte sind erfundene Beispieldaten. Problem melden ist vorübergehend deaktiviert.

Passende Logik- und Desktop-UI-Prüfungen sind Pflicht. Materialänderungen zusätzlich gegen Kristall prüfen. Keine Installer automatisch erstellen; alte Builds niemals löschen. Prüfartefakte sind lokale Nachweise, keine neue Auslieferung.
