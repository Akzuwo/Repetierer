# Repetierer – Architektur

Stand: 9. November 2026. Diese Datei beschreibt den aktuellen Quellcode. Gestaltung: [DESIGN.md](DESIGN.md); Arbeitsregeln: [AGENTS.md](AGENTS.md).

## Laufzeit und Grenzen

Die Lehrpersonen-App läuft als Windows-Electron-Anwendung mit einem rahmenlosen, transparenten Desktopfenster von 1100 × 800 Pixeln. `app.js` startet den separaten Splash, das Hauptfenster und den zustimmungsbasierten Updater. Entwicklerwerkzeuge und das Meldungs-Debug-Menü sind an die Debug-Flags gebunden. Fensterpräferenzen verwenden gegenwärtig `nodeIntegration: true` und `contextIsolation: false`: Renderer sind privilegierter Anwendungscode, kein Ort für ungeprüftes Fremd-HTML.

Die Schülerumfrage unter `cloudflare/` ist eine getrennte Browseranwendung mit Worker/API und D1. Ihre mobile Bedienung ist kein Grund, die Electron-App für Smartphones umzubauen oder dort zu testen. Lokale UI-Arbeit veröffentlicht weder Cloudflare-Dienste noch Releases.

## Module und Datenfluss

| Bereich | Dateien | Verantwortung |
| --- | --- | --- |
| Einstieg/Fenster | `app.js`, `public/splash.*` | Splash, Hauptfenster, Fensteraktionen, Updateereignisse |
| Hauptprozess | `src/mainProcess.js` | IPC, Dateiauswahl, Session- und Excelaktionen, Protokoll, Backup, Undo/Redo, Integration der Fachmodule |
| Hausaufgaben | `src/homework.js`, `public/homeworkPage.js` | Freiwillige Auswahl, Note/Datum, unveränderliches Notengewicht, Joker, Migration und Undo/Redo |
| Einstellungen | `src/storage.js`, `public/settingsScope.js` | Globale Optionen und lokale Regeln pro Datei/Klasse |
| Repetition | `src/program.js`, `src/repetitionTimer.js` | Personenauswahl, Gewichtung, Joker/Wiggers-Regel, Noten, Auswahl-Token und Zeitmessung |
| Excel | `src/excelReader.js`, `src/jokerDates.js` | Arbeitsmappen lesen/schreiben, Klassen, Datumsspalten, Migrationsplan und Sicherung |
| Lokale Daten | `src/storage.js`, `src/teachingSessions.js` | Einstellungen, Backup/Warteschlange, Protokoll, Unterrichtssessions |
| Statistik | `src/classStats.js` | Klassenwerte und durchschnittliche Repetitionsdauer |
| Mitmachnote | `src/assessment/*` | Runden, Antworten, Zuordnung, Lehrerbewertung, API, Versand und Konfiguration |
| Lokale KI | `src/ai/*` | Ollama, begrenzte Werkzeuge, Validierung, Bestätigung vor Schreibaktionen |
| Seiten | `public/renderProcess.js`, `assessmentRenderer.js`, `excelEditorPage.js`, `aiRenderer.js` | Lebende DOM-Ansichten, Zustände, Benutzeraktionen und IPC-Ergebnisse |
| Darstellung | `public/kristall/*`, `uiMotion.js`, `scrollEffects.js`, `customSelect.js` | Kristall-Shader, native Eingaben, Seiten/Dialogbewegung, Scrollbewegung, Dropdowns |
| Meldungen | `src/notificationPolicy.js`, `src/notifications.js`, `public/notifications.js` | Normalisierung, Dauer, Hauptprozesszustellung und eine gemeinsame sichtbare Anzeige |
| Debug | `src/debugFlags.js`, `src/debugUpdate.js`, `public/debugMode.js`, `src/assessment/debugMail.js` | Meldungen, Updatevorschau, Neustartsimulation und SMTP-Testmail nur mit Debug-Flag |

Eine typische Schreibaktion läuft von einem nativen Button über IPC und das Fachmodul zum jeweiligen Speicher. Erst das bestätigte Ergebnis verändert den sichtbaren Abschlusszustand. Ein lokal gesicherter, noch nicht nach Excel geschriebener Eintrag erzeugt eine Warnung; er darf keine erfolgreiche Excel-Speicherung vortäuschen. Der Exceleditor bearbeitet eine explizit angegebene Klasse mit Anfrage-ID und setzt dafür keine aktive Repetitionsklasse voraus.

## Persistenz und Migration

`storage.getPaths()` legt die Daten unter Electrons `userData` ab:

| Datei | Inhalt |
| --- | --- |
| `settings.json` | Appoptionen, zuletzt verwendete Daten und unter anderem `notificationDurationSeconds` |
| `backup.json` | Änderungsverlauf und ausstehende Exceländerungen |
| `migrations.json` | Migrationsstatus |
| `repetierer.log` | Anwendungsprotokoll |
| `unterrichtssessions.json` | Anwesenheitsschnappschüsse, Mitarbeit und Repetitionseinträge |
| `mitmachnoten.json` | Beurteilungsrunden, Antworten, Bewertungen und Versandstatus |
| `repetierer.env`, altes `brevo-smtp.env` | Lokale Integrationskonfiguration; sensible Werte über den Betriebssystem-Schlüsselspeicher |

Die eigentliche Exceldatei bleibt am gewählten Pfad. `Excel Template.xlsx` enthält auf den Klassenblättern sechs Joker-Datumsspalten ab Q; die gemeinsame Spaltenlogik erweitert den Block bei Bedarf ohne eigene Spalten zu überschreiben. Alte Dateien erhalten eine ausdrückliche Migrationsvorschau mit optionalem Start und Sicherungskopie. Geänderte oder gesperrte Quellen werden nicht blind überschrieben. Nicht vorhandene historische Joker-Daten werden nicht erfunden.

Repetitionsdauer beginnt nach dem Roulette beim bereitstehenden Namen. Sie wird vor dem Excel-I/O eingefroren und mit dem Verlauf gespeichert. Der Durchschnitt berücksichtigt passende, akzeptierte Noteneinträge mit bekannter Dauer; Altbestände ohne Zeitwert werden nicht als null Sekunden gerechnet.

## Benachrichtigungen

Renderer initialisieren die gemeinsame Anzeige einmal in `renderProcess.js`. Andere Module importieren dieselbe Fassade:

```js
const notify = require('./notifications'); // innerhalb public/
notify.error('Die Datei konnte nicht gespeichert werden.');
notify.warning('Excel ist geöffnet.', {title: 'Speichern ausstehend'});
notify.info('Der Export wird vorbereitet.', {key: 'export'});
notify.success('Die Datei wurde gespeichert.', {title: 'Export', key: 'export'});
```

Im Renderer sind auch `window.notify` und `window.Notifications` verfügbar; `notify.Error(text)` ist ein Alias. Der native JavaScript-Konstruktor `Error` bleibt unverändert. Ein gleicher `key` aktualisiert eine vorhandene Meldung und startet ihre Dauer neu. Leerer Text wird ignoriert; Texte werden über `textContent` ausgegeben. Die Policy begrenzt die Dauer auf 2–8 Sekunden, Standard 6, sowie die Anzeige auf drei Meldungen.

Der Hauptprozess verwendet `require('./src/notifications')` ab dem App-Einstieg beziehungsweise `require('./notifications')` innerhalb `src/`. `app.js` registriert das Hauptfenster vor `loadFile`. Meldungen vor dem Laden werden begrenzt gepuffert und über `ui:notification` zugestellt. `options.target` kann ein konkretes Fenster/WebContents angeben. Die gespeicherte Dauer wird beim Erzeugen übernommen; Settingsereignisse aktualisieren den Renderer.

Die Anzeige ist ein nicht blockierendes manuelles Popover der obersten Ebene. Eine einzige lokale transparente Zeichenfläche zeichnet bis zu drei farbige Linsen mit demselben Shader wie die App. Es gibt keinen Kontext pro Meldung und keine dauerhafte Animationsschleife. Beim WebGL-Ausfall greift das CSS-Material. Timer, GSAP-Bewegung, Entfernen alter Karten und Lebenszyklus werden zentral verwaltet. Reduzierte Bewegung beendet laufende Bewegungen sauber.

`showUpdatePanel` ist ein Kompatibilitätsadapter für ältere Aufrufer: passive Meldungen gehen an das neue System; eine Updatefreigabe bleibt ein Aktionsdialog. Fachliche Entscheidungen, KI-Änderungsbestätigungen, Migrationsvorschauen sowie dauerhafte Lade-/Konfigurationszustände sind keine wegzutimernden Benachrichtigungen.

## Integrationen und Sicherheit

`src/assessment/deadlinePolicy.js` prüft echte Kalendertage und das Tagesende in Europe/Zurich gegen die API-Grenze von 90 Tagen, bevor eine Runde lokal angelegt wird. `assessment:reopen-round` veröffentlicht einen korrigierten Entwurf oder ruft den besitzgeschützten Worker-Endpunkt `POST /api/teacher/rounds/:id/reopen` mit einer neuen Frist auf. Erst der bestätigte API-Erfolg ändert Frist und Status lokal. Wiederöffnen behält Antworten und Duplikatsperren; die Aufbewahrung wird ab der ausdrücklichen Wiederöffnung auf 90 Tage gesetzt. Bereits endgültig gelöschte Online-Runden können nicht wieder geöffnet werden. Für diesen Endpunkt ist die Veröffentlichung des geänderten Workers nötig; ein App-Build veröffentlicht die API nicht.

`src/assessment/autoSync.js` startet im Hauptprozess einen 30-Sekunden-Timer für veröffentlichte offene und geschlossene Runden der aktuellen Exceldatei. Ausstehende Entwürfe werden nicht ungefragt veröffentlicht. Manuelle Synchronisierung, Schliessen, Wiederöffnen und Löschen teilen eine Warteschlange pro Datei/Klasse/Runde. Langsame Durchläufe überlappen nicht; nach Fehlern erfolgt ein neuer Versuch im nächsten Takt. `assessment:auto-synced` aktualisiert die sichtbare Übersicht, ohne Lehrerentwürfe oder fokussierte Eingaben zu ersetzen. Der Timer endet beim Beenden der App.

`src/ai/integration.js` verbindet den Ollama-Dienst mit dem persistierten Schalter `aiIntegrationEnabled` (Altbestände bleiben kompatibel aktiv). Entfernen deaktiviert Statuszugriffe, Chat und Schreibbestätigungen und verwirft ausstehende Bestätigungen, ohne Ollama oder Modelle zu deinstallieren. Die neue Einrichtung setzt den Schalter erst nach Erfolg zurück. Die Erklärung `ai-info-modal` wird vor dem Einrichten und über die Einstellungs-Hilfe verwendet. `ai:setup` verlangt eine bestätigte Freigabe. Andere Einstellungen überschreiben diesen Integrationsschalter nicht.

Cloudflare-API-Zugang und Brevo SMTP können gemeinsam als `.repetierer` exportiert/importiert werden. Das portable Paket enthält Zugangsdaten; Quellen und bestehende Konfiguration werden bei fehlgeschlagener Prüfung erhalten. Lokale Zugangsdaten nutzen `safeStorage`. Absender von Resultatmails ist fest `Philosophie Mündlich beurteilung <no-reply@akzuwo.ch>`; SMTP-Anmeldung und sichtbare Absenderidentität sind getrennt.

Die KI arbeitet lokal über Ollama. Werkzeuge erhalten geprüfte Argumente und begrenzten Kontext; Schreibaktionen benötigen eine konkrete Änderungsbestätigung. Die Zeichenanimation einer bereits erhaltenen KI-Antwort ist kein Netzwerkstreaming.

## Prüfung und Lieferung

`npm run test:integrations` prüft Datumsgrenzen, Sync-Warteschlange und KI-Sperre sowie native Desktopaktionen für Erklärung, Entfernen, erneute Einrichtung und Wiederöffnen. Der lokale Cloudflare-E2E-Test prüft zusätzlich eine echte automatische Übernahme nach 30 Sekunden und eine verspätete Schülerabgabe nach der Wiederöffnung über echte IPC, Worker und lokales D1. Er verwendet ausschliesslich synthetische Daten und veröffentlicht nichts.

`npm run test:notifications` prüft Policy, Persistenz, Hauptprozesspuffer und die echte Electron-Anzeige. `npm run test:ui`, `npm run test:debug` und `npm run test:kristall` prüfen Desktopabläufe, Simulation und Optik. Für eine Appänderung ohne Browserumfrage werden die Node-Tests ohne `test/cloudflareE2E.test.js` ausgeführt; die getrennte Umfrage hat die privaten lokalen Backendtests und eine private Worker-Buildprüfung.

UI-Harnesses verwenden temporäre Benutzerdaten. `REPETIERER_DEV_USER_DATA` erlaubt auch beim Entwicklungsstart ein getrenntes Verzeichnis. Für Electron-Prüfungen muss `ELECTRON_RUN_AS_NODE` entfernt sein. Renderer-/IPC-Fehler und reale Eingaben prüfen, nicht allein programmatische Klicks.

Das Debug-Menü verwendet für seine drei Meldungsbuttons direkt `public/notifications.js`; es ersetzt keine IPC-Methoden und spielt keine Demo-Arbeitsmappe ein. Zusätzlich führt `debug-update-show` zur Updatevorschau oder zur vollständigen Simulation. `src/debugUpdate.js` verwaltet Zustimmung, Abbruch, Fortschritt und Doppelklicks; der reale Updater wird umgangen. Ein separater Marker `debug-update-state.json` im Nutzerprofil überlebt `app.relaunch()` / `app.exit(0)`. Nach dem Start wird er konsumiert, der neue Prozess meldet Erfolg und schreibt `debug-update-completed.json` mit alter/neuer PID. Der produktive `update-state.json` bleibt unberührt.

`debug:send-test-mail` prüft das Debug-Flag auch im Hauptprozess. Es baut aus erfundenen Daten mit `buildResultMessage` die produktive Text-/HTML-Mail und sendet sie über den bestehenden SMTP-Adapter an genau eine validierte Adresse. Weder Arbeitsmappe noch Beurteilungs-/Versandstatus werden verändert. Der UI-Test fängt Nodemailer lokal ab und prüft Erfolg, Authentifizierungsfehler und die Sperre ohne Debug-Flag; keine echte Mail wird zu Testzwecken versendet. Alte gespeicherte Szenarioschlüssel werden weiterhin entfernt.

`npm run test:debug:exe` kopiert die vorhandene Windows-Laufzeit aus `builds/win-unpacked` in ein temporäres Verzeichnis, integriert den aktuellen Quellstand und setzt vor der Initialisierung ein temporäres Nutzerprofil. Es erstellt keinen Installer und verändert keine vorhandenen Builds. Native Mausklicks wählen Meldungen und Update-Simulation. Nach dem Neustart wird eine neue native PID im selben temporären Profil nachgewiesen. Auch Hamburger/X in beide Richtungen, Unterbrechungen, Reduced Motion sowie getrennte Version/Build-Anzeige werden geprüft. Weitere Kriterien sind die sichtbare Meldung innerhalb der Fenstergrenzen, unverdeckter Hauptinhalt und eine ausgeblendete Repetition ohne Session. `node scripts/debug-exe-check.cjs --baseline` prüft stattdessen den alten gepackten App-Code. Die EXE verwendet `--debug-mode`, da `--debug` vor dem App-Start von Electron abgefangen wird. `REPETIERER_DEBUG_VISIBLE=1` aktiviert auch beim regulären Debug-Test ein sichtbares Fenster.

Alte Builds werden nie gelöscht. Installer und Veröffentlichung erfolgen ausschliesslich auf ausdrücklichen Auftrag. Eine Quellcodeprüfung ist kein Nachweis eines neu gebauten Installers oder eines echten Mailversands.

`ai-info-btn` wird vom Kristall-Surface-Controller ausdrücklich ausgenommen. Die CSS-Darstellung enthält nur das Fragezeichen mit Hover-Glow und Tastaturfokus. Die Meldungsicons werden mit `align-items:center` und ohne oberen Versatz in der gesamten Karte zentriert. `scripts/export-app-icon.cjs` exportiert die freigestellte PNG in zehn PNG-basierte 32-Bit-ICO-Grössen; die Brandingtests dekodieren die Alpha-Pixel jeder Grösse. Originale und Bearbeitungsnotiz bleiben unter `design-reference/branding-before-transparency-2026-11-09/`.

## Öffentlicher Repository-Umfang

Versioniert werden App, appseitiger API-Client, Dokumentation und lokale App-Prüfungen. `cloudflare/`, `scripts/cloudflare-*`, `test/cloudflareE2E.test.js` und die Deployment-Anleitung bleiben lokal und stehen in `.gitignore`. Die Cloudflare-npm-Aliase wurden aus dem öffentlichen Paket entfernt; ihre bisherige Belegung bleibt lokal unter `.cloudflare-private/` erhalten. Der Server wird unabhängig verwaltet. `npm run test:repository` prüft den Git-Index; `.gitleaks.toml` ergänzt einen vollständigen Secret-Scan um die Anwendungs- und SMTP-Tokenformate.


## Version und Buildidentität

`package.json` hält `version` (Semver, aktuell 7.0.0) unabhängig von `buildNumber` (28) und `buildDate` (09112026). `src/buildInfo.js` liest diese gespeicherten Werte ohne neues Startdatum und erzeugt die Anzeige `28.09112026`. `get-version` liefert Version und Buildinformationen getrennt; beide sind auch ausserhalb des Debug-Modus im Drawer sichtbar. Der Release-Plan übergibt Buildnummer und Datum per `extraMetadata` an beide Builder-Schritte und verwendet sie im Installer-Dateinamen. Das Datum wird einmal beim Bauen in Europe/Zurich ermittelt. `build-local.ps1` zählt von der höheren lokalen/Paketnummer weiter und speichert den Zähler erst nach Erfolg. Eine Prüfung des Argumentplans erstellt keinen Installer. `test/debugTools.test.js` prüft diese Regeln sowie die Update-Zustände und Testmail-Vorlage.


## Hausaufgaben und lokale Klassenregeln

`settings.json.classSettings` ordnet Klassenregeln über den kanonischen Dateipfad und Klassennamen zu. Alle Regeln bleiben im App-Profil. `getClassSettings` übernimmt bisherige flache Einstellungen als Vorgaben, bis eine Klasse eigene Werte erhält. `saveScopedSettings` speichert globale Werte und die ausgewählten Klassenwerte zusammen, ohne andere Klassen, KI-Freigaben oder Noten zu verändern. Wiggers-Sperren bleiben pro Datei/Klasse/Person getrennt; Deaktivieren entfernt nur die Sperren der betroffenen Klasse. `get-ui-settings` und Klassenwechsel liefern die effektiven Regeln. Die Auswahl im Einstellungsfenster wechselt keine Session. Ein geänderter Dateikontext wird vor dem Speichern zurückgewiesen.

Klassenspezifisch: `extraJokerAfterThreeGrades`, Nachtragen (Migrationsstatus pro Klasse), `probabilityDecreaseFactor`, `boostNeverSelected`, `neverSelectedBoostFactor`, `wiggersRuleEnabled`, `wiggersRuleDurationMinutes`, `homeworkEnabled` (true), `homeworkJokerThreshold` (2) und `homeworkWeight` (1.2). Global: Theme/visuelle Effekte, Meldungsdauer, KI und API-/SMTP-Integration. Historisch abgeschlossene dateiweite Joker-Nachträge bleiben als abgeschlossen erhalten; neue Nachträge betreffen ausschliesslich die ausgewählte Klasse.

Hausaufgaben sind eigenständige Einträge vom Typ `homework`. Die Auswahl verwendet alle anwesenden Schüler, auch nach sechs Repetitionen. Neue Sessions halten auch ihren Dateipfad fest. Der Hauptprozess prüft Session, Datei, Klasse, Anwesenheit, Notenbereich und Aktivierung; eine laufende Session darf keine Hausaufgabennote in eine andere Datei schreiben. Jede zweite abgeschlossene Vorführung vergibt standardmässig einen Joker; Schwelle und Gewicht kommen aus den effektiven Klassenregeln. Das zum Zeitpunkt der Bewertung verwendete Gewicht ist Teil des Noteneintrags und bleibt bei späteren Einstellungsänderungen erhalten. Repetitionswahrscheinlichkeit und ihre sechs Notenplätze werden nicht durch Hausaufgaben verändert.

`homework:status` liefert die Vorschau für die aktive Klasse. `homework:prepare` verlangt deren Dateifingerabdruck, prüft Dateisperren, erstellt eine unveränderte Sicherung, schreibt/prüft eine temporäre Datei und ersetzt die Quelle erst nach erneutem Fingerabdruck-/Sperrtest. Die Erweiterung hängt hinter vorhandenen eigenen Spalten an. Rechts stehen eine gewichtete Gesamt-/gerundete Note und sechs anfängliche Noten-/Datumspaare. Weitere Plätze werden ohne feste Obergrenze angehängt. Verborgene Spalten halten das historische Gewicht und eine stabile Eintragskennung samt bereits vergebenem Jokerbonus. Das sind Daten einzelner Bewertungen, keine Klassenregeln; leere Plätze enthalten weder Gewicht noch Konfiguration. Die ursprünglichen Repetitionsformeln bleiben erhalten. Die mitgelieferte Vorlage enthält diesen Bereich bereits auf allen drei Klassenblättern.

Gesamtwert: `(Summe Repetitionsnoten + Summe Hausaufgabennote × gespeichertes Gewicht) / (Anzahl Repetitionsnoten + Summe gespeicherter Hausaufgabengewichte)`. Leere Noten zählen nicht als null. Eigene Spalten und Joker-Daten bleiben erhalten. Die Statistik nennt die gewichtete Herkunft ihres Notendurchschnitts. Schülerhistorien liefern Hausaufgabennoten getrennt.

Backup, ausstehende Exceländerungen, Undo/Redo und Sessionprotokoll unterstützen `homework`. `homeworkId` bleibt auch bei neuen Backup-/Warteschlangen-IDs stabil und verhindert doppeltes Nachschreiben und Jokervergabe. Ausstehende Einträge zählen für die nächste Joker-Schwelle. Nach erfolgreichem Nachschreiben werden Backup-/Undo-Status aktualisiert. Undo entfernt Note, Datum und den zugehörigen Bonus; Redo stellt denselben Eintrag wieder her. Personen werden vor dem Nachschreiben mit ihrer Excelzeile abgeglichen; beim Umordnen im Editor bewegen sich Noten, Daten und Kennungen gemeinsam.

`npm run test:homework` prüft Regeln, gewichtete Werte, Erweiterung, Sicherung, Dateisperren, Warteschlange, idempotentes Nachschreiben, Undo/Redo, Umordnen und die gelieferte Vorlage. Der native Desktoptest durchläuft zwei benotete Vorführungen, Joker, die Klassenumschaltung sowie die Trennung globaler/Klassenoptionen. `test:ui`, `test:kristall` und `--settings-sliders` ergänzen die bestehenden Prüfungen.


Die Einstellungsnavigation `#settings-navdock` liegt als `.session-tabs` direkt im Container, ausserhalb des bewegten `#content`. Dadurch teilt sie Material, Marker und Geometrie mit dem Session-Dock. `settings-active` blendet sie ein; Speichern/Verwerfen stehen darüber. Der Klassenwähler ist ein fester Kopfbereich der Klassenansicht. `get-settings` wählt ausschliesslich die aktive Sessionklasse derselben Datei vor. Ohne Session bleibt die Auswahl leer; eine explizite Dropdownwahl wird unabhängig von der aktiven Session geladen. Der Session-Dock ordnet Mitarbeit, Hausaufgaben, Repetition in dieser Reihenfolge an.

Das Einstellungsdock hebt die geerbte Electron-Ziehregion explizit mit `app-region: no-drag` auf. `settingsScope.js` und `homeworkPage.js` sind Teil der EXE-Dateiliste; `test/rendererPackaging.test.js` prüft alle lokalen Renderer-Einstiegsmodule gegen diese Liste.

## Abhängigkeiten (13. November 2026)

Electron 45, Electron Builder 26, Nodemailer 10 und ExcelJS 4.4 sind auf die aktuellen npm-Versionen aktualisiert. Tailwind 4 nutzt die separate CLI und explizite Theme-/Utilities-Imports ohne Preflight, um die bestehende Kristallgestaltung zu erhalten. Der Icon-Schritt lädt den benannten ESM-Export von rcedit bei Bedarf. Unbenutzte direkte Abhängigkeiten electron-packager, autoprefixer und postcss entfallen. Node >=22.12 ist erforderlich; der Release-Workflow verwendet Node 24.

Gezielte npm-Overrides schliessen bekannte Lücken in brace-expansion 1, den Minimatch-Versionen von app-builder-lib/@electron/universal und UUID von ExcelJS. Der Packagingtest verwendet den tatsächlichen FileMatcher des Builders, die Exceltests prüfen Lesen/Schreiben und die öffentliche Suite die Mailvorlagen. Installierte Versionen sind im Lockfile festgeschrieben.

`node scripts/ui-polish-smoke.js --settings-page` prüft die Einstellungsseite separat. Die Animationsprüfung wartet auf den IPC-getriebenen Seitenstart und liest eine definierte Eingangsposition der tatsächlich gestarteten GSAP-Timeline; Reduced-Motion-Prüfungen warten auf die asynchrone Medienänderung.

Der gemeinsame Kristalltakt beendet `animateFor` bei reduzierter Bewegung unmittelbar. Auch nachträgliche Neuzeichnungen dürfen keinen neuen Animationsframe einplanen; die Kristallprüfung sichert dies ab.
