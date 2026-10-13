# Project overview

Repetierer ist eine Electron-Desktop-App fuer Lehrpersonen. Sie liest Klassen und Schueler aus einer Excel-Arbeitsmappe, waehlt Schueler gewichtet zufaellig fuer Repetitionen aus, speichert Noten/Joker wieder in Excel und fuehrt zusaetzlich lokale Unterrichts-Sessions, Mitarbeitseintraege, Mitmachnoten, Protokolle, Backups und einen optionalen lokalen KI-Assistenten.

Tech-Stack:
- Electron/Node.js/CommonJS mit `nodeIntegration: true` und `contextIsolation: false`.
- Renderer: statisches HTML/CSS/JS in `Repetierer/public`, Tailwind wird in `public/tailwind.css` gebaut.
- Persistenz: Excel ueber `exceljs`, lokale JSON/env/log-Dateien im Electron-`userData`-Ordner.
- Externe Integrationen: GitHub Releases/electron-updater, AnalyticsManager, Cloudflare Workers + D1, Brevo SMTP, lokale Ollama-API.

Wichtige Einstiegspunkte:
- `Repetierer/package.json`: npm-Skripte, Electron-Builder-Konfiguration.
- `Repetierer/app.js`: Electron-App, Splash/MainWindow, Auto-Update.
- `Repetierer/src/mainProcess.js`: zentrale IPC-Schicht und Main-Prozess-Orchestrierung.
- `Repetierer/public/index.html`: Haupt-UI mit allen Views/Modals.
- `Repetierer/public/renderProcess.js`, `public/assessmentRenderer.js`, `public/aiRenderer.js`: Renderer-Logik.

# Repository structure

| Pfad | Zweck und Verantwortung | Wichtige Dateien |
| --- | --- | --- |
| `Repetierer/` | Eigentliche App. Root dieses Repositories enthaelt zusaetzlich `AGENTS.md`, `nextup.txt` und diese Datei. | `package.json`, `app.js`, `README.md`, `DESIGN.md` |
| `Repetierer/src/` | Main-Prozess-Fachlogik, Excel-Zugriff, Storage, Sessions, Integrationen. | `mainProcess.js`, `program.js`, `excelReader.js`, `storage.js`, `teachingSessions.js`, `classStats.js` |
| `Repetierer/src/assessment/` | Mitmachnote: lokale Daten, Cloudflare-Client, Matching, Lehrerbewertung, Mailversand. | `mainAssessment.js`, `assessmentStore.js`, `assessmentApiClient.js`, `integrationConfig.js`, `teacherCredentials.js`, `assessmentDelivery.js` |
| `Repetierer/cloudflare/` | Isolierter Worker, D1-Schema, responsive Schülerumfrage und Deployment. | `wrangler.jsonc`, `src/worker.js`, `migrations/0001_assessment.sql`, `web/*`, `assets/fonts/*` |
| `Repetierer/src/ai/` | Lokaler Ollama-Assistent mit Tool-Registry und bestaetigten Schreibaktionen. | `ollamaService.js`, `assistantService.js`, `toolRegistry.js` |
| `Repetierer/public/` | Renderer-UI, Styles, Icons, Splash. | `index.html`, `renderProcess.js`, `assessmentRenderer.js`, `aiRenderer.js`, `style.css`, `tailwind.input.css`, `splash.html`, `splash.js` |
| `Repetierer/docs/` | Funktions- und Integrationsdokumentation. | `mitmachnote.md`, `cloudflare-deployment.md`, `lokale-ki.md` |
| `Repetierer/scripts/` | Release-, Smoke- und Hilfsskripte. | `release-with-icon.js`, `generate-release-notes.js`, `cloudflare-ci-check.js`, `cloudflare-electron-e2e.js`, `ai-live-smoke.js` |
| `Repetierer/test/` | Node-Test-Suite fuer Fachmodule, Integrationen und UI-nahe Logik. | `*.test.js` |
| `Repetierer/builds/` | Electron-Builder-Ausgabe. Alte Builds sollen nicht geloescht werden; hoechstens ueberschreiben. | wird von Build-Skripten erzeugt |

# Modules

## Electron shell and IPC

- Verantwortung: App-Lifecycle, Fenster, Splash, Auto-Update, IPC-Grenze zwischen Renderer und Fachlogik.
- Zentrale Dateien/Funktionen: `app.js` (`createSplashWindow`, `createMainWindow`, `checkForUpdatesWithConsent`), `src/mainProcess.js` (IPC-Handler, Export/Import, Undo/Redo, Reporting).
- Oeffentliche Schnittstelle: Electron-IPC-Kanaele wie `file`, `class`, `start`, `ok`, `joker`, `start-teaching-session`, `assessment:*`, `ai:*`.
- Abhaengigkeiten: `program.js`, `storage.js`, `teachingSessions.js`, Assessment-Module, AI-Module, Electron `dialog`, `BrowserWindow`, `net`.

## Repetition and Excel domain

- Verantwortung: aktuelle Datei/Klasse, Personenliste, gewichtete Auswahl, Noten/Joker, Abwesenheiten, Wiggersche Regel, Excel-Migrationen.
- Zentrale Dateien/Funktionen: `src/program.js` (`setFile`, `setClass`, `selectPerson`, `saveGrade`, `setJoker`, `setAbsences`, `getProbabilities`), `src/excelReader.js` (`init`, `read`, `write_grade`, `write_joker`, `apply_entries`, `undo_entry`).
- Oeffentliche Schnittstelle: Methoden aus `program.js`, aufgerufen nur aus Main-Prozess/AI-Registry.
- Abhaengigkeiten: `excelReader.js`, `storage.js`, `classStats.js`, `exceljs`, Dateisystem/Windows-Lockcheck.

## Local storage and logs

- Verantwortung: App-Einstellungen, zuletzt gewaehlte Excel-Datei, Backups, Pending-Excel-Eintraege, Migrationstatus, Logs, Pfade.
- Zentrale Datei/Funktionen: `src/storage.js` (`getPaths`, `getAppSettings`, `saveAppSettings`, `addBackupEntry`, `addPendingExcelEntry`, `logEvent`).
- Oeffentliche Schnittstelle: CommonJS-Exports aus `storage.js`.
- Abhaengigkeiten: Electron `app.getPath('userData')`, `fs`, `path`.

## Teaching sessions and participation

- Verantwortung: aktive Unterrichts-Session, An-/Abwesenheits-Snapshot, Mitarbeitspunkte, Repetitionshistorie pro Session, Auswertung.
- Zentrale Datei/Funktionen: `src/teachingSessions.js` (`createTeachingSessionStore`, `startSession`, `endSession`, `addParticipation`, `addRepetition`, `getParticipationSummary`).
- Oeffentliche Schnittstelle: Store-Objekt, im Main-Prozess lazily erzeugt.
- Abhaengigkeiten: lokales JSON unter `userData`, aktuelle Personen aus `program.js`.

## Mitmachnote / assessment

- Verantwortung: Beurteilungsrunden, Selbstbeurteilungen aus externer API, Lehrerbewertungen, Matching auf Schueler, Einladungslinks, Versand persoenlicher Resultate.
- Zentrale Dateien/Funktionen: `src/assessment/mainAssessment.js` (`registerAssessmentHandlers`), `assessmentStore.js` (`AssessmentStore`), `assessmentApiClient.js`, `matching.js`, `integrationConfig.js`, `assessmentDelivery.js`, `smtpMailer.js`.
- Oeffentliche Schnittstelle: `ipcMain.handle`-Kanaele `assessment:get-overview`, `assessment:create-round`, `assessment:get-round`, `assessment:sync-responses`, `assessment:save-teacher`, `assessment:send-round`, Config-/Brevo-Kanaele.
- Abhaengigkeiten: aktuelle Excel-Klasse ueber `program.getClassPersons`, lokale `mitmachnoten.json`, Cloudflare-Worker-API, Electron `safeStorage` fuer Lehrerzugang, Brevo SMTP ueber `nodemailer`.

## Cloudflare survey service

- `cloudflare/src/worker.js` liefert Umfrage und JSON-API unter `/app/services/repetierer/` auf `timolu.ch` und `akzuwo.ch`; bestehende Projekte und DNS bleiben erhalten. Beide Domains verwenden dieselbe produktive D1-Datenbank; Staging hat einen eigenen Worker und eine eigene Datenbank. Details/IDs: `docs/cloudflare-deployment.md`.
- D1 speichert `teachers` (individuelle Token-Hashes mit Ablauf/Widerruf), `rounds` (UUID, Eigentümer, Raster, Fristen) und `responses` (Name, Mail, Antworten, rundenspezifischer Browser-Hash). Keine Klassenlisten, IP-Adressen oder User-Agents. Parametrisierte SQL-Abfragen und eindeutige Indizes sichern einmalige Abgaben atomar ab.
- Verwaltungszugriffe brauchen ein individuelles Anwendungstoken und werden auf den Rundeneigentümer begrenzt. Electron speichert dieses mit dem Betriebssystem-Schlüsselspeicher; Cloudflare-OAuth/Verwaltungsschlüssel sind nie Teil der App. Provisionierung/Widerruf: `scripts/cloudflare-teacher.js`.
- Rundenerstellung ist ueber die UUID wiederholbar; `pending` wird erst nach Online-Bestaetigung `open`. `closed` sperrt Abgaben. Antwortseiten und stabile IDs erlauben idempotente lokale Imports. `cloudFormUrl` bewahrt die zuletzt erreichbare Umfragedomain auch nach Neustart/offline. Alte lokale Runden bleiben Archive.
- Abgelaufene Runden werden sofort unzugänglich, täglich per Cron samt Antworten gelöscht; höchstens 90 Tage ab Erstellung bzw. 30 Tage nach Schliessung. Anbieter-Wiederherstellungskopien und lokale Archive sind separat zu berücksichtigen.
- Die Webumfrage nutzt eigene statische Vanilla-Dateien mit denselben Farben und der lokalen Geist-Schrift. Keine fremden Scripts/Fonts; strikte CSP, `no-store, no-transform`, Origin-Prüfung und Datenvalidierung.
- Cloudflare Workers Builds zieht `Akzuwo/Repetierer`, Branch `master`, aus GitHub. `npm run cloudflare:check` testet lokal simuliertes D1 ohne Desktop-GUI; `npm run cloudflare:deploy` veröffentlicht Worker und Webumfrage zusammen. Andere Branches deployen nicht produktiv. SQL-Schemaänderungen bleiben versionierte D1-Migrationen und werden vor Übernahme inkompatibler Änderungen gezielt angewendet; Schülerdaten werden bei einem Push nicht neu angelegt.

## Local AI assistant

- Verantwortung: lokaler Chat, Ollama-Setup/Status, Tool-Calling, Navigation und bestaetigte Schreibaktionen.
- Zentrale Dateien/Funktionen: `src/ai/ollamaService.js` (`createOllamaService`, Modell `qwen3:4b-instruct`), `assistantService.js` (`createAiAssistantService`), `toolRegistry.js` (`createAiToolRegistry`).
- Oeffentliche Schnittstelle: IPC-Handler `ai:status`, `ai:setup`, `ai:ask`, `ai:confirm`, `ai:reset`; Renderer-Events `ai-activity`, `ai-navigate`.
- Abhaengigkeiten: lokale Ollama-API `http://127.0.0.1:11434`, `program.js`, `teachingSessions.js`, Main-Prozess-Callbacks fuer Schreiben.

## Renderer UI

- Verantwortung: alle sichtbaren Screens, Navigation, Modals, Statusanzeigen, Formulare, Renderer-seitige State-Variablen.
- Zentrale Dateien: `public/index.html`, `public/renderProcess.js`, `public/assessmentRenderer.js`, `public/aiRenderer.js`, `public/style.css`.
- Oeffentliche Schnittstelle: DOM-Events und Electron `ipcRenderer.send/invoke`.
- Abhaengigkeiten: Main-Prozess-IPC; keine direkte Nutzung der Fachmodule.

## Build, release and docs

- Verantwortung: CSS-Build, Electron-Packaging, lokale Buildnummern, Release Notes, Smoke-Tests.
- Zentrale Dateien: `package.json`, `build-local.ps1`, `scripts/release-with-icon.js`, `scripts/generate-release-notes.js`, `release-notes.txt`.
- Abhaengigkeiten: `electron-builder`, `rcedit`, Tailwind CLI, GitHub-Publish-Konfiguration.

# Data flow

## Repetition / grade flow

`public/renderProcess.js` sendet `start`/`ok`/`joker` -> `src/mainProcess.js` prueft aktive Session und ruft `program.js` -> `program.js` waehlt Person oder erstellt Entry -> `excelReader.js` schreibt Excel -> `mainProcess.js` schreibt Backup/Pending-Entry in `storage.js`, aktualisiert Session ueber `teachingSessions.js` und sendet UI-Events zurueck.

Wenn Excel gesperrt ist, wird die Aktion lokal als Backup/Pending-Excel-Eintrag gespeichert und spaeter ueber `flush-pending-excel` mit `program.applyPendingExcelEntries` in Excel angewendet.

## Teaching session / participation flow

UI waehlt Klasse und Abwesenheiten -> `start-teaching-session` -> `program.setAbsences` und `teachingSessions.startSession` -> UI wechselt zwischen Mitarbeit/Repetition. Mitarbeitspunkte gehen ueber `add-participation-entry` direkt in `unterrichtssessions.json`; Repetitionsaktionen werden nach erfolgreichem Backup zusaetzlich als `repetitionEntries` in der aktiven Session protokolliert.

## Mitmachnote flow

`public/assessmentRenderer.js` -> `assessment:*` IPC -> `mainAssessment.js` validiert Datei/Klasse -> `AssessmentStore` speichert UUID/Raster lokal -> `AssessmentApiClient` erstellt die Online-Runde mit Bearer-Token. Schueler-Webumfrage -> Worker -> atomare D1-Einfuegung. Sync -> paginierte Eigentümer-API -> lokale Antwort-IDs/Matching -> UI. Bei Netzwerkfehlern wechselt der Client von `timolu.ch` auf `akzuwo.ch`; Einladungen folgen der erreichbaren Domain. Versand bleibt `assessmentDelivery.js` + `mailTemplates.js` + `smtpMailer.js`.

## AI assistant flow

`public/aiRenderer.js` -> `ai:ask` -> `assistantService.js` ruft Ollama mit Tool-Definitionen -> `toolRegistry.js` liest Daten aus `program.js`/`teachingSessions.js` oder fordert Bestaetigung fuer Schreibaktionen an -> UI bestaetigt ueber `ai:confirm` -> Main-Prozess schreibt Note/Mitarbeit ueber bestehende Services.

## Problem report / analytics flow

UI oeffnet Vorschau -> `mainProcess.buildReportPayload` liest bereinigte Logs aus `storage.js` -> `send-problem-report` sendet JSON an `https://analytics.akzuwo.ch` mit `appId: "Repetierer"` und optionalem API-Key.

# Data and persistence

- Excel-Arbeitsmappe: externe Datei, Pfad in `settings.json`. Gueltige Repetierer-Sheets haben `A1 = "repetierer"`, Namen in `A6:A30`, bis 6 Noten in `B:G`, Joker in `H`, Notendaten in `K:P`, Joker-Historie in `Q`.
- `settings.json`: Theme, Wahrscheinlichkeitsfaktor, Boosts, Extra-Joker, visuelle Effekte, Wiggersche Regel, letzte Excel-Datei, Update-News-Status, Wiggers-Penalties.
- `backup.json`: gespeicherte Aktionen und `pendingExcelEntries` fuer Excel-Konflikte.
- `migrations.json`: pro Excel-Datei abgeschlossene lokale Migrationen, z.B. Extra-Joker-Nachtrag.
- `repetierer.log`: technische Ereignisse, fuer UI-Log und Problemberichte.
- `unterrichtssessions.json`: Sessions mit Klasse, anwesenden/abwesenden Schuelern, Mitarbeit und Repetitionseintraegen.
- `mitmachnoten.json`: Assessment-Schema v3, Workbooks/Klassen/Schueler/Runden/Antworten/Bewertungen/Mailstatus.
- Lehrerbewertungen enthalten optional `teacherCriterionComments` nach Kriteriums-ID; die Felder erscheinen nach einer numerischen Lehrerbewertung. Kommentare bleiben lokal und werden escaped in die persönliche Brevo-Auswertung übernommen. Bestehende Daten bleiben ohne Kommentare gültig.
- `repetierer.env`: importierte Mitmachnoten- und Brevo-Konfiguration. Legacy `brevo-smtp.env` wird noch gelesen.
- `update-state.json`: in `app.js` gespeicherter Installationsstatus fuer Auto-Updates.
- Import/Export: Klassenlistenimport aus `.xlsx`/`.csv`; Sitzungsprotokoll-Export als CSV/PDF; lokale env-Konfigurationsdateien fuer Integrationen werden importiert und in `userData` kopiert.

# UI architecture

- `public/index.html` enthaelt die gesamte Haupt-App als Single-Window-UI: Titlebar, Drawer, Klassenauswahl, Repetition, Mitarbeit, Statistiken, Mitmachnote, Update-Panel, Settings und alle Modals.
- `public/renderProcess.js` steuert die Kern-UI: Drawer, Datei-/Klassenwahl, Session-Start/Ende, Repetition, Mitarbeit, Backup/Log, Import, Settings, Updates, Problemberichte und Animationen.
- `public/assessmentRenderer.js` steuert die Mitmachnoten-Views im bestehenden `#assessment-view`.
- `public/aiRenderer.js` steuert KI-Modal, Chat, Setup-Status, Tool-Aktivitaet, Bestätigungskarten und KI-Navigation.
- `public/visualFeedback.js` bindet die originale Thinking-Orbs-Canvas-Engine an den Vanilla-Renderer. `public/uiMotion.js` steuert GSAP-Timelines fuer Modals, Drawer, Ansichten, Schuelername, Chat und Fenster-Minimierung/Wiederherstellung. Das lokale Bundle `public/vendor/ui-motion.js` und Lizenzhinweise werden mit der App verpackt; Quellen und Versionen stehen in `docs/ui-sources.md`.
- Aktive Sessions zeigen Datei/Klassen als flache Navigation ohne Cards links und Mitarbeit/Repetition rechts. Kompakter Header und unterstrichene Bereichstabs bilden den App-Rahmen. Die vorhandene `#class-list` wird zwischen der flachen Startliste und `#session-class-slot` verschoben. Repetitionsauswahl steht unter dem Titel, Bewertung am unteren Rand; der Schuelername skaliert mit seiner verfuegbaren Breite. Bereichsbuttons spiegeln ihre Auswahl auch ueber `aria-pressed`.
- Modals sind statisch in `index.html` definiert und werden ueber CSS-Klassen/`display` geoeffnet. Wiederkehrende Klassen: `.modal`, `.modal-content`, `.btn-1`, `.btn-2`, `.btn-danger`, `.update-hidden`.
- Alle Modal-Oeffner verwenden `uiMotion.openDialog`/`closeDialog` mit `display: flex`; das Overlay zentriert den Dialog horizontal und vertikal. Hohe Inhalte scrollen innerhalb der viewportbegrenzten Dialogbox. Ein Interaktionselement besitzt genau eine Timeline; Wiedereroeffnen bricht alte Exit-Callbacks ab. Reduzierte Bewegung beendet laufende Uebergaenge im korrekten Endzustand.
- Navigation ist Renderer-State plus DOM-Toggling; es gibt keine Routing-Library. Viele Views liegen gleichzeitig im DOM und werden per Klasse/Visibility umgeschaltet.
- Styling: `public/style.css` ist Haupt-Designsystem mit CSS-Variablen fuer Themes (`:root`, `:root[data-theme="classic"]`), dunklem Standardtheme und gemeinsamen Button-/Modal-/Zeilen-Klassen. Geist Variable liegt samt OFL-Lizenz unter `public/fonts` und wird lokal vorgeladen; Roboto bleibt fuer die separate Splash-Schrift erhalten. `public/tailwind.input.css` liefert Utilities und die optionale Logoanimation; doppelte Surface-/Control-Regeln und Routineanimationen wurden entfernt. Im HTML wird Tailwind zuerst und das Designsystem danach geladen.
- Ladefeedback: `renderProcess.js` stellt `window.createRepetiererLoadingState` fuer echte IPC-Lesevorgaenge in Statistiken und Mitmachnote bereit. Schreibvorgaenge erhalten ihren Inhalt und sperren ihn mit `inert`; der KI-Renderer zeigt waehrend echter Anfragen eine Thinking Orb und Tool-Status an. Neue Animationen respektieren `prefers-reduced-motion`.
- Neue UI-Funktionen werden normalerweise in `index.html` markupped, in `renderProcess.js`/spezialisiertem Renderer verdrahtet und in `style.css` gestylt. Bei UI-Aenderungen zuerst `Repetierer/DESIGN.md` lesen.

# External integrations

- GitHub Updates: `app.js` nutzt `electron-updater` mit Feed `Akzuwo/Repetierer`, Auto-Download erst nach Zustimmung.
- AnalyticsManager: `src/analyticsConfig.js` und `mainProcess.js` senden Problemberichte an `https://analytics.akzuwo.ch`; Key aus `REPETIERER_ANALYTICS_KEY`, `ANALYTICS_KEY`, `analytics-config.local.json` oder legacy `report-config.local.json`.
- Cloudflare: `ASSESSMENT_API_URL`, `ASSESSMENT_TEACHER_TOKEN` (lokal verschluesselt), optional `ASSESSMENT_FALLBACK_API_URL`. Worker + D1, Konfiguration in `cloudflare/wrangler.jsonc`; Einrichtung in `docs/cloudflare-deployment.md`.
- Brevo SMTP: `BREVO_SMTP_HOST`, `BREVO_SMTP_PORT`, `BREVO_SMTP_USER`, `BREVO_SMTP_PASSWORD`, `BREVO_FROM_EMAIL`, `BREVO_FROM_NAME`; Versand ueber `nodemailer`.
- Ollama: lokale API `127.0.0.1:11434`, Modell `qwen3:4b-instruct`, Windows-Auto-Installation von `https://ollama.com/download/OllamaSetup.exe`.
- Windows/Excel Lockcheck: `excelReader.js` prueft `~$`-Lockdateien und unter Windows exklusiven Dateizugriff per PowerShell.

# Build and runtime

- Entwicklung: in `Repetierer/` `npm install`/`npm ci`, dann `npm start`. Das Skript baut CSS (`npm run build:css`) und startet `electron .`.
- Tests: `npm test` fuehrt `node --test --test-concurrency=1 test/*.test.js` aus.
- CSS: `npm run build:css` nutzt `tailwind.config.js`, `public/tailwind.input.css` -> `public/tailwind.css`.
- UI-Vendor-Bundle: `npm run build:ui-vendor` nutzt esbuild und `scripts/ui-vendor-entry.mjs` -> `public/vendor/ui-motion.js`; benoetigt beim App-Start keinen Build und kein React.
- Renderer-Smoke: `npm run test:ui` startet ein isoliertes, unsichtbares Electron-Fenster mit kontrolliertem IPC und schreibt Screenshots nach `.ui-polish-smoke/`.
- Packaging: `npm run build` nutzt `electron-builder`; Ausgabe in `Repetierer/builds`.
- Lokaler Windows-Build: `npm run build:local` bzw. `build-local.ps1`; installiert bei fehlendem `node_modules` per `npm ci`, baut CSS, setzt lokale Buildnummer und ruft `scripts/release-with-icon.js`.
- Release: `npm run release` oder `npm run release:ci`; `release-with-icon.js` baut erst `win-unpacked`, setzt Icon via `rcedit`, schreibt `app-update.yml`, baut dann NSIS-Installer.
- Plattformbesonderheiten: App ist klar auf Windows ausgerichtet (NSIS, `.ico`, AppUserModelId, PowerShell-Lockcheck, automatische Ollama-Installation nur Windows), Electron laeuft konzeptionell auch auf anderen Plattformen mit Einschraenkungen.

# Dependency map

- `Renderer UI -> Electron IPC -> mainProcess.js -> program.js -> excelReader.js -> Excel workbook`
- `Renderer UI -> Electron IPC -> mainProcess.js -> storage.js -> userData JSON/log/env`
- `Repetition Save -> program.js -> excelReader.js -> storage backup/pending -> teachingSessions.addRepetition -> UI`
- `Session UI -> mainProcess.js -> teachingSessions.js -> unterrichtssessions.json`
- `Mitmachnote UI -> assessmentRenderer.js -> assessment:* IPC -> mainAssessment.js -> AssessmentStore -> mitmachnoten.json`
- `Mitmachnote Create/Sync/Close/Delete -> mainAssessment.js -> AssessmentApiClient -> Cloudflare Worker -> D1`
- `Schueler-Webumfrage -> Cloudflare Worker -> D1 -> Electron Sync -> mitmachnoten.json`
- `Mitmachnote Mail -> mainAssessment.js -> assessmentDelivery/mailTemplates/smtpMailer -> Brevo SMTP`
- `AI UI -> aiRenderer.js -> ai:* IPC -> assistantService -> Ollama -> toolRegistry -> program/teachingSessions`
- `Problem Report UI -> mainProcess.js -> storage logs + analyticsConfig -> AnalyticsManager`
- `Build -> package.json scripts -> Tailwind/electron-builder/release-with-icon -> builds`

# Where to make changes

| Aufgabe | Primaer relevanter Bereich |
| --- | --- |
| Repetitions-Auswahl/Wahrscheinlichkeiten aendern | `Repetierer/src/program.js`, Tests in `test/repetition.test.js` |
| Excel-Format, Noten/Joker schreiben, Locking aendern | `Repetierer/src/excelReader.js`, ggf. `src/program.js` |
| Pending-Excel, Backup, Logs, Settings aendern | `Repetierer/src/storage.js`, `src/mainProcess.js` |
| Unterrichts-Session oder Mitarbeit aendern | `Repetierer/src/teachingSessions.js`, IPC in `src/mainProcess.js`, UI in `public/renderProcess.js` |
| Repetition-UI oder Hauptnavigation aendern | `Repetierer/public/index.html`, `public/renderProcess.js`, `public/style.css`, vorher `DESIGN.md` |
| Modals/Dialoge aendern | `public/index.html`, `public/renderProcess.js`, `public/style.css`, vorher `DESIGN.md` |
| Mitmachnote-Datenmodell/Rundenlogik aendern | `src/assessment/assessmentStore.js`, `mainAssessment.js`, Tests `assessmentStore.test.js` |
| Mitmachnote-UI aendern | `public/assessmentRenderer.js`, `public/index.html`, `public/style.css` |
| Umfrage/API-Vertrag aendern | `cloudflare/src/worker.js`, `cloudflare/web/*`, `src/assessment/assessmentApiClient.js`, `test/cloudflareE2E.test.js` |
| Brevo/SMTP-Versand aendern | `src/assessment/assessmentDelivery.js`, `mailTemplates.js`, `smtpMailer.js`, `integrationConfig.js` |
| Lokalen KI-Assistenten aendern | `src/ai/*`, IPC in `src/mainProcess.js`, UI in `public/aiRenderer.js`, Doku `docs/lokale-ki.md` |
| Problemberichte/Analytics aendern | `src/analyticsConfig.js`, `src/mainProcess.js` |
| Auto-Update aendern | `Repetierer/app.js`, `package.json` build/publish-Konfiguration |
| Build/Installer/Release aendern | `package.json`, `build-local.ps1`, `scripts/release-with-icon.js`, `scripts/generate-release-notes.js` |
| Styling/Designsystem aendern | `public/style.css`, `public/tailwind.input.css`, `tailwind.config.js`, vorher `DESIGN.md` |

# Architecture rules / conventions

- Renderer sprechen mit Daten/Fachlogik ueber Electron IPC; direkte Zugriffe auf `src/*` passieren nur im Main-Prozess.
- Fachlogik ist CommonJS-basiert und nutzt explizite Exports statt Klassen/Frameworks, ausser `AssessmentStore` und API-/Service-Factories.
- Excel bleibt Quelle fuer Klassen, Schueler und Repetitionsnoten; lokale JSON-Stores ergaenzen App-Zustand, Sessions, Backups und Mitmachnoten.
- Speicherzugriffe sind dateibasiert; Assessment- und Integrationskonfiguration schreiben atomar ueber Tempdatei/Rename, `storage.js` schreibt einfache JSON-Dateien.
- Schreibaktionen mit Excel-Konflikt werden nicht verworfen, sondern als Backup/Pending-Entry erhalten.
- KI-Schreibaktionen muessen bestaetigt und unmittelbar vor Commit erneut validiert werden.
- Tests laufen mit Node Test Runner und sind nach Fachbereichen benannt.
- UI-Konventionen stehen in `DESIGN.md`: dunkle ruhige Oberflaeche, pro Screen klarer Fokus, Modals kompakt, bestehende Klassen/Patterns wiederverwenden.
- Alte Builds werden nicht geloescht; Build-Skripte duerfen vorhandene Artefakte hoechstens ueberschreiben.

# Known architectural weaknesses

- `Repetierer/src/mainProcess.js` ist sehr gross und buendelt IPC, Session-Orchestrierung, Import/Export, Reporting, Undo/Redo und Hilfsfunktionen. Aenderungen an zentralen Flows brauchen dort oft gezielte Kontextsuche.
- `Repetierer/public/renderProcess.js` ist ebenfalls gross und verwaltet viele Views/Modals mit gemeinsamem Renderer-State. UI-Aenderungen koennen unerwartet mehrere Screens beruehren.
- `Repetierer/public/index.html` enthaelt fast die gesamte UI in einer Datei; Markup, Modal-Struktur und IDs sind stark mit den Renderer-Dateien gekoppelt.
- `program.js` und `excelReader.js` halten aktuellen Workbook-/Klassenstatus als Modulzustand. Das ist einfach, aber erschwert parallele Dateien/Klassen und macht Tests/Flows zustandsabhaengig.
- Renderer laeuft mit `nodeIntegration: true` und `contextIsolation: false`; das entspricht dem vorhandenen Electron-Muster, ist aber sicherheitstechnisch eine starke Kopplung zwischen UI und Node-Kontext.
- Es gibt mehrere Persistenzorte mit unterschiedlichen Quellen der Wahrheit: Excel fuer Repetitionsdaten, `unterrichtssessions.json` fuer Sessions/Mitarbeit, `mitmachnoten.json` fuer Assessments. Neue Features muessen bewusst entscheiden, welche Quelle massgeblich ist.
