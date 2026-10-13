# Cloudflare-Deployment

## Ressourcen

| Umgebung | Worker | D1 |
| --- | --- | --- |
| Produktion | `repetierer-assessment` | `f891e94d-70cf-49f7-883f-4571693afd3e` |
| Staging | `repetierer-assessment-staging` | `aa6e995f-d679-4cd4-b94e-06926fa0742e` |

Account: `b28433d3be7f249de5b132c4ac59d351`. Beide D1-Datenbanken wurden mit `wrangler d1 create … --jurisdiction eu` erstellt. Die Jurisdiktion ist eine Datenbankeigenschaft, kein Binding-Feld.

Nur `timolu.ch/app/services/repetierer/*` und `akzuwo.ch/app/services/repetierer/*` werden an den neuen Worker geroutet. Bestehende Custom Domains und Worker `portfolio`, `hwm3` und `analytics-manager` bleiben erhalten. DNS wurde nicht geändert. Produktion hat keine workers.dev-/Preview-URLs. Staging: `https://repetierer-assessment-staging.timowigger8.workers.dev/app/services/repetierer/`.

Keine kostenpflichtige Zusatzfunktion wird benötigt oder aktiviert. Höchstens 200 Antworten je Runde, normalerweise die lokale Klassengrösse, 200 aktive Runden pro Lehrer, 32 KiB Anfragekörper und begrenzte Textfelder beschränken Last und Datenmenge. Die Free-Tier-Kontingente im Dashboard überwachen; Missbrauch kann Tageskontingente erschöpfen. Kein IP-basiertes Anwendungs-Rate-Limit.

## Deployment

Im App-Ordner mit Node 24 / Wrangler 4:

```powershell
npm ci
$env:NODE_USE_SYSTEM_CA='1'
npx wrangler login
npm run test:cloudflare
npx wrangler d1 migrations apply repetierer-assessment-staging --config cloudflare/wrangler.jsonc --env staging --remote
npx wrangler deploy --config cloudflare/wrangler.jsonc --env staging
npx wrangler d1 migrations apply repetierer-assessment --config cloudflare/wrangler.jsonc --remote
npm run cloudflare:deploy
```

`NODE_USE_SYSTEM_CA=1` verwendet den Windows-Zertifikatsspeicher bei schulischer TLS-Inspection; die Zertifikatsprüfung bleibt aktiv. Wrangler-OAuth liegt nur in Wranglers Benutzerkonfiguration. Neue Installationen passen Account-/DB-/Routenwerte in `cloudflare/wrangler.jsonc` an. `npm run cloudflare:types` regeneriert Binding-/Runtime-Typen.

Lokal:

```powershell
npx wrangler d1 migrations apply repetierer-assessment --config cloudflare/wrangler.jsonc --local
npx wrangler dev --config cloudflare/wrangler.jsonc --local --test-scheduled
```

HTML/CSS/JS werden als Textmodule gebündelt; die vorhandene Geist-Schrift wird mit OFL-Lizenz selbst gehostet. `run_worker_first` erzwingt die Response-Header. `no-transform` verhindert zusätzlich injiziertes JavaScript ausschliesslich auf der Umfrage und ändert keine Einstellungen anderer Projekte.

## Automatische Veröffentlichung aus GitHub

Der produktive Worker ist mit Cloudflare Workers Builds an `Akzuwo/Repetierer` angebunden. Produktionsbranch: **master**. Root: `/`, Build-Befehl: `npm run cloudflare:check`, Deploy-Befehl: `npm run cloudflare:deploy`. Vorschau-Builds sind ausgeschaltet. Die vorhandene Cloudflare-GitHub-Verbindung und der bestehende Build-Token werden genutzt; an deren Berechtigungen wurde nichts geändert. Kein Deployment-Token liegt im Repository oder in Electron.

Nach jedem Push auf `master` lädt Cloudflare den Quellstand, installiert die gelockten npm-Abhängigkeiten, führt die Backend-/D1-/Import-/Mailtests ohne Electron-GUI aus und veröffentlicht bei Erfolg Worker und statische Webumfrage zusammen. `scripts/cloudflare-ci-check.js` ist dafür der reproduzierbare Einstiegspunkt. Die Desktop-GUI wird weiterhin lokal mit `npm run test:ui` und `npm run test:cloudflare` geprüft.

Der lokale Arbeitsbranch `ai-integration-testbranch` wird zunächst gepusht; erst nach dessen Übernahme in `master` nutzt Cloudflare diese Konfiguration. Andere Branches verändern die Produktion nicht. Bereits manuell veröffentlichte Änderungen werden dadurch nicht zurückgesetzt.

Der D1-Zugriffscode ist Teil des Workers und wird automatisch aktualisiert. **Schema-SQL wird nicht bei jedem Push erneut ausgeführt.** Dafür bleiben die nummerierten Dateien unter `cloudflare/migrations/` massgeblich. Neue Migrationen zuerst lokal/Staging prüfen und vor Übernahme einer darauf angewiesenen Worker-Version gezielt auf Produktion anwenden. Das schützt aktive Runden vor unbedachten Schemaänderungen. Für die Kriteriumskommentare ist keine D1-Migration nötig, da sie lokal gespeichert werden.

Git-Anbindung: [Cloudflare Workers Builds](https://developers.cloudflare.com/workers/ci-cd/builds/).

## Lehrerzugänge einrichten

```powershell
node scripts/cloudflare-teacher.js "Lehrer-Kürzel"
node scripts/cloudflare-teacher.js --revoke <Lehrer-UUID>
node scripts/cloudflare-teacher.js --rotate <Lehrer-UUID>
```

Provisionierung erzeugt 256 Bit Zufall, schreibt nur SHA-256 nach D1 und eine individuelle `.env` nach `.cloudflare-private/`. Ausgabe enthält Lehrer-ID und Dateipfad, niemals Tokenwerte. Zugangsdauer: ein Jahr. Kürzel ohne unnötige Personendaten wählen. Die Datei sicher an den betreffenden Lehrer übergeben und in den Einstellungen importieren. Widerruf sperrt den Zugang sofort und schliesst dessen Runden. Bei Geräteverlust widerrufen. Provisionierung ist nur mit Wrangler-Zugriff möglich; kein öffentlicher Provisionierungsendpunkt, kein globaler Cloudflare-Schlüssel im Client.

Für Staging `--staging` hinzufügen und `ASSESSMENT_STAGING_API_URL` auf die Staging-Adresse mit `/api` setzen. Testzugänge nach Nutzung widerrufen. `.cloudflare-private/`, `.wrangler/`, `.cloudflare-test/` sind ignoriert und nicht im Electron-Paket enthalten.

`--rotate` ersetzt das Token derselben Lehrer-ID und verlängert den Zugang um ein Jahr; die Eigentümerschaft vorhandener Runden bleibt erhalten. Das alte Token ist danach sofort ungültig. Die neue Datei erneut importieren. Eine Rotation kann einen widerrufenen Zugang wieder aktivieren, öffnet aber keine geschlossenen Runden.

Ein lokales Profil lässt sich unsichtbar konfigurieren:

```powershell
npx electron scripts/cloudflare-configure-local.js <Lehrer-env>
```

Der Helfer bewahrt Brevo, prüft den neuen Zugang und speichert ihn verschlüsselt. Nach erfolgreicher Rückleseprüfung wird die Quelldatei wie beim UI-Import entfernt.

## API und Sicherheit

Pfad relativ zu `/app/services/repetierer/api`:

| Methode / Pfad | Zugriff | Zweck |
| --- | --- | --- |
| `GET /health` | öffentlich | Dienst-/DB-Erreichbarkeit |
| `GET /public/rounds/:id` | Rundenlink | Titel, Raster, Status, Frist |
| `POST /public/rounds/:id/responses` | Rundenlink | Validierte einmalige Abgabe |
| `GET /teacher/me` | Bearer | Zugang prüfen |
| `PUT /teacher/rounds/:id` | Eigentümer | Idempotent erstellen |
| `GET /teacher/rounds/:id/responses?cursor=…` | Eigentümer | Antwortseiten à 100 |
| `POST /teacher/rounds/:id/close` | Eigentümer | Idempotent schliessen |
| `DELETE /teacher/rounds/:id` | Eigentümer | Runde samt Antworten löschen |

Parametrisierte SQL-Abfragen, gehashte Token, Ablauf/Widerruf und Eigentümerprüfung. Fremde Runden liefern 404. Keine CORS-Freigabe für Lehrerzugriffe; fremde Browser-Origins können keine Schülerabgabe senden. Textausgabe über `textContent`. CSP erlaubt Ressourcen/Verbindungen nur aus der eigenen Origin und blockiert Einbettung/Inline-Scripts. `no-store`, `no-transform`, `no-referrer`, HSTS und `nosniff` ergänzen den Schutz. Keine Inhalts-/Tokenlogs.

Der bestehende Electron-Renderer hat weiterhin Node-Integration. Betriebssystem-Verschlüsselung schützt ruhende Zugangsdaten, nicht ein kompromittiertes Gerät. Eine vollständige Renderer-Isolation wäre eine separate App-Migration.

Cron `17 2 * * *` entfernt abgelaufene Runden mit `ON DELETE CASCADE`. Schon vorher sind sie unzugänglich. Technische Wiederherstellungskopien folgen [D1 Time Travel](https://developers.cloudflare.com/d1/reference/time-travel/). Lokale Archivlöschung und Schul-Aufbewahrungsregeln: [mitmachnote.md](mitmachnote.md).

## Tests und Betrieb

`npm run test:cloudflare` startet Worker und D1 in einem isolierten Testordner. Geprüft werden Retry, Raster, Datenminimierung, acht konkurrierende Duplikate, Browser-Code-Sperre, fehlende/fremde Lehrerrechte, wiederholter Sync mit lokalem Neustart, Schliessung, Löschung, Widerruf und Cron-Kaskaden. Ein echtes unsichtbares Electron-Fenster prüft IPC, Betriebssystem-Verschlüsselung und das Browserformular auf Desktop und Mobil.

```powershell
$env:NODE_USE_SYSTEM_CA='1'
node scripts/cloudflare-remote-e2e.js <Lehrer-env>
```

Remote-Tests nutzen nur synthetische Namen und `e2e@example.ch`, versenden keine E-Mails und löschen ihre Online-Runden. Sie prüfen HTTPS, Fallback, API-Rechte und parallele Abgaben auf dem realen Worker. Ausgaben/Screenshots bleiben in `.cloudflare-test/`. Alte Builds werden nicht entfernt.

Deployment mit `wrangler deployments list --config cloudflare/wrangler.jsonc` prüfen. Rollback: `wrangler rollback <Version> --config cloudflare/wrangler.jsonc`. Ein Code-Rollback setzt keine D1-Migration zurück. Schemaänderungen zuerst in Staging prüfen; Sicherungen zeitlich begrenzt und geschützt halten, niemals Schülerdaten in Git exportieren.

## Verifikation vom 13. Oktober 2026

Lokales E2E und Staging-HTTPS-E2E bestanden. Produktiver Electron-/Browser-/API-Ablauf unter `akzuwo.ch` samt automatischem Fallback, Duplikatsrennen, Sync, Bewertung, Schliessung und Testdatenlöschung bestanden. Beide Domainrouten sind deployed. Der direkte Zugriff auf `timolu.ch` bleibt im aktuellen Schulnetz durch FortiGuard („Newly Observed Domain“) blockiert; `akzuwo.ch` ist der geprüfte, nutzbare Fallback. Bestehende DNS-/Domainkonfigurationen blieben erhalten.

Die vollständige Node-Suite besteht mit 49 Tests, der bestehende Electron-UI-Smoke-Test ebenfalls. Der Worker-Dry-Run und das lokale Startup-Profil sind geprüft. Die Produktion enthält nach den Tests keine Runden oder Schülerantworten; Testzugänge wurden widerrufen. Ein produktiver Lehrerzugang ist im lokalen Profil verschlüsselt eingerichtet. Der Windows-Installer liegt in `builds/cloudflare-migration-20261013-final/Repetierer-7.0.0.exe`; vorhandene Builds bleiben erhalten. Das geprüfte App-Paket enthält keine Deployment-Zugangsdaten.
