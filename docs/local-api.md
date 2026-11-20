# Lokale Lese-API (7.1.0)

Die API läuft als eigener fensterloser Electron-Hauptprozess ohne zusätzliche Node-Installation. Der normale Installer enthält den Dienst und startet ihn nach der Installation. Der Dienst richtet `Repetierer API.vbs` im persönlichen Windows-Autostart ein; auch die installierte Hauptapp gleicht den Start ab.

Die globale Einstellung **Lokale Lese-API** ist standardmässig aktiv. Deaktivieren wird unmittelbar gespeichert, beendet den Dienst und entfernt seinen Autostarteintrag. Appdaten und Schlüssel bleiben erhalten. Erneutes Aktivieren startet ihn wieder. Uninstall entfernt den Autostarteintrag und erhält die bestehende Option zum Behalten der Nutzerdaten.

Der Dienst importiert keine IPC-/Schreibmodule. Ohne Anfrage wird keine Exceldatei geladen; es gibt keine Rendererfenster oder Cloud-Synchronisierung. Ein Dateiwächter erkennt die Deaktivierung; ein 30-Sekunden-Abgleich dient als Rückfall. Die feste Loopback-Adresse verhindert eine zweite Dienstinstanz.

## Zugriff

Basis: `http://127.0.0.1:17839/v1`, nur IPv4-Loopback, kein LAN. Alle Endpunkte benötigen `Authorization: Bearer <Schlüssel>`. Der zufällige Schlüssel liegt im App-Profil unter `local-api-token`. `local-api.json` enthält Adresse, PID und Schlüsseldateipfad. **Zugriffspfad kopieren** kopiert Adresse und Dateipfad. Programme desselben Windows-Kontos können wie die App auf das Profil zugreifen. Schlüsseldateien nicht teilen oder versionieren.

```powershell
$profilePath = Join-Path $env:APPDATA 'Repetierer'
$apiToken = (Get-Content -LiteralPath (Join-Path $profilePath 'local-api-token') -Raw).Trim()
$headers = @{ Authorization = "Bearer $apiToken" }
$students = Invoke-RestMethod 'http://127.0.0.1:17839/v1/students?class=3a' -Headers $headers
$students.data
```

```python
import json, os, pathlib, urllib.request
token = (pathlib.Path(os.environ['APPDATA']) / 'Repetierer' / 'local-api-token').read_text().strip()
request = urllib.request.Request('http://127.0.0.1:17839/v1/sessions', headers={'Authorization': 'Bearer ' + token})
with urllib.request.urlopen(request) as response:
    sessions = json.load(response)['data']
```

## Endpunkte

Alle verwenden `GET`. Erfolg: `{ "success": true, "readAt": "ISO-Zeit", "data": … }`.

| Pfad | Inhalt |
| --- | --- |
| `/v1` | Endpunktverzeichnis |
| `/v1/health` | Dienstname, API-Version, `readOnly: true` |
| `/v1/settings` | Gespeicherte globale/Klasseneinstellungen und Dateizuordnung |
| `/v1/backup` | Änderungsverlauf und ausstehende Exceländerungen |
| `/v1/sessions` | Unterrichtssessions, Anwesenheit, Mitarbeit und Repetition |
| `/v1/assessments` | Lokal gespeicherte Runden, Antworten, Lehrerbewertungen und Versandstatus aller gespeicherten Arbeitsmappen |
| `/v1/migrations` | Migrationszustände |
| `/v1/workbook` | Aktuelle Exceldatei als getrennt eingelesene Klassen/Schülerakten: Noten, Hausaufgaben, Joker, Mitarbeit, Auswertungen und ausstehende Änderungen |
| `/v1/classes` | Klassenname und Schüleranzahl der aktuellen Exceldatei |
| `/v1/students?class=3a` | Schülerakten einer Klasse; ohne Parameter alle Klassen |
| `/v1/snapshot` | Obige Profildaten gemeinsam mit der aktuellen Arbeitsmappe |

Ausgegeben werden persistierte Daten, keine ungespeicherten Rendererentwürfe. Es gibt keinen Zugriff auf beliebige Dateipfade. Zugangsdaten, `.env`-Dateien, Tokens, SMTP-Konfiguration, Chromium-Caches und rohe Diagnoseprotokolle werden nicht ausgegeben. Schlüsselähnliche Felder werden zusätzlich rekursiv entfernt. Fachdateien werden nie verändert; nur dienstinterne Schlüssel-/Statusdateien werden erzeugt.

`/snapshot` ist keine Transaktion über Excel und mehrere JSON-Dateien. Bei gleichzeitigem Speichern kann `503 data_unavailable` auftreten; später erneut lesen. Fehlende optionale Profildateien liefern `{}`, eine noch nicht gewählte Exceldatei eine leere Klassenliste. Weitere Fehler: `401 unauthorized`, `403 origin_denied` (Browser-Origin oder fremder Host), `404 not_found`, `405 read_only` mit `Allow: GET`. Keine CORS-Freigabe und keine Schreibendpunkte.

## Entwicklung und Prüfung

```powershell
npx electron . --background-api --api-profile C:\Temp\Repetierer-API-Test
node --test test/localApi.test.js
node scripts/local-api-process-check.cjs
node scripts/api-mail-ui-check.cjs
```

Entwicklungs-/Testprozesse registrieren keinen echten Windows-Autostart. Der Prozesstest prüft HTTP nach dem Ende des Elternprozesses, der Autostarttest einen temporären Startupordner. Installerintegration: `build/api-service.nsh`. Ein echter Installations-/Windows-Neustarttest wurde ohne angeforderten Installer nicht durchgeführt.
