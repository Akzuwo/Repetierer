# Mitmachnote

Excel bleibt die Quelle für Klassen und Schüler. Selbstbeurteilungen, Lehrerbewertungen, Eigenkorrektur, Zuordnungen und Versandstatus liegen in `mitmachnoten.json` unter Electrons `userData`. Bestehende Bewertungen und synchronisierte Antworten bleiben erhalten und offline bearbeitbar.

## Ablauf

1. In den Einstellungen den individuellen Lehrerzugang importieren und die API-Verbindung prüfen.
2. Klasse wählen und eine Runde erstellen. Electron speichert lokal eine zufällige UUID und erstellt damit die Online-Runde. Die Klassenliste wird nicht hochgeladen.
3. **Einladung kopieren** liefert `https://timolu.ch/app/services/repetierer/?round=<UUID>`. Bei Netzwerksperren verwendet Electron automatisch `https://akzuwo.ch/app/services/repetierer/?round=<UUID>`.
4. Schüler antworten ohne Anmeldung. Das bestehende Raster mit sieben Kriterien, fünf Skalenwerten, zusätzlicher Qualität und Bemerkungen bleibt erhalten.
5. **Antworten aktualisieren** übernimmt Antworten über stabile IDs. Wiederholte Synchronisierung verändert bestehende Antworten und Lehrerbewertungen nicht.
6. **Runde schliessen** verhindert weitere Abgaben auf beiden Domains. Die Frist gilt bis 23:59:59 Uhr in `Europe/Zurich`; ohne Frist ist die Runde 30 Tage offen.
7. Bewertung, Eigenkorrektur und Brevo-Versand erfolgen lokal wie bisher. Eine Schülerabgabe verschickt keine automatische Mail.

Unter jedem Kriterium erscheint nach Auswahl der Lehrerbewertung automatisch ein optionales Kommentarfeld. Beim erneuten Öffnen bleiben Bewertung und Kommentar erhalten. Diese Kommentare werden zusätzlich zum allgemeinen Lehrerkommentar in der persönlichen Auswertung angezeigt. Sie werden nicht an die öffentliche Webumfrage oder D1 hochgeladen.

Bei fehlgeschlagener Online-Erstellung bleibt die lokale Runde `pending`. Öffnen oder Aktualisieren wiederholt denselben idempotenten Auftrag. Einladungen werden erst nach erfolgreicher Veröffentlichung angeboten. Alte lokale Runden bleiben Archive und werden nicht ungefragt hochgeladen; für neue Abgaben eine neue Runde erstellen.

## Konfiguration

Der Betreiber provisioniert pro Lehrer ein widerrufbares Anwendungstoken, keinen Cloudflare-Verwaltungsschlüssel:

```dotenv
ASSESSMENT_API_URL=https://timolu.ch/app/services/repetierer/api
ASSESSMENT_FALLBACK_API_URL=https://akzuwo.ch/app/services/repetierer/api
ASSESSMENT_TEACHER_TOKEN=<individuelles rt_-Token>
```

Der Fallback ist für diese Hauptadresse bereits voreingestellt; sein Konfigurationswert ist optional. Bei Netzwerkfehlern, Timeouts, Filterseiten oder Serverausfällen wird die zweite Domain verwendet. Ein gültiger Authentifizierungsfehler löst keinen Domainwechsel aus. Beide Domains teilen dieselben Daten. Die erfolgreiche Domainauswahl bleibt zehn Minuten im Main-Prozess gespeichert; Einladungen folgen ihr.

Die zuletzt erfolgreich verwendete Umfragedomain wird zusätzlich pro Runde lokal gespeichert. Damit behalten Einladungen auch nach einem App-Neustart oder offline den erreichbaren Fallback-Link.

Der Import prüft den Zugang online, verschlüsselt ihn mit Electron `safeStorage` und schreibt `repetierer.env` atomar. Unter Windows bindet DPAPI die Entschlüsselung an das Benutzerkonto. Ohne sicheren Schlüsselspeicher wird unverschlüsseltes Speichern abgelehnt. Token gelangen nicht zum Renderer, in Einladungen oder Protokolle. Für andere Benutzerkonten eine eigene ursprüngliche Zugangsdatei verwenden.

Die importierte Datei darf zusätzlich die unveränderten Brevo-Werte enthalten:

```dotenv
BREVO_SMTP_HOST=smtp-relay.brevo.com
BREVO_SMTP_PORT=587
BREVO_SMTP_USER=
BREVO_SMTP_PASSWORD=
BREVO_FROM_EMAIL=
BREVO_FROM_NAME=
```

Die App parst nur Key/Value-Zeilen, validiert vollständige Gruppen und prüft die gespeicherte Kopie vor dem Entfernen der importierten Datei. Bei Fehlern bleibt die Quelle erhalten. `brevo-smtp.env` wird weiterhin gelesen. Nicht mehr unterstützte Umfrage-Einstellungen werden beim Start entfernt; Brevo und lokale Beurteilungen bleiben erhalten.

## Duplikate und Datenschutz

D1 erzwingt `UNIQUE(round_id, name_key)` und ergänzend `UNIQUE(round_id, browser_hash)`. Namen werden Unicode-normalisiert, unsichtbare Formatzeichen entfernt, Leerzeichen vereinheitlicht und kleingeschrieben. Die Einfügung prüft Offenheit, Frist und Antwortlimit in derselben SQL-Anweisung. Gleichzeitige Abgaben umgehen die Sperre nicht. Es gibt keine IP-Sperre. Bei gemeinsam verwendeten Geräten separate Browserprofile verwenden, da auch der Browser-Code nur eine Abgabe pro Runde erlaubt.

Ohne Anmeldung wird die Identität eines Namens nicht bewiesen. Einen weitergegebenen Link können auch andere Personen benutzen. Die Lehrperson prüft falsche Zuordnungen lokal. Öffentliche Endpunkte liefern keine Antworten oder Schülerlisten.

Gespeichert werden Name, E-Mail für den persönlichen Resultatversand, Bewertungswerte, freiwillige Texte, Antwort-ID und Abgabezeit. Browser-Codes werden rundenspezifisch gehasht. Keine Anwendungsspeicherung von IP-Adressen, User-Agent oder Geräteprofilen. Keine Inhalts-/Tokenlogs; persistente Invocation-Logs sind ausgeschaltet. Anbieter-Infrastrukturlogs unterliegen dessen eigenen Richtlinien.

Online-Runden laufen spätestens nach 90 Tagen ab, nach Schliessung innerhalb von höchstens 30 Tagen. Danach sind sie sofort unzugänglich; ein täglicher Cron entfernt sie samt Antworten physisch. Cloudflares technische Wiederherstellungskopien können während der tarifabhängigen Time-Travel-Frist weiterexistieren. **Runde lokal und online löschen** entfernt die aktive Onlinekopie sofort und danach die lokale Runde. Bei Netzwerkfehlern bleibt die lokale Runde bis zur bestätigten Online-Löschung erhalten.

Vor Ablauf synchronisieren und lokale Kopien nach schulischen Aufbewahrungsregeln löschen. E-Mails bei Empfängern und persönliche Backups werden nicht automatisch entfernt. D1 wurde mit EU-Jurisdiktion erstellt; Worker-Anfragen werden im Cloudflare-Netz verarbeitet. Der verantwortliche Betreiber prüft die schulisch geltenden Vorgaben für Auftragsverarbeitung und Aufbewahrung.

Deployment, Infrastruktur, Zugangswiderruf und Tests: [cloudflare-deployment.md](cloudflare-deployment.md).
