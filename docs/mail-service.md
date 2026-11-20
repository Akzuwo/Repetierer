# Cloudmail und lokales SMTP (7.1.0)

Jede Mailfunktion prüft vor dem Versand über den vorhandenen authentifizierten Cloudzugang `GET /teacher/mail/capabilities`. Kein neues Konfigurationspaket ist nötig. Resultat- und Debug-Testmails verwenden `src/assessment/mailService.js`; weitere Mailfunktionen sollen denselben Resolver verwenden.

| Serverergebnis | Versandweg |
| --- | --- |
| `mail.supported: true`, `mail.configured: true` | Ausschliesslich Cloudflare → Brevo HTTPS; lokale SMTP-Daten werden ignoriert |
| Explizit `mail.configured: false` | Lokales SMTP; falls nicht eingerichtet, Einrichtung verlangen |
| Alter Server: `404` mit `error: not_found` am Capability-Endpunkt | Bisheriges lokales SMTP |
| Keine Cloudkonfiguration vorhanden | Lokales SMTP |
| Netzwerk-, Auth-, Schemafehler, Limit erreicht oder Cloudversand fehlgeschlagen | Fehler, kein SMTP-Fallback |

Die Einstellungen zeigen den Versandweg. Die authentifizierte Cloudverbindung und die Mailfähigkeiten werden beim Appstart und danach alle fünf Minuten automatisch geprüft, auch bei minimierter App. Dafür wird keine Mail verschickt. Die Einstellungen zeigen den letzten Status ohne manuellen Testbutton; Konfigurationsänderungen lösen sofort eine neue Prüfung aus. Cloud zeigt die gemeinsame Restquote; alle lokalen SMTP-Aktionen und das lokale Limit sind ausgeblendet. „Konfiguration entfernen“ steht am Ende der globalen Einstellungen über die volle Breite. Einstellungen werden direkt gespeichert, ohne Speichern-/Verwerfen-Buttons. Eigenes lokales SMTP erlaubt ein Tageslimit von 1 bis 1000. Vorhandene lokale SMTP-Zugangsdaten werden behalten, bis die Lehrperson sie ausdrücklich entfernt. Neue Pakete aus einer Cloudmail-Installation enthalten nur portable Cloudzugangsdaten; 7.1.0 importiert weiterhin alte Pakete mit Cloud und SMTP. Originaldateien bleiben unverändert. Cloudpakete müssen nicht in 7.0.0 importierbar sein.

## API

Basis bleibt `https://akzuwo.ch/app/services/repetierer/api` bzw. die bestehende Route auf `timolu.ch`. Beide Endpunkte verlangen den Lehrer-Bearer-Token.

- `GET /teacher/mail/capabilities` liefert `{success:true, mail:{supported:true, configured:true/false, dailyLimit:200, used, remaining, resetTimezone:"UTC", scope:"shared"}}`.
- `POST /teacher/mail/send` erwartet `{idempotencyKey:"UUID", message:{to:"einzelne Adresse", subject, text?, html?, attachments?:[{filename:"Akte.pdf",content:"Base64"}]}}`.

Maximal ein Empfänger und ein PDF-Anhang (Base64 höchstens 7 MB), Anfrage höchstens 7,5 MB. Absender ist serverseitig fest `Philosophie Mündlich beurteilung <no-reply@akzuwo.ch>`. Keine Dateipfade, URL-Anhänge, CC/BCC oder Provideroptionen. Erfolg bestätigt die Annahme durch Brevo, nicht die Zustellung im Postfach.

Alle Installationen und Lehrerzugänge der bereitgestellten Cloudkonfiguration teilen **200 Versandversuche pro UTC-Kalendertag**. Ein D1-Batch reserviert nur bei freier Quote; ein Trigger erhöht den gemeinsamen Zähler in derselben Transaktion. Neue Installationen erzeugen keine neue Quote. Fehler und unklare Zustellungen bleiben vorsichtshalber gezählt. Das Limit wird nicht vom Client übergeben. Direkter Alt-SMTP-Versand aus 7.0.0 kann zusätzlich Brevos Kontingent verbrauchen und lässt sich durch den Worker nicht begrenzen.

Aufträge speichern Lehrer-ID, UUID, Payload-Hash, Tag, Status und Provider-ID, keine Mailinhalte oder Empfängeradresse. Derselbe Auftrag wird höchstens einmal an Brevo weitergereicht. Erfolgreiche Wiederholungen liefern `replayed:true`; geänderter Inhalt unter derselben ID wird abgewiesen. Transportabbruch, Provider-5xx oder Prozessabbruch erlauben keinen blinden Neuversand. Vor weiteren Aktionen den Versandverlauf prüfen. Definitiv abgelehnte Provideraufträge erhalten erst durch ausdrückliche Wiederholung in der App eine neue ID; die vorherige Reservierung bleibt gezählt. Resultatmails speichern ihre ID vor dem ersten Versand dauerhaft in der Bewertung.

Lokales SMTP reserviert sein einstellbares Kontingent vor dem Versand in `smtp-mail-usage.json`, mit gemeinsamer Warteschlange im App-Hauptprozess. Dies begrenzt keine anderen Programme und ersetzt keine Cloudquote.

## Betrieb und Tests

Brevo-HTTPS-Schlüssel als Worker-Secret `BREVO_API_KEY`, niemals im Electron-Paket oder neuen Cloud-Konfigurationspaket. Lokale Betreiberdatei: `.cloudflare-private/brevo.env`, durch `.gitignore` ausgeschlossen. Schema: `cloudflare/migrations/0002_mail.sql`. Bestehende Endpunkte und Antwortformate bleiben erhalten. [Brevo-HTTPS-API](https://developers.brevo.com/reference/send-transac-email).

Nur Staging hat `MAIL_SANDBOX=true`. Der Worker setzt selbst `headers.X-Sib-Sandbox=drop` im Payload; Clients können den Modus nicht wählen. Produktion hat diese Variable nicht. Der [Brevo-Sandboxmodus](https://developers.brevo.com/docs/using-sandbox-mode) validiert ohne Zustellung und beweist keine echte Zustellung.

Prüfungen: `node --test test/mailService.test.js test/delivery.test.js test/integrationConfig.test.js test/assessmentApiClient.test.js`, privater D1-Test `node --test test/cloudflareMail.test.js`, Desktop `node scripts/api-mail-ui-check.cjs`. Infrastruktur und Betreiberwerkzeuge bleiben unversioniert. Die lokale Lese-API überträgt keine Schülerdaten in die Cloud.
