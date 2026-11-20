# Cloud-API für Selbstbeurteilungen (7.1.0)

Diese API veröffentlicht Umfragerunden und liefert Schülerantworten. Sie ist von der lokalen, ausschliesslich lesenden API getrennt. Cloud-Schreibaktionen verändern Online-Runden, keine lokale Exceldatei. Automatisierungen sollten zuerst lesend arbeiten; insbesondere `DELETE` entfernt eine Online-Runde samt Antworten endgültig.

## Basis und Anmeldung

Basis: `https://timolu.ch/app/services/repetierer/api`, alternativ `https://akzuwo.ch/app/services/repetierer/api`. Die bereitgestellten Zugänge gelten für beide Domainrouten desselben Dienstes. Lehrerendpunkte benötigen `Authorization: Bearer <Lehrer-Token>`; der Token stammt aus der persönlichen Cloudkonfiguration. Den Token nicht in URLs, Ausgaben oder eigene Dokumentationen schreiben. Die App speichert ihn mit dem Betriebssystem-Schlüsselspeicher. Schülerendpunkte verlangen den vollständigen Rundenlink, keinen Lehrer-Token.

JSON-Anfragen verwenden `Content-Type: application/json`. Erfolgreiche Antworten enthalten `success:true`; Fehler enthalten `success:false` und einen maschinenlesbaren `error`-Code. Alle Pfade der folgenden Tabelle sind relativ zur Basis.

## Endpunkte

| Methode / Pfad | Zugriff | Ergebnis und Verwendung |
| --- | --- | --- |
| `GET /health` | Öffentlich | Erreichbarkeit von Dienst und Datenbank; `service` bezeichnet den Dienst. |
| `GET /teacher/me` | Lehrer | Anmeldung prüfen; liefert `teacherId`. |
| `PUT /teacher/rounds/:id` | Lehrer | Neue Runde unter eigener UUID erstellen; liefert `roundId`. Gleiche ID mit gleichen Angaben ist wiederholbar; andere Angaben ergeben `409 round_conflict`. |
| `GET /teacher/rounds/:id/responses?cursor=0` | Eigentümer | Liefert `responses`, `closed`, `nextCursor`. Seiten enthalten höchstens 100 Antworten. Mit `nextCursor` weiterlesen, bis der Wert `null` ist. |
| `POST /teacher/rounds/:id/close` | Eigentümer | Runde wiederholbar schliessen; liefert `closed:true`. |
| `POST /teacher/rounds/:id/reopen` | Eigentümer | Mit neuer Frist wieder öffnen; vorhandene Antworten und Duplikatsperren bleiben erhalten. Liefert `closed:false` und `deadline`. |
| `DELETE /teacher/rounds/:id` | Eigentümer | Runde mit Antworten löschen; liefert `deleted:true`. |
| `GET /public/rounds/:id` | Rundenlink | Liefert `round` mit `id`, `title`, `criteria`, `deadline`, `expiresAt`, `closed`; keine Antwortliste anderer Personen. |
| `POST /public/rounds/:id/responses` | Rundenlink | Einzelne Selbstbeurteilung abgeben; liefert `responseId`. Namen und Browsercode sind pro Runde gegen Duplikate geschützt. |
| `GET /teacher/mail/capabilities` | Lehrer | Cloudmail-Fähigkeit und gemeinsames Tageskontingent; siehe Cloudmail-Abschnitt. |
| `POST /teacher/mail/send` | Lehrer | Einzelne Mail mit wiederholbarer Auftrags-ID; siehe Cloudmail-Abschnitt. |

## Anfragen und Antwortdaten

Erstellen einer Runde (`PUT`), mit selbst erzeugter UUID als `:id`:

```json
{
  "title": "Selbstbeurteilung",
  "criteria": [
    {"id":"beteiligung","label":"Beteiligung","scaleLabels":["Nie","Selten","Manchmal","Oft","Immer"]}
  ],
  "deadline": "<zukünftiger ISO-Zeitpunkt mit Zeitzone>",
  "maxResponses": 25
}
```

`title` hat höchstens 160 Zeichen. Eine Runde enthält 1–20 Kriterien mit eindeutigen IDs aus Kleinbuchstaben, Ziffern und Unterstrich (höchstens 64 Zeichen), je einer Beschriftung (höchstens 250 Zeichen) und genau fünf Skalenbeschriftungen (je höchstens 80 Zeichen). `maxResponses` liegt zwischen 1 und 200. Die Frist muss beim Erstellen in der Zukunft und höchstens 90 Tage entfernt liegen; ohne Frist verwendet der Server 30 Tage. Ein Lehrerzugang kann höchstens 200 aktive Runden besitzen.

Wiederöffnen (`POST /teacher/rounds/:id/reopen`) erwartet `{"deadline":"<zukünftiger ISO-Zeitpunkt mit Zeitzone>"}` innerhalb der nächsten 90 Tage. Die App berechnet Fristen als Tagesende in Europe/Zurich. Die öffentlichen Zeitangaben sind ISO-Zeitpunkte. Eine bereits abgelaufene und endgültig entfernte Online-Runde kann nicht wieder geöffnet werden.

Abgabe (`POST /public/rounds/:id/responses`):

```json
{
  "name":"Synthetische Testperson",
  "email":"person@example.invalid",
  "browserToken":"<zufälliger Browsercode aus 32–100 Zeichen>",
  "answers":{"beteiligung":4},
  "additionalQuality":"",
  "comment":""
}
```

Der Browsercode enthält nur Buchstaben, Ziffern, Unterstrich oder Bindestrich. `answers` enthält genau die Kriterien der Runde mit ganzzahligen Werten von 1 bis 5. Name höchstens 200 Zeichen, Mailadresse höchstens 254 Zeichen; optionale Texte jeweils höchstens 2000 Zeichen. Die Anfrage ist höchstens 32 KiB gross. Fremde Browser-Origins sind für Abgaben gesperrt; Lehrerzugriffe haben keine CORS-Freigabe.

Eine Lehrer-Antwortseite enthält je Eintrag `responseId`, `roundId`, `timestamp`, `name`, `email`, `answers`, `additionalQuality`, `comment`. Lokal bereits importierte `responseId`-Werte merken und nicht erneut als neue Antwort speichern.

## Fehler und Wiederholung

`401 unauthorized` bedeutet fehlenden, abgelaufenen oder widerrufenen Lehrerzugang. Fremde bzw. abgelaufene Runden liefern `404 round_not_found`. `409 duplicate_response` verhindert eine weitere Schülerabgabe; `410 round_closed` bzw. `round_closed_or_full` verhindert Abgaben nach Schliessung/Frist oder bei voller Runde. `400` betrifft ungültige Angaben, `413` zu grosse Anfragen, `429 round_limit` zu viele aktive Runden. Netzwerkfehler sind kein Beweis dafür, dass eine Schreibaktion nicht ausgeführt wurde: Runden mit derselben UUID und unveränderten Angaben wiederholen; Mailaufträge nur mit derselben Auftrags-ID. Siehe dazu die Mail-Idempotenz im folgenden Abschnitt.
