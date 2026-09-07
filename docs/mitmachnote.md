# Mitmachnote

## Architektur und Datenhaltung

Die Excel-Arbeitsmappe bleibt die Quelle für Klassen, Schüler und sämtliche Repetitionsdaten. Mitmachnoten werden getrennt in `mitmachnoten.json` unter Electrons persistentem `userData`-Ordner gespeichert. Schema 3 ergänzt das Raster des ursprünglichen Selbstbeurteilungsbogens und ein separates Freitextfeld für zusätzliche Qualitäten. Alte Daten werden beim Laden migriert; bestehende Runden behalten ihr jeweiliges Raster und alle Selbst- und Lehrerbeurteilungen.

Die Selbstbeurteilung verwendet ein dauerhaftes, außerhalb von Repetierer vorbereitetes Google Form mit verknüpfter Antworttabelle. Ein dazugehöriges Google Apps Script liest die Tabelle und stellt die Antworten als kleine HTTPS-JSON-API bereit. Repetierer benötigt weder Google-OAuth-Client noch Google-Anmeldung und erstellt keine Formulare. Die Arbeitskopie bleibt lokal; ohne Internet können bereits synchronisierte Antworten und Lehrerbewertungen weiter bearbeitet werden.

## Formular vorbereiten

Das Formular muss folgende Fragen enthalten. Die Titel müssen mit der mitgelieferten Script-Vorlage übereinstimmen:

- `Beurteilungsrunde` als Pflicht-Kurzantwort
- `Vorname, Nachname` als Pflicht-Kurzantwort
- `E-Mail Adresse` als Pflicht-Kurzantwort
- `Beteiligung mit Äusserungen`
- `Falsch, unbefriedigend oder nicht ausreichend`
- `Originell, aber unpassend im Lektionsverlauf`
- `Korrekt, aber stichwortartig`
- `Passend und im Lektionsverlauf weiterführend`
- `Passend, eigenständig und in eine neue, interessante Richtung führend`
- `Zusätzliche Qualität meiner Äusserungen` als Absatz
- `Teilnahme insgesamt`
- `Sonstige Bemerkungen` als Absatz

Für die ersten sechs Bewertungsfragen gilt die Skala `1 – (Fast) nie`, `2 – Ab und zu`, `3 – Manchmal`, `4 – Häufig`, `5 – Sehr häufig`. Für `Teilnahme insgesamt` gilt `1 – Mangelhaft`, `2 – Genügend`, `3 – Recht`, `4 – Gut`, `5 – Sehr gut`. Das Script akzeptiert sowohl diese ausgeschriebenen Optionen als auch reine Zahlen von 1 bis 5.

Über **Mehr > Link zum vorausgefüllten Formular abrufen** lässt sich für `Beurteilungsrunde` ein Beispielwert einsetzen. Aus dem erzeugten Link wird der Parametername `entry.123456...` für `ASSESSMENT_ROUND_FIELD_ID` übernommen. Repetierer setzt dessen Wert pro Runde automatisch; die technische ID erscheint nicht in der Oberfläche.

## Apps Script bereitstellen

1. Das Formular mit einer Google-Tabelle für Antworten verknüpfen und in dieser Tabelle über **Erweiterungen > Apps Script** ein gebundenes Script-Projekt öffnen.
2. Den Inhalt aus [`google-apps-script/Code.gs`](google-apps-script/Code.gs) in `Code.gs` einsetzen.
3. Unter **Projekteinstellungen > Script Properties** `ASSESSMENT_API_KEY` mit einem langen zufälligen Wert hinterlegen. Optional können `SPREADSHEET_ID` und `RESPONSES_SHEET_NAME` gesetzt werden; im gebundenen Script werden sonst die aktive Tabelle und ihr erstes Blatt verwendet.
4. **Bereitstellen > Neue Bereitstellung > Web-App** wählen, Ausführung als Eigentümer festlegen und den Zugriff so konfigurieren, dass Repetierer die URL ohne Google-Anmeldung aufrufen kann.
5. Die `/exec`-URL als `ASSESSMENT_API_URL` verwenden. Bei Scriptänderungen eine neue Version der Bereitstellung veröffentlichen.

Der aktuelle Script-Vertrag akzeptiert `GET` mit `apiKey` und `roundId`. Repetierer unterstützt zusätzlich den früher dokumentierten JSON-POST-Vertrag und erkennt die Variante automatisch. Der API-Schlüssel erscheint niemals im Einladungslink oder in Repetierer-Protokollen. Die API und der Client filtern exakt nach `roundId`; Antwort-IDs verhindern lokale Doppelimporte.

## Gemeinsame Konfiguration importieren

Die importierte Datei kann Mitmachnote, Brevo oder beide vollständigen Gruppen enthalten:

```dotenv
ASSESSMENT_FORM_URL=https://docs.google.com/forms/d/e/.../viewform?usp=pp_url
ASSESSMENT_API_URL=https://script.google.com/macros/s/.../exec
ASSESSMENT_API_KEY=
ASSESSMENT_ROUND_FIELD_ID=entry.123456789

BREVO_SMTP_HOST=smtp-relay.brevo.com
BREVO_SMTP_PORT=587
BREVO_SMTP_USER=
BREVO_SMTP_PASSWORD=
BREVO_FROM_EMAIL=
BREVO_FROM_NAME=Herr Schaufelberger
```

`ASSESSMENT_ROUND_FIELD_ID` ist optional. Fehlt der Wert, erkennt Repetierer das Kurzantwortfeld `Beurteilungsrunde` beim ersten Onlinezugriff und speichert seine `entry...`-ID intern. Danach kann der Einladungslink auch offline erzeugt werden.

Repetierer parst die Datei ausschließlich als Key/Value-Konfiguration. Vorhandene Gruppen werden zusammengeführt, vollständig validiert und atomar als `repetierer.env` im persistenten `userData`-Ordner geschrieben. Nach erfolgreicher Rückleseprüfung wird erst dann die ausgewählte Originaldatei gelöscht. Bei Lese-, Validierungs-, Schreib- oder Verifikationsfehlern bleibt das Original erhalten. Eine bestehende `brevo-smtp.env` wird weiterhin gelesen und kann beim nächsten Import ohne Datenverlust in die gemeinsame Konfiguration übernommen werden.

## Synchronisation und Versand

Neue Runden erhalten eine ID wie `assessment-3a-hs-2026-a8f2d1`. Der Einladungslink enthält diese ID als vorausgefüllten Formularwert. Beim Aktualisieren sendet Repetierer die Runden-ID und bereits bekannte Antwort-IDs an die API. Fremde Runden und ungültige Datensätze werden ignoriert und gezählt. Alle gültigen Antworten werden lokal aufbewahrt; bei gleicher E-Mail wird standardmäßig die neueste verwendet, eine ältere Abgabe kann in der Rundenansicht ausgewählt werden.

Der Brevo-Verbindungstest und Klassenversand verwenden die importierten SMTP-Werte. Jede Person erhält eine getrennte HTML- und Text-Mail. Status und Versuchszahl werden nach jedem Empfänger gespeichert; **Fehlgeschlagene erneut senden** berücksichtigt nur fehlgeschlagene Mails. Passwörter, API-Schlüssel und vollständige Antworten werden nicht protokolliert oder exportiert.
