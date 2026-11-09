# Lokaler KI-Assistent

Der Assistent ist optional. Repetierer startet und arbeitet ohne Ollama wie bisher. In **Einstellungen → Lokale KI** richtet **Lokale KI einrichten** Ollama und das Modell `qwen3:4b-instruct` ein. Danach ist der Chat über **KI** oder **Ctrl+K** erreichbar.

## Architektur

- `src/ai/ollamaService.js`: lokale Ollama-API (`127.0.0.1:11434`), Windows-Installation, Modell-Download mit Fortschritt und Funktionstest. Der Modellname ist zentral definiert.
- `src/ai/assistantService.js`: zentraler Systemprompt, Chatverlauf der laufenden App-Sitzung, nativer Tool-Call-Loop mit maximal zwölf Modellschritten und Unterbrechung vor jeder Schreibaktion.
- `src/ai/toolRegistry.js`: explizite Tool-Liste, Argument- und ID-Prüfung sowie Bestätigungsdaten. Die Registry verwendet vorhandene Programm-, Excel- und Session-Methoden.
- `src/mainProcess.js`: Electron-IPC, App-Kontext, Navigation und bestätigte Schreibaktionen. Nur der Main-Prozess kann Datenservices aufrufen.
- `public/aiRenderer.js`, `public/index.html`, `public/style.css`: Chat, Fortschritt, Tool-Aktivität und Bestätigungskarten.
- `src/program.js` und `src/excelReader.js`: schmale Lesezugänge für Klassen und einzelne Notenhistorien. Noten werden weiterhin über `program.saveGrade`, Mitarbeit über `teachingSessions.addParticipation` gespeichert.

## Tools

Lesen: `searchStudents`, `getStudent`, `getStudentGrades`, `getStudentParticipation`, `getClassStudents`, `getCurrentSession`, `getCurrentAppContext`.

Navigation: `openStudent`, `openClass`, `openEvaluation`, `openCurrentSession`.

Schreiben, immer mit Bestätigung: `createRepetitionGrade`, `addParticipationPoints`. Die Bestätigung ist zentral in der Agent-Schicht angeordnet, damit später eine explizite Richtlinie für einzelne Aktionen ergänzt werden kann.

Schüler-IDs sind innerhalb einer Klasse gültig; jede Aktion benötigt daher `classId` und `studentId`. Schreibaktionen verlangen eine eindeutige vorherige Namenssuche oder den aktuell ausgewählten Schüler im App-Kontext. Noten müssen zwischen 1 und 6 liegen; Mitarbeitspunkte sind ganze Zahlen von 1 bis 6. Schreibaktionen verlangen eine aktive Sitzung der gewählten Klasse und einen anwesenden Schüler. Die Repetition ist durch das bestehende Excel-Format auf sechs Noten pro Person begrenzt.

## Einrichtung

1. Die App prüft die lokale API und die Modellliste.
2. Unter Windows startet sie eine vorhandene Ollama-Installation oder lädt den offiziellen `OllamaSetup.exe` von `ollama.com` herunter und installiert ihn für das Benutzerkonto. Die Installation läuft ohne Terminal. Die App prüft die Erreichbarkeit danach erneut.
3. Die App lädt das Modell über Ollamas `/api/pull` und zeigt den gemeldeten Fortschritt. Bei erneuter Einrichtung kann ein unterbrochener Download fortgesetzt werden.
4. Ein kurzer lokaler `/api/chat`-Aufruf prüft das Modell. Erst dann zeigt die UI **Bereit**.

Für die Installation werden konservativ 7 GB freier Speicher geprüft, für einen separaten Modell-Download 3 GB. Netzwerk-, Installations-, Verbindungs-, Speicher- und Modellfehler erscheinen im Einstellungsbereich. Auf anderen Betriebssystemen muss Ollama derzeit bereits installiert sein; ein erreichbarer Dienst kann auch dort für den Modell-Download verwendet werden.

## Sicherheitsgrenzen und Datenschutz

Der Assistent benutzt ausschliesslich die lokale Ollama-API. Schülerdaten werden für die KI nicht an eine Cloud-LLM-API gesendet. Der Prompt enthält nur den kleinen UI-Kontext; weitere Daten holt das Modell einzeln über freigegebene Tools. Unbekannte Tools und zusätzliche oder ungültige Argumente werden abgewiesen. Es gibt kein SQL-, Shell-, JavaScript- oder generisches Dateitool. Das Modell erhält keinen direkten Datenbank- oder Excel-Zugriff. Jede Änderung wird vor dem Speichern in der UI bestätigt und unmittelbar davor erneut validiert.

## Einschränkungen

- Die App hat keine eigene Schülerdetailseite. `openStudent` wählt die Klasse in der vorhandenen Navigation und zeigt eine kompakte Schülerkarte im Assistenten.
- Repetitionsnoten stammen aus der Excel-Datei; Session- und Mitarbeitshistorie stammen aus dem lokalen Session-Store. Die Datenquellen können bei älteren oder ausstehenden Excel-Einträgen voneinander abweichen.
- Die Qualität natürlicher Sprachbefehle hängt vom lokalen Modell ab. Bei mehrdeutigen Namen soll der Assistent nachfragen; die Registry lässt ohne eindeutige Auflösung keine Schreibaktion zu.
- Der Chatverlauf bleibt nur während der laufenden App-Sitzung im Speicher.

## Tests

`npm test` prüft Tool-Registry, IDs und Argumente, eindeutige und mehrdeutige Schülernamen, Excel-Notenlesen, Bestätigung und Abbruch von Schreibaktionen, fehlendes Ollama/Modell sowie den Agent-Loop mit mehreren Tool-Calls. `node scripts/ai-live-smoke.js` kann bei lokal installiertem Zielmodell zehn typische Sprachbefehle mit ausschliesslich synthetischen Schülerdaten gegen die echte Ollama-API prüfen; dabei werden vorgeschlagene Schreibaktionen abgebrochen.
