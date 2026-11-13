# Arbeitsregeln für Repetierer

Stand: 30. Oktober 2026. Projekt: Windows-Electron-App für Lehrpersonen; die Cloudflare-Schülerumfrage ist eine getrennte Browseranwendung.

## Verbindliche Vorgaben

- Lösche nie alte Builds. Sie dürfen höchstens überschrieben werden. Keine Bereinigung von `builds/`, alten Installern oder früheren Artefakten.
- Teste jedes Ergebnis mit den für die Änderung passenden Prüfungen. Fehler beheben und erneut gezielt prüfen. Nicht geprüfte Grenzen ehrlich nennen.
- Erstelle keinen Installer und veröffentliche keine Builds oder Cloudflare-Änderungen ohne ausdrücklichen Auftrag. Der Nutzer erstellt seine Installer selbst.
- Die App ist Desktop-only. Keine mobilen App-Ansichten prüfen oder aus allgemeinen Responsive-Vorgaben einen Smartphone-Umbau ableiten. UI-Harnesses prüfen 1100 × 800 und bei Bedarf 1440 × 900. Die getrennte Schülerumfrage darf responsiv geprüft werden.
- Bestehende Nutzeränderungen, Exceldateien, Noten, Joker, lokale Einstellungen und Originale erhalten. Keine echten Mails, KI-Schreibaktionen oder fremden Datenänderungen für UI-Tests auslösen. Temporäre Profile und Fixtures verwenden.

## Gestaltung und Verhalten

Lies [DESIGN.md](DESIGN.md) und den betroffenen Quellcode vor einer UI-Änderung. Die vorhandene Kristallwelt weiterführen: abstrakte Szene ohne Foto, klare abgerundete Glaslinsen, angebundener Drawer, stationärer Header und Session-Chrome. Native Semantik, Tastaturbedienung und Reduced Motion erhalten. Keine sichtbaren Scrollbars, keine zusätzlichen Pillrahmen um reine Werkzeugicons oder ungerahmte KI-Antworten.

Transienten Erfolg, Fehler, Warnung und Information immer über `public/notifications.js` beziehungsweise im Hauptprozess `src/notifications.js` ausgeben. API: `notify.error(text, options)`, `warning`, `info`, `success`; Standard 6 Sekunden, Einstellung 2–8. Keine neuen Toastcontainer, Inline-Ergebnisbanner, eigenen Timer oder Schliessen-Buttons. Entscheidungen, Feldbeschriftungen, Ladezustände und dauerhafte Statusinformationen bleiben an ihrem fachlichen Ort. Speichererfolg erst nach bestätigtem Ergebnis melden; ausstehende Exceländerungen als Warnung behandeln.

Gemeinsame Bewegung in `uiMotion.js`, wiederholbare Scrollbewegung in `scrollEffects.js`; Material in `public/kristall/`. Optische Grenzen stets aus tatsächlichen DOM-/Canvas-Bounds berechnen. Neue Steuerelemente müssen native Maus-/Tastatureingaben akzeptieren und dürfen keine Electron-Drag-Region erben.

## Architektur und Daten

[Architekture.md](Architekture.md) beschreibt Module, IPC und Persistenz. Exceländerungen über die bestehenden Fachmodule führen. Explizite Klasse/Anfrage-ID im Editor beibehalten. Joker-Datumsspalten, Backup, Warteschlange und Undo/Redo gemeinsam berücksichtigen. Migration nur mit sichtbarer Vorschau und Sicherung; keine historischen Werte erfinden. Zugangsdaten nicht ausgeben oder in Tests/Screenshots aufnehmen. `safeStorage`, Konfigurationsvalidierung und bestätigte KI-Schreibaktionen erhalten.

## Prüfung

- Logik: `node --test --test-concurrency=1` mit den betroffenen `test/*.test.js`. `npm test` prüft den gesamten öffentlichen Appbereich. Die privaten Cloudflare-Backendtests sind nicht Teil des Git-Index.
- Hausaufgaben und Klasseneinstellungen: `npm run test:homework`. Regeln bleiben ausschliesslich im App-Profil; bestehende Notengewichte und Daten beim Deaktivieren erhalten.
- Meldungen: `npm run test:notifications`.
- Mitmachnoten: `npm run test:assessment`; laufende Session beim Schliessen: `npm run test:session-close`.
- Desktopabläufe: `npm run test:ui`; Debug-Einstieg: `npm run test:debug`.
- Änderungen am Material/Renderer: `npm run test:kristall`.
- Gezielte Abläufe: `node scripts/liquid-refinement-focused.cjs --settings-sliders`, `--editor-select`, `--assessment-polish`, `--stats-drawer`, `--ai-flight` oder `--notifications`.
- UI-Prüfungen nacheinander ausführen; native Fenster/Fokus dürfen nicht konkurrieren. `ELECTRON_RUN_AS_NODE` für Electron entfernen. `git diff --check` vor Abschluss.

Keine unbegrenzte Screenshot-Polierschleife: Änderungen sammeln, Desktopzustände gemeinsam ansehen, konkrete Fehler korrigieren und gezielt bestätigen. Tests müssen das neue Verhalten nachweisen; alte Assertions bei einem gewollten Wechsel fachlich ersetzen, nicht kommentarlos entfernen.

## Dokumentation

Bei Änderungen an Verhalten, Modulen, Tests oder Design diese Dateien aktualisieren. `DESIGN.md` ist die aktuelle Gestaltung, `Architekture.md` die technische Übersicht. Die äusseren Dateien im Workspace verweisen auf diese Projektdateien. Originale Designreferenzen und alte Brandassets unter `design-reference/` erhalten; sie sind keine aktuellere Produktspezifikation.

Cloudflare-Servercode, Deployment-Konfiguration und Betreiber-Skripte sowie lokale Zugangsdaten bleiben unversioniert. Vor einem Push `npm run test:repository` und einen Secret-Scan des tatsächlichen Index/Snapshots ausführen. Eine Ignore-Regel allein entfernt keine bereits versionierten Dateien.
