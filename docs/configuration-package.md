# Ein Paket für eine Lehrperson

In den Einstellungen unter **Integrationen** zuerst Cloudflare API und SMTP einrichten. Die bisherige `.env` lässt sich über den gemeinsamen Import weiterhin auswählen.

Mit **Paket exportieren** entsteht eine einzelne `.repetierer`-Datei mit beiden Verbindungen. Diese Datei enthält den Lehrerzugang zur Umfrage-API und die SMTP-Zugangsdaten; nur an die vorgesehene Lehrperson weitergeben. Betriebssystemgebundene Verschlüsselung wird für das transportierbare Paket im Hauptprozess geöffnet; Schülerdaten und lokale Sessions werden nicht exportiert.

Die empfangende Lehrperson öffnet **Einstellungen → Integrationen → .repetierer importieren** und wählt die Datei. Die App prüft beide Konfigurationsgruppen, testet den vorhandenen API-Zugang und speichert erst danach atomar. Der Lehrerzugang wird auf dem Zielgerät erneut mit dem Betriebssystem-Schlüsselspeicher geschützt. Die importierte `.repetierer`-Datei bleibt erhalten.

Ein unvollständiges, beschädigtes oder nicht unterstütztes Paket ersetzt die vorhandene Konfiguration nicht. Die getrennten Verbindungstests für API und SMTP bleiben in den Einstellungen verfügbar.
