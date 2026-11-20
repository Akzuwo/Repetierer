# Repetierer

## Platform

web

Electron-Desktop für Windows mit getrenntem fensterlosem API-Prozess; responsive Webumfrage für Schüler. Die Desktopansichten werden ausschliesslich bei Desktopgrössen geprüft.

## Users and purpose

Lehrpersonen führen Repetitionen, Mitarbeit und Beurteilungsrunden mit lokalen Excel-Klassenlisten. Schüler geben eine Selbstbeurteilung ohne Konto über einen teilbaren Rundenlink ab.

## Constraints

Bewertung und Eigenkorrektur bleiben lokal. Cloudmail läuft bei eingerichteter Serverfähigkeit über Cloudflare und Brevos HTTPS-API mit gemeinsamer 200er-Tagesquote; eigenes lokales SMTP bleibt kompatibel. Eine unabhängige lokale Lese-API erschliesst gespeicherte Appdaten auf demselben Gerät. Die Umfrage folgt der bestehenden Designsprache, ist minimalistisch, mobil bedienbar und per Tastatur zugänglich. Keine IP-Sperren; eine Antwort pro Name und Runde. Alte Builds bleiben erhalten.
