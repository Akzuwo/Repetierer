# UI-Referenzen und eingebundene Komponenten

## Aktuelle Kristall-Integration (23. Oktober 2026)

Die nachfolgenden Abschnitte dokumentieren frühere UI-Arbeiten. Aktuell gilt die vom Nutzer gelieferte [Kristall-Spezifikation](kristall-material.md), mit der Korrektur **ohne Foto**. Die frühere DESIGN.md ist in [design-before-kristall.md](design-before-kristall.md) archiviert. Das Material verwendet nun die native Systemschrift; die alte Geist-Datei bleibt für bestehende Artefakte erhalten.

`design-reference/kristall/` enthält die fünf unverändert extrahierten Referenzdateien. Der lokale Referenz-Test liefert für deren Hintergrund-URL die selbst geschriebene abstrakte SVG-Szene aus. Es gibt keine Fotoabhängigkeit und keinen externen Asset-Aufruf. Die Referenz ist kein zweiter App-Einstieg und wird nicht paketiert.

`public/kristall/` enthält die CommonJS-Portierung, die gemeinsame Federuhr und die DOM-Anbindung. Shader und Federkonstanten bleiben erhalten. Ergänzt wurden Lifecycle-Cleanup, dynamische Controls, Clipping aller Scroll-Vorfahren und die Synchronisierung mit den vorhandenen GSAP-Dialogen. Beim Schliessen wird die Hintergrundebene vor der Fokus-Rückgabe wieder eingeblendet.

`npm run test:kristall` prüft die echte WebGL-Ausgabe mit neutraler Testtextur, sichtbaren Scrollbalken und DPR 1/1.25/2; dazu Navigation, Regler, Schalter, reduzierte Bewegung, Kontextverlust/Wiederherstellung sowie 680×560 und 390×844. Die Produkt-UI-Prüfung verwendet isolierte IPC-Fixtures und belegt keine externen Mail- oder Serverdienste.

Das bestehende dunkle Farbtheme und die Vanilla-JavaScript/Electron-Architektur bleiben erhalten. Alle Animationen laufen lokal; es werden keine Skripte von CDNs geladen.

## Repetitions-Hierarchie

Referenz: [Jonas-Luetolf/Repetierer](https://github.com/Jonas-Luetolf/Repetierer/tree/a3e3df9388bbce7fc6b46d9ef9d2a577b941dbef).

Die aktive Session orientiert sich am ursprünglichen Aufbau: Datei und Klassen als flache Navigation ohne Cards links, Repetition rechts mit Klassenname im Titel, Start/Freiwillig darunter, grossem Schülernamen und Bewertung am unteren Rand. Mitarbeit und Sessionaktionen bleiben über eine zurückhaltende Tab-Leiste zugänglich. Die vorhandene Klassenauswahl wird zwischen Startansicht und Navigation verschoben; sie existiert nur einmal im DOM.

## Thinking Orbs

- Quelle: [Jakubantalik/thinking-orbs](https://github.com/Jakubantalik/thinking-orbs).
- Paket: `thinking-orbs@0.3.2`, MIT; Lizenz: `public/vendor/thinking-orbs.LICENSE.txt`.
- Verwendet wird der originale Export `thinking-orbs/engine` mit `MODE_DRAWS` und `resolvePreset`. Die Canvas-Geometrie und Presets werden unverändert verwendet. Der lokale Adapter in `public/visualFeedback.js` ersetzt nur die React-Bindung.
- Zustände: `breathing` im leeren Chat, `composing` während der Antwort, `searching` bei Schülersuche, `solving` beim Vorbereiten von Bewertungen und `working` beim Lesen anderer Daten.
- Die Darstellung verwendet die originalen hellen Punkte für das dunkle App-Theme. Canvas-Auflösung berücksichtigt die Pixeldichte bis 2×. Unsichtbare/inaktive Orbs pausieren; reduzierte Bewegung ergibt ein statisches Bild. Entfernte Orbs räumen Observer und Listener auf.
- React ist lediglich ein von npm aufgelöster Entwicklungs-Peer des Pakets; es wird weder in das UI-Bundle importiert noch vom Renderer benötigt.

## GSAP

- Quelle: [GSAP](https://github.com/greensock/GSAP), [Timeline-Dokumentation](https://gsap.com/docs/v3/GSAP/Timeline/).
- Paket: `gsap@3.15.0`, Standard No Charge License; Hinweis: `public/vendor/gsap.LICENSE.txt`, originale Copyright-/Lizenzheader im Bundle.
- `public/uiMotion.js` verwendet Vanilla-GSAP-Timelines mit `expo.out` für Dialoge, Drawer, Ansichtswechsel, Schülernamen, Nachrichten und Minimieren/Wiederherstellen. Ein- und Ausblendungen dauern 120–280 ms; Ausgänge sind kürzer als Eingänge. Klassen und Anwesenheitslisten bleiben statisch. AutoAnimate wurde ersetzt.
- Pro Interaktionselement existiert eine Timeline. Ein neuer Auftrag beendet die vorherige Timeline, sodass alte Exit-Callbacks keinen wieder geöffneten Dialog ausblenden. Bei nachträglich aktivierter Bewegungsreduktion werden laufende Übergänge korrekt abgeschlossen. Geschlossene Drawer sind `inert` und aus dem Accessibility-Tree ausgeblendet.
- Die bestehenden optionalen Auswahl-/Feier-Effekte, Skeleton-Shimmer und die eigenständige Splash-Animation behalten ihre passenden CSS-/Canvas-Renderer. Routineanimationen aus dem zusätzlichen Tailwind-Layer wurden entfernt, damit App-Übergänge ausschliesslich GSAP gehören.

## Angewandte Design-Skills

- [Taste Skill: Redesign existing projects](https://github.com/Leonxlnx/taste-skill/blob/18dfc928b135629e0eddfdd445a06400d04ed439/skills/redesign-skill/SKILL.md), im Internet recherchiert und unter `C:/Users/timow/.codex/skills/redesign-existing-projects/SKILL.md` installiert. Angewandt wurde der Ablauf **Scan → Diagnose → Fix** mit dem vorhandenen Vanilla-Stack.
- [Anthropic frontend-design](https://github.com/anthropics/skills/blob/683bc88e56f3e09ba94f7055977f3d3aa499f202/skills/frontend-design/SKILL.md), installiert unter `C:/Users/timow/.codex/skills/frontend-design`.
- [Impeccable 4.5.1](https://github.com/pbakaus/impeccable/blob/9dad388a41944a0d2b8d1fb547c8556d2ecb49e7/plugin/skills/impeccable/SKILL.md), installiert unter `C:/Users/timow/.codex/skills/impeccable`; angewandt: `distill`, `animate`, Craft Floor und mechanischer Design-Scan.
- Diagnose der bisherigen Oberfläche: konkurrierende Titelgrössen, durchgehend grosse/runde Controls, Container um Container in der Startansicht, ein dekorativer Hintergrund-Glow, Pill-Tabs, grossgeschriebene Schülernamen und eine marketingartige Chat-Einleitung. Ein zweiter CSS-Layer fügte erneut Schatten und Routineanimationen hinzu.
- Umsetzung im Modus **Operate**: feste Typografieskala, lokal eingebundene Geist-Schrift, flache Klassen-/Mitarbeits-/Rundenzeilen, kompakter App-Rahmen, unterstrichene Bereichstabs und ein klarer Namensfokus. Dialoge bleiben zentriert. Die dunklen/blauen Farb-Tokens bleiben erhalten. Der KI-Dialog verwendet konkrete Aufgaben und eine originale Thinking Orb; sein dekorativer Border Beam wurde entfernt.
- Der mechanische Scan ergänzt die Prüfung der berechneten Styles und Screenshots. Er erfasst auch weiterhin vorhandene optionale Feier-/Splash-Effekte und deren separate Roboto-Schrift. Die sichtbaren bearbeiteten Ansichten werden zusätzlich bei 1100×800 und 680×560 in Electron geprüft.

## Professionelle Referenzen und Schrift

- [Linears UI-Redesign](https://linear.app/now/how-we-redesigned-the-linear-ui): Orientierung an kompakter App-Chrome, konsistenten Ausrichtungen, dichter Navigation und einer ruhigen Inhaltsfläche. Die veröffentlichte Produktoberfläche wurde visuell geprüft. Die Repetierer-Funktionen und die ursprüngliche Datei/Klassen-Hierarchie bleiben erhalten.
- [Vercel Geist](https://vercel.com/geist/typography): Orientierung an eindeutigen Schriftrollen und zurückhaltenden Komponenten.
- Schrift: [vercel/geist-font](https://github.com/vercel/geist-font/tree/10dc7658f13c38a474cde201bb09a4617267545b), Datei `fonts/Geist/webfonts/Geist[wght].woff2`, lokal als `public/fonts/Geist-Variable.woff2` eingebunden. SIL Open Font License 1.1 liegt unter `public/fonts/Geist-OFL.txt`. Das Packaging nimmt beide Dateien über das bestehende `public/fonts/**/*` auf. Der Renderer benötigt keine Netzwerkverbindung zum Laden der Schrift.

## Build und Prüfung

`npm run build:ui-vendor` bündelt nur Orb-Engine und GSAP nach `public/vendor/ui-motion.js` und kopiert die Orb-Lizenz. Das generierte Bundle ist eingecheckt und wird zusammen mit den Lizenzhinweisen über `build.files` verpackt. Zum Aktualisieren der Komponenten zuerst ihre festgelegten Paketversionen ändern und das Bundle neu bauen.

`npm run test:ui` prüft den echten Electron-Renderer mit kontrollierten IPC-Antworten: alle 21 Modals bei zwei Fenstergrössen, fehlende Sidebar-Cards und sekundäre Kontextlabels, GSAP-Tweens, unterbrochene Drawer-/Modal-Ausgänge, Minimieren/Wiederherstellen, Repetitions-Hierarchie und Namensskalierung, bemalte Canvas-Orbs mit den tatsächlichen Aktivitätszuständen, Animation/Pause und reduzierte Bewegung einschliesslich eines Wechsels während der Animation. Themes und bestehende Kernabläufe werden ebenfalls geprüft. Screenshots entstehen unter `.ui-polish-smoke/`. Die Prüfung benötigt weder echte Schülerdaten noch eine laufende Ollama-Instanz.
