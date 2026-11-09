---
name: Kristall — Liquid Glass
description: Portierbare, ausführbare Referenz des freigegebenen Kristall-Materials aus Liquid Glass 1.3.0.
colors:
  ink: '#ffffff'
  muted: rgba(255, 255, 255, .78)
  panel-muted: rgba(255, 255, 255, .94)
  crystal-fallback: rgba(220, 243, 235, .104)
  switch-track: '#ffffff30'
  switch-on: '#96dab899'
  focus-ring: '#d5fff0'
  scene-base: '#244b47'
typography:
  body:
    fontFamily: -apple-system, BlinkMacSystemFont, 'Helvetica Neue', 'Segoe UI', sans-serif
    fontSize: 16px
    fontWeight: 400
  title:
    fontFamily: -apple-system, BlinkMacSystemFont, 'Helvetica Neue', 'Segoe UI', sans-serif
    fontSize: 22px
    fontWeight: 550
    letterSpacing: -.025em
  navigation:
    fontFamily: -apple-system, BlinkMacSystemFont, 'Helvetica Neue', 'Segoe UI', sans-serif
    fontSize: 13px
    letterSpacing: .005em
  label:
    fontFamily: -apple-system, BlinkMacSystemFont, 'Helvetica Neue', 'Segoe UI', sans-serif
    fontSize: 12px
  small:
    fontFamily: -apple-system, BlinkMacSystemFont, 'Helvetica Neue', 'Segoe UI', sans-serif
    fontSize: 11px
rounded:
  pill: 50px
  circle: 50%
  widget: 30px
  panel: 32px
  range: 22px
spacing:
  nav-inset: 5px
  gap: 12px
  widget-inset: 22px
  panel-inset: 26px
components:
  icon-button:
    textColor: '{colors.ink}'
    rounded: '{rounded.circle}'
    padding: '0'
    width: 46px
    height: 46px
  navigation:
    textColor: '{colors.ink}'
    typography: '{typography.navigation}'
    rounded: '{rounded.pill}'
    padding: '{spacing.nav-inset}'
    height: 52px
  range:
    rounded: '{rounded.range}'
    height: 38px
    width: 100%
---

# Design System: Kristall — Liquid Glass

## Overview

Diese Datei beschreibt **das vom Nutzer freigegebene Kristall-Material der lokalen Liquid-Glass-Anwendung, Version 1.3.0 vom 22. Oktober 2026**. Sie ist eine eigenständige Übergabe an Coding Agents. Die vollständige Optik- und Bewegungsimplementierung steht weiter unten direkt in dieser Datei. Das Ursprungsprojekt, ein Plugin, eine bestimmte Komponentenbibliothek und ein externer Dienst werden für die Übernahme nicht benötigt.

**Ziel:** maximal klares, sichtbar lichtbrechendes Glas mit einer einzigen sauberen Außenkontur. Navigation, Buttons, Schalter und Regler verwenden dasselbe Material. Bewegte Formen ziehen sich elastisch, reagieren unmittelbar auf Eingaben und federn dezent zurück. Der Hintergrund bleibt durch das Glas erkennbar.

Die Referenz ist eine **WebGL/HTML/CSS-Adaption**, kein natives Apple-Material und keine vollständige physikalische Glassimulation. Ihre Linsen brechen eine gemeinsame Hintergrundtextur. Sie brechen keine beliebigen HTML-Inhalte hinter einer Komponente. Ein CSS-Fallback erhält Bedienbarkeit und Grundform; er erzeugt keine gleichwertige Lichtbrechung.

### Auftrag an den Coding Agent

> Übernimm das Kristall-Material aus dieser Design.md einschließlich des eingebetteten Referenzcodes. Implementiere zuerst die Referenz unverändert und prüfe ihre Optik. Integriere sie anschließend in die bestehende Architektur. Erhalte Materialkonstanten, Koordinatenabbildung, die gemeinsame Federuhr und die weichen Reglerfelder. Passe Inhalt, Hintergrund und Anordnung an das Zielprojekt an. Ersetze den Shader nicht durch gewöhnliches verschwommenes CSS-Glas. Verwende native semantische Eingabeelemente. Prüfe gerasterte Glasflächen und HTML gemeinsam, mit sichtbaren Scrollbalken und mehreren Pixeldichten. Liefere eine funktionierende, getestete Integration.

### Was feststeht und was frei bleibt

| Verbindlich für denselben Materialeindruck | Darf das Zielprojekt bestimmen |
| --- | --- |
| Kristall-Werte, Shaderkonstanten, ein Rand, ein Schatten | Seitenaufbau, Inhalte und Navigationseinträge |
| Federverhalten, weiche Reglerfelder, Verformung | Datenmodell und Anwendungslogik |
| Korrekte CSS-/Canvas-/Pixel-Koordinaten | Framework, Dateistruktur und Komponenten-API |
| Klare Icons und Text außerhalb der Brechung | Hintergrundmotiv und produktbezogene Akzentfarben |
| Tastatur, Pointer-Abbruch und reduzierte Bewegung | Ob Musik und Wasservisualisierung benötigt werden |

Das Hintergrundmotiv verändert den sichtbaren Farbeindruck stark. Das gleiche Material über einer anderen Szene sieht deshalb nicht pixelidentisch aus. Für einen Pixelvergleich müssen auch Bild, Ausschnitt, Helligkeit, Vignette, Schriftumgebung und Zeigerposition übereinstimmen.

### Schnellstart mit dieser einen Datei

1. Extrahiere die fünf ausdrücklich benannten Referenzdateien aus den Codeblöcken unter **Components**: `index.html`, `styles.css`, `app.js`, `glass.js` und `motion.js`.
2. Lege ein eigenes Hintergrundfoto unter `assets/background.jpg` daneben. Verwende ein lokal verfügbares, ausreichend aufgelöstes Bild.
3. Starte einen lokalen HTTP-Server in diesem Ordner. Öffne die Seite über HTTP; die ES-Module nicht per `file://` starten.
4. Prüfe Navigation, große und kleine Schalter, breite und kompakte Regler, Buttons sowie Wasserbewegung.
5. Übernimm die geprüften Klassen beziehungsweise deren Logik in das Zielprojekt. Das kleine Beispielraster ist nur ein Prüfgerüst und keine vorgeschriebene Seitenkomposition.

## Colors

Die Frontmatter enthält die verbindlichen CSS-Farben. Die breite Farbwirkung kommt aus dem Hintergrund. Text und Icons bleiben weiß; eingeschaltete Schalter erhalten eine transparente, blassgrüne Zustandsfarbe.

### Materialparameter

UI-Prozentwerte und normalisierte Shaderwerte nicht verwechseln:

| Parameter | UI-Wert | Rendererwert |
| --- | ---: | ---: |
| Lichtbrechung / `refraction` | 100 % | `1.0` |
| Mattierung / `frost` | 0 % | `0.0` |
| Glasdichte / `tint` | 8 % | `0.08` |
| Hintergrundhelligkeit / `brightness` | 100 % als Referenz | `1.0` |
| Zusätzlicher Kontrast / `contrast` | aus | `0.0` |
| Fokusabdunklung / `focus` | aus | `0.0` |
| Zeigerlicht / `motion` | an | `true` |

Kristall ist der Standard für **alle** Glasflächen, auch für bewegte Marker und Schalterknöpfe. Nicht jeder Komponente einen etwas anderen Blur-, Tint- oder Refraction-Wert geben. Ein bewusster Kontrastmodus ist ein zusätzlicher Zustand, nicht der normale Kristall-Look.

Die Fallback-Oberfläche entspricht `alpha = .08 + 8 × .003 = .104`. Diese Alpha-Zahl ist **kein** Ersatz für den Shader und wird nicht zusätzlich über dessen Glas gelegt. Im WebGL-Modus sind CSS-Flächenfüllung und Backdrop-Filter transparent beziehungsweise deaktiviert.

## Typography

Die native Systemschrift aus der Frontmatter verwenden. Sie hält die Umgebung ruhig, während das Material sichtbar wird. Navigation nutzt die Navigation-Rolle, Feldnamen die Label-Rolle und Zahlen die Small-Rolle mit `font-variant-numeric: tabular-nums`.

Die Referenz enthält kompakte 11px-Ausgaben. Das ist eine gemessene Größe dieser Vorlage, keine allgemeine Empfehlung für Fließtext. Der Lesbarkeitsbedarf des Zielprodukts darf größere Texte verlangen. Icons bleiben scharfe SVGs über der optischen Ebene; weder Text noch Icons werden durch den Hintergrundshader verzerrt.

## Layout

### Ebenenvertrag

| Ebene | Aufgabe |
| --- | --- |
| Hintergrundfoto / CSS-Fallback | `position: fixed; inset: 0; width: 100%; height: 100%` |
| Ein WebGL-Canvas | Dasselbe Foto und danach alle Linsen, `z-index: 1`, `pointer-events: none` |
| Optionale Vignette | `z-index: 2`, ebenfalls ohne Pointer-Ereignisse |
| Semantisches HTML | `z-index: 3` oder höher; tatsächliche Buttons, Labels und Inputs |

Nur **ein** gemeinsames Canvas verwenden. Es misst die echten DOM-Rechtecke über `getBoundingClientRect()`. Das Canvas liegt in der Referenz am Viewport-Ursprung `(0, 0)` und ist nicht selbst transformiert. Wird es im Zielprojekt versetzt oder in einen transformierten Container gelegt, müssen alle DOM-, Zeiger- und Clip-Koordinaten in sein lokales Koordinatensystem umgerechnet werden.

### Die Scrollbalken-Regel

`window.innerWidth` darf nicht als CSS-Zeichenbreite vorausgesetzt werden. Ein klassischer Scrollbalken kann beispielsweise `innerWidth = 1391` und eine tatsächliche Canvasbreite von `1385` erzeugen. Wird das Canvas trotzdem mit 1391 CSS-Pixeln berechnet, rutschen die Glasflächen zunehmend gegenüber ihren Icons nach links.

Die verbindliche Abbildung lautet:

```js
const bounds = canvas.getBoundingClientRect();
const requestedDpr = Math.min(devicePixelRatio || 1, 1.5);
canvas.width = Math.round(bounds.width * requestedDpr);
canvas.height = Math.round(bounds.height * requestedDpr);
const pixelRatioX = canvas.width / bounds.width;
const pixelRatioY = canvas.height / bounds.height;
```

Im Fragmentshader wird jeder Pixel mit der **effektiven Achsenauflösung** zurückgerechnet:

```glsl
vec2 p = vec2(
  gl_FragCoord.x / u_pixelRatio.x,
  u_resolution.y - gl_FragCoord.y / u_pixelRatio.y
);
```

Auch Scissor-Rechtecke verwenden diese beiden Verhältnisse. Canvasgrößen per `ResizeObserver` beobachten, einschließlich Änderungen beim Ein- und Ausblenden von Scrollbalken. Keine Canvasdimension erneut setzen, wenn sie gleich geblieben ist: Das würde den Drawing Buffer unnötig löschen.

### Framework- und Scrollintegration

- DOM-Rechtecke erst nach dem Mounten beziehungsweise nach dem Layout messen. Ein externes DOM erst nach dem Laden der Client-Komponenten initialisieren.
- Bruchteilige Maße erhalten. `offsetWidth` und `offsetLeft` nicht zur alleinigen Zielgeometrie machen; deren Rundung kann Marker um halbe Pixel versetzen.
- Die Markerziele bei skalierten Eltern in deren unskaliertes lokales Koordinatensystem zurückrechnen. Der Referenzcontroller erledigt das in `follow()`.
- Scroll-Ereignisse auch aus inneren Scrollcontainern erfassen. Der Referenzrenderer nutzt dafür `capture: true`.
- Canvas-Clipping muss HTML-Overflow-Clipping entsprechen. Die Referenz behandelt `.home-view` innerhalb von `main` bei `max-width: 760px`. Bei anderen Scrollcontainern diesen Clip-Vertrag ausdrücklich anpassen.
- Die Referenz sammelt die Glasflächen bei jedem Draw neu. Nach dynamischem Einfügen neuer Größenquellen zusätzlich den ResizeObserver registrieren und die Marker initialisieren.
- Die Referenz besitzt globale Listener und einen Renderer. In React/Vue/Svelte nicht bei jedem Render neue Instanzen anlegen. Beim Unmounten Listener, Observer und Frames aufräumen; dazu dem Renderer beim Portieren einen passenden `dispose()`-Lebenszyklus ergänzen. Die eingebettete Referenz ist für eine einmalig gemountete Seite geschrieben.

## Elevation & Depth

### Eine Kontur, ein Schatten

Der einzige normale Umgebungsschatten lautet:

```css
box-shadow: 0 10px 28px -12px rgba(0, 0, 0, .28);
```

Beim Icon-Hover darf dessen Alpha auf `.36` steigen. Bewegte Marker, Range-Knöpfe und Toggle-Knöpfe erhalten **keinen eigenen CSS-Schatten**. Keine Ketten aus Inset-Shadows, weißen Außenringen und mehreren halbtransparenten Duplikaten erzeugen. Im WebGL-Modus gehört der optische Rand vollständig dem Shader. Der CSS-Fallback besitzt stattdessen einen einzelnen weißen 1px-Rand.

### Verbindliche Shaderkonstanten

Der vollständige Shader steht in der Referenzdatei `glass.js`. Folgende Zahlen nicht durch vermeintlich bessere Standardwerte ersetzen:

| Wirkung | Konstante / Formel |
| --- | --- |
| Antialiasing der SDF-Kontur | `1 - smoothstep(-0.8, 0.5, sdf)` |
| Bevelbreite | `min(radius * 0.85, 24.0)` |
| Randprofil | `lens = sin(bevel * 1.570796)` |
| Einwärtsbrechung | `-normal * lens * refraction * bevelWidth * 0.34` |
| Leichte innere Vergrößerung | `-local * (0.012 * refraction)` |
| Chromatische Dispersion | `lens * refraction * 0.35` |
| RGB-Aufspaltungsgewicht | `bevel * 0.65` |
| Farbtonmischung | `vec3(0.84, 0.92, 0.89)` mit `tint * 0.31 + frost * 0.10` |
| Grundhelligkeit der Linse | `col *= 0.91` |
| Markeraufhellung | `vec3(0.94, 0.98, 0.96)` mit `u_control * 0.14` |
| Randabfall | `exp(-depth * 0.8)` |
| Randlichtgewicht | `0.10 + highlight * 0.52 + backlight * 0.12` |
| Zeigerlicht | Ellipse `300 × 230`, Stärke `0.045` |
| Popup-Innendichte | `0.66 * smoothstep(2.0, 18.0, depth)` gegen `vec3(0.027, 0.094, 0.078)` |

Die begrenzte Einwärtsbrechung verhindert zurückgefaltete Bilddetails, die wie gestapelte Formen wirken. Dispersion bleibt unter einem CSS-Pixel. Bei Kristall überspringt der Shader den Neun-Sample-Blur, weil `frost = 0` ist. Normale kleine Widgets erhalten keine Popup-Innendichte; `.floating-panel` und seine optischen Kinder schon.

### Übergänge und Federuhr

HTML-Geometrie und optische Linsen werden in **demselben Animationsframe** aktualisiert. Es darf keine unabhängige CSS-Transition auf der vom Controller gesetzten Markertransformation liegen. Ein neuer Klick lenkt die laufende Feder um; er wartet nicht auf das Ende einer alten Animation.

| Feder / Effekt | Verbindliche Werte |
| --- | --- |
| Grundfeder | Steifigkeit `460`, Dämpfung `30` |
| Langsamere Markerkante | `370 / 27` |
| Schnellere Markerkante | `570 / 31` |
| Vertikale Position und Höhe | `460 / 31` |
| Druck | `510 / 25`, gedrückt Ziel `.87`, losgelassen `1` |
| Griffwachstum | `440 / 29`, Ziel `1` beim Halten, sonst `0` |
| Breiter Regler während Drag | Linke Kante `1000`, rechte `1100`, beide Dämpfung `48` |
| Glasbutton | `540 / 32`, gedrückt `scale(.955)` |
| Zeitintegration | Substeps höchstens `1/180 s`; Frame-dt zwischen `1/240` und `1/30 s` |
| Exaktes Einrasten | Distanz `< .012` und Geschwindigkeit `< .06` |
| Äußere Nav-Kapsel | Maximal `scale(1.008, .988)` |

Die beiden horizontalen Kanten tauschen ihre schnellere/langsamere Rolle mit der Bewegungsrichtung. Druck und Geschwindigkeit reduzieren die Höhe begrenzt auf `.78 … 1.065` der Zielhöhe. Der Knopf wächst beim Halten um bis zu ungefähr 30 % seiner Ruhebreite, mit zusätzlicher geschwindigkeitsabhängiger Streckung bis 18 %. Seine Höhe wächst um 9 % der Engagement-Auslenkung. Trackgrenzen begrenzen diese Verformung.

Popover öffnen: `460ms`, von `scale(.94,.90)` über `scale(1.004,1.008)` bei Offset `.72` nach `scale(1)`. Schließen: `180ms`. Easing: `cubic-bezier(.2,.8,.25,1)`. Die Ursprungstranslation wird aus dem auslösenden Button abgeleitet und klein gehalten. Bei Umkehr zuerst die aktuelle Transformationsmatrix und Opazität lesen, dann die alte Animation abbrechen und von diesem Zustand starten. Glas und HTML müssen dieselbe Opazität erhalten. Bei zusätzlich laufenden CSS-Fades die optische Uhr für deren endliche Dauer mit `animateFor()` versorgen.

Bei `prefers-reduced-motion: reduce` erreichen Federn sofort ihr Ziel; Popovers öffnen und schließen ohne räumliche Verformung. Ein laufender Wechsel der Einstellung muss auch bereits aktive Controller vollständig beenden. Pointer-Up außerhalb eines Controls, Pointer-Cancel und Fenster-Blur lösen gehaltene Zustände. Die optische Federuhr endet nach dem Einrasten; nur ausdrücklich erforderliche sichtbare Wiedergabeanimationen dürfen weiterlaufen.

## Shapes

| Komponente | Referenzmaße |
| --- | --- |
| Iconbutton | `46 × 46px`, Icon `20 × 20px`, Padding `0`, kreisförmig |
| Kleine untere Aktion | `41 × 41px`, Icon `18 × 18px` |
| Obere Navigation | Höhe `52px`, Innenabstand `5px`, Tab-Padding `11px 23px` |
| Untere Navigation | Padding `7px 11px`, Gap `6px`, Buttons `48 × 48px`, Icons `23px` |
| Mobile Dock-Variante | Padding `6px 9px`, Gap `3px`, Buttons `46px` |
| Großer Toggle | Track `52 × 30px`, Ruheknopf `26px`, Inset `2px` |
| Kleiner Toggle | Track `40 × 24px`, Ruheknopf `20px`, Inset `2px` |
| Breiter Regler | Höhe `38px`, Ruheknopf `32px`, horizontales Inset `2px` |
| Kompakter Regler | Wrapper `24px`, Track `6px`, Ruheknopf `20px` |
| Widget / großes Panel | Radius `30px` / `32px`; mobile Widgets `27px` |

Kreise im Renderer mit `data-glass="circle"` kennzeichnen. Pillen und abgerundete Rechtecke nutzen `data-glass="pill"` beziehungsweise `data-glass="panel"`; der Renderer liest ihren tatsächlichen CSS-Radius. Bei Verformung folgt der Marker-Radius `min(controllerRadius, aktuelleHöhe / 2)`.

Icons besitzen explizit `viewBox="0 0 24 24"` und `display: block`. Nicht nur den SVG-Rahmen prüfen, sondern auch die tatsächlich gezeichneten Pfade. Die Referenz-Playform besitzt die Punkte `(6,4)`, `(18,12)` und `(6,20)`; ihre gezeichneten Grenzen sind um `(12,12)` zentriert. Ein Pausewechsel darf keinen seitlichen Sprung erzeugen.

## Components

### DOM-Vertrag des Referenzrenderers

| Selektor / Attribut | Bedeutung |
| --- | --- |
| `[data-glass]` | Sichtbare DOM-Linse; Reihenfolge im DOM ist die Zeichenreihenfolge |
| `.liquid-marker` | Absolut positionierte, rein optische Controllerform; kein eigener Schatten |
| `.filled-range` | Breite Reglerlinse; Wrapper enthält den Input und `.range-thumb` |
| `.range-thumb` | Tatsächliche animierte Knopfgeometrie; breite Variante wird nicht separat gerendert |
| `data-fill` / `--visual-fill` | Aktuelles animiertes Knopfzentrum als Fortschritt, nicht nur der Inputwert |
| `.floating-panel` | Tiefenabhängige Lesbarkeitsdichte für Panel und optische Kinder |
| `.workspace` / `.intro` | Vorhandene Vorfahren-Opazität wird mit dem optischen Alpha multipliziert |
| `.home-view` innerhalb `main` | Besonderer mobiler Overflow-Clip der Ursprungsanwendung |
| `data-renderer="webgl"` / `"css"` am HTML-Element | CSS-Hintergrund und Rand zwischen Optik und Fallback umschalten |

Bei Umbenennung diese Verträge im Renderer mit anpassen. Bei zusätzlichen Fade-Vorfahren deren Alpha ebenfalls berücksichtigen. Eine zweite Canvas-Linse sieht den Inhalt einer früher gerenderten Linse nicht: Beide samplen dieselbe Hintergrundtextur.

### Breite Regler: keine senkrechte Schnittkante

Der Knopf ist ein weicher optischer Hügel **innerhalb derselben Linse**. Der Renderer überspringt deshalb eine `.range-thumb` innerhalb von `.filled-range` als separaten Draw. Die Knopfgeometrie wird stattdessen als `u_knob` an den Wrapper-Shader übergeben.

- Knopfmaske: `1 - smoothstep(-4, 7, knobDistance)`.
- Füllung: gerundetes SDF-Rechteck bis zur rechten aktuellen Knopfkante, weich mit `1 - smoothstep(-6, 6, fillDistance)`.
- Innere Knopfbrechung: `samplePoint -= knobLocal * (0.055 * refraction * knob)`.
- Aufhellung: `filled * .23 + knob * .13`.
- Weicher Knopfreflex: `exp(-pow(knobDistance / 4.2, 2)) * .085`.

Keine harte vertikale Maske bei `value%` einsetzen. Keine zweite deckende Hintergrundkopie als Knopf darüberlegen. Inputschritte werden mit den schnellen Reglerfedern interpoliert; direktes Snappen bei jeder Eingabe löscht die Bewegungsgeschwindigkeit und fühlt sich starr an.

Breite native Inputs haben eine unsichtbare `36px`-Daumenbreite; damit stimmt ihr Zentrum mit dem `32px`-Knopf plus beidseitig `2px` Inset überein. Die Ruheposition ist `x = 2 + (Wrapperbreite - Knopfgröße - 4) * Fortschritt`. Füllung aus der **aktuellen animierten Knopfmitte** ableiten. Labels, native Tastaturbedienung, 0–100-Endanschläge und sichtbarer Fokus bleiben erhalten.

### Toggles und Navigation

Toggle-Ziele: aus `x = 2`, an `x = Trackbreite - 2 - Knopfgröße`; vertikal immer zentriert. Der native Checkbox-Input deckt den ganzen Track ab. Pointer-Capture hält den Drag stabil. Nach mehr als `3px` horizontaler Bewegung kontinuierlich vorschauen; beim Loslassen ab der Trackmitte auf „an“ wechseln. Ein Pointer-Abbruch stellt den Ausgangszustand wieder her. Den anschließenden synthetischen Klick nach einem Drag unterdrücken, sonst schaltet der Toggle doppelt.

Ein Navigationsmarker ist ein separates absolutes Element unter dem Text beziehungsweise Icon. Die Buttons erhalten keine zweite aktive Hintergrundfüllung. Beim Wechsel den bestehenden Marker umlenken; keine alten Marker als ausblendende Kopien liegenlassen. `aria-current` und gegebenenfalls `aria-pressed` mit dem echten Zustand synchronisieren.

### Optionale Wasserbewegung

Drei transparente SVG-Strömungen bilden Wasseroberflächen statt Equalizer-Balken. Perioden: `5.6s`, `8.4s` rückwärts und `6.8s`, jeweils linear über `110px` in einem `110 × 30`-Viewport. Wasserflächen mit `.10` Alpha, Oberflächenlinien mit `.75` Stroke-Alpha und `.8px` Strichbreite. Den unteren Rand ausblenden. Nur die Wasseroberfläche zeichnen; geschlossene Körperpfade nicht rundum umranden, sonst entsteht ein sichtbares Rechteck.

Die Strömung läuft nur bei aktiver Wiedergabe und sichtbarem Motiv. Pause und ausgeblendetes Dokument stoppen sie; reduzierte Bewegung entfernt sie. Die Referenzdemo hat dafür einen ausdrücklich beschrifteten Bewegungsschalter ohne Audio. In einer echten Musikanwendung an den realen Wiedergabezustand binden.

### Referenzdatei: index.html

```html
<!doctype html>
<html lang="de">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Kristall — Materialreferenz</title>
  <link rel="icon" href="data:,">
  <link rel="stylesheet" href="styles.css">
  <script src="app.js" type="module"></script>
</head>
<body>
  <div class="wallpaper" aria-hidden="true"></div>
  <canvas id="glass-canvas" aria-hidden="true"></canvas>
  <div class="vignette" aria-hidden="true"></div>
  <div class="example">
    <header>
      <nav class="top-nav glass" data-glass="pill" aria-label="Materialnavigation">
        <span class="nav-selection liquid-marker glass" data-glass="pill" aria-hidden="true"></span>
        <button class="nav-tab" data-key="one" aria-current="page">Übersicht</button>
        <button class="nav-tab" data-key="two">Sammlung</button>
        <button class="nav-tab" data-key="three">Einstellungen</button>
      </nav>
      <button class="icon-button glass" data-glass="circle" aria-label="Beispielaktion">
        <svg class="icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 7h5m4 0h7M4 17h9m4 0h3"/><circle cx="11" cy="7" r="2.5"/><circle cx="15" cy="17" r="2.5"/></svg>
      </button>
    </header>
    <main class="example-grid">
      <section class="card glass" data-glass="panel" aria-labelledby="reference-title">
        <h1 id="reference-title">Kristall</h1><p>Klares Glas. Weiche Bewegung.</p>
        <label class="inline-slider"><svg class="icon" viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2m0 16v2M2 12h2m16 0h2M5 5l1.4 1.4m11.2 11.2L19 19M5 19l1.4-1.4M17.6 6.4L19 5"/></svg><input id="brightness" type="range" min="0" max="100" value="100" aria-label="Helligkeit"><output id="brightness-value">100%</output></label>
        <label class="inline-slider"><svg class="icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M11 4 5 9H2v6h3l6 5Zm4 4a6 6 0 0 1 0 8m3-11a10 10 0 0 1 0 14"/></svg><input id="volume" type="range" min="0" max="100" value="35" aria-label="Lautstärke-Beispiel"><output id="volume-value">35%</output></label>
        <div class="water-row">
          <svg class="water-wave" viewBox="0 0 110 30" aria-hidden="true">
            <g class="water-current back"><path d="M0 12Q27.5 3 55 12T110 12T165 12T220 12V36H0Z"/><path class="water-line" d="M0 12Q27.5 3 55 12T110 12T165 12T220 12"/></g>
            <g class="water-current"><path d="M0 16Q27.5 7 55 16T110 16T165 16T220 16V36H0Z"/><path class="water-line" d="M0 16Q27.5 7 55 16T110 16T165 16T220 16"/></g>
            <g class="water-current front"><path d="M0 21Q27.5 14 55 21T110 21T165 21T220 21V36H0Z"/><path class="water-line" d="M0 21Q27.5 14 55 21T110 21T165 21T220 21"/></g>
          </svg>
          <button id="water-toggle" class="icon-button glass" data-glass="circle" aria-label="Wasserbewegung starten" aria-pressed="false"><svg class="icon" viewBox="0 0 24 24" aria-hidden="true"><path id="play-path" d="m6 4 12 8-12 8Z" fill="currentColor" stroke="none"/></svg></button>
        </div>
      </section>
      <section class="floating-panel glass" data-glass="panel" aria-labelledby="options-title">
        <h2 id="options-title">Liquid Controls</h2>
        <label class="switch-row" for="option"><span>Option</span><span class="liquid-switch"><input id="option" class="switch-input" type="checkbox" checked><span class="switch-track" aria-hidden="true"></span><span class="switch-thumb liquid-marker glass" data-glass="pill" aria-hidden="true"></span></span></label>
        <label class="switch-row" for="small-option"><span>Kleine Variante</span><span class="mini-switch"><input id="small-option" class="switch-input" type="checkbox"><span class="switch-track" aria-hidden="true"></span><span class="switch-thumb liquid-marker glass" data-glass="pill" aria-hidden="true"></span></span></label>
        <label class="material-slider" for="fine"><span>Feinregler <output id="fine-value">42%</output></span><input id="fine" type="range" min="0" max="100" value="42"></label>
        <p class="small">Die Option-Schalter und die Lautstärke dienen hier der Materialprüfung. Das Beispiel erzeugt kein Audio.</p>
      </section>
    </main>
    <nav class="dock glass" data-glass="pill" aria-label="Schnellnavigation">
      <span class="dock-selection liquid-marker glass" data-glass="pill" aria-hidden="true"></span>
      <button class="dock-button" data-key="one" aria-label="Übersicht" aria-current="page"><svg class="icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m3 10 9-7 9 7v10a1 1 0 0 1-1 1h-6v-7h-4v7H4a1 1 0 0 1-1-1Z"/></svg></button>
      <button class="dock-button" data-key="two" aria-label="Sammlung"><svg class="icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3.5" y="3.5" width="6" height="6" rx="1.7"/><rect x="14.5" y="3.5" width="6" height="6" rx="1.7"/><rect x="3.5" y="14.5" width="6" height="6" rx="1.7"/><rect x="14.5" y="14.5" width="6" height="6" rx="1.7"/></svg></button>
      <button class="dock-button" data-key="three" aria-label="Einstellungen"><svg class="icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 7h5m4 0h7M4 17h9m4 0h3"/><circle cx="11" cy="7" r="2.5"/><circle cx="15" cy="17" r="2.5"/></svg></button>
    </nav>
    <p class="small status" id="renderer-status" role="status">Material wird geladen.</p>
  </div>
</body>
</html>
```

### Referenzdatei: styles.css

```css
:root {
  color-scheme: dark;
  --surface: rgba(220, 243, 235, .104);
  --wallpaper: url('assets/background.jpg');
  font-family: -apple-system, BlinkMacSystemFont, 'Helvetica Neue', 'Segoe UI', sans-serif;
  font-synthesis: none;
  -webkit-font-smoothing: antialiased;
}
* { box-sizing: border-box; }
body { margin: 0; min-width: 320px; color: #fff; background: #244b47; }
button, input { font: inherit; color: inherit; -webkit-tap-highlight-color: transparent; }
button { border: 0; cursor: pointer; touch-action: manipulation; transition: background .2s, color .2s, box-shadow .25s; }
button:focus-visible { outline: 2px solid #d5fff0; outline-offset: 5px; }
button:disabled { opacity: .5; cursor: progress; }
[hidden] { display: none !important; }
::selection { color: #173e35; background: #d2ffec; }
::-webkit-scrollbar { width: 6px; }
::-webkit-scrollbar-thumb { background: #ffffff50; border-radius: 8px; }
::-webkit-scrollbar-track { background: transparent; }
.wallpaper, #glass-canvas, .vignette { position: fixed; inset: 0; width: 100%; height: 100%; pointer-events: none; }
.wallpaper { background: var(--wallpaper) center / cover; }
#glass-canvas { z-index: 1; }
.vignette { z-index: 2; background: linear-gradient(90deg, rgba(4,19,17,.22), transparent 65%), linear-gradient(180deg, rgba(5,24,23,.15), transparent 25%, transparent 68%, rgba(6,24,22,.38)); }
.glass { position: relative; isolation: isolate; background: var(--surface); backdrop-filter: blur(0px) saturate(1.35); -webkit-backdrop-filter: blur(0px) saturate(1.35); box-shadow: 0 10px 28px -12px rgba(0,0,0,.28); }
.glass::before { content: ''; position: absolute; inset: 0; z-index: -1; border-radius: inherit; pointer-events: none; border: 1px solid rgba(255,255,255,.24); }
[data-renderer="webgl"] .glass { background: transparent; backdrop-filter: none; -webkit-backdrop-filter: none; }
[data-renderer="webgl"] .glass::before { border: 0; }
.glass.liquid-marker { position: absolute; top: 0; left: 0; margin: 0; pointer-events: none; box-shadow: none; z-index: 0; transform-origin: center; }
.switch-thumb { z-index: 1 !important; }
.icon { display: block; width: 24px; height: 24px; fill: none; stroke: currentColor; stroke-width: 1.65; stroke-linecap: round; stroke-linejoin: round; flex-shrink: 0; }
.icon-button { display: grid; place-items: center; width: 46px; height: 46px; padding: 0; border-radius: 50%; }
.icon-button .icon { width: 20px; height: 20px; }
.icon-button:hover { box-shadow: 0 10px 28px -12px rgba(0,0,0,.36); }
.top-nav { display: flex; align-items: center; position: relative; height: 52px; padding: 5px; border-radius: 50px; transform-origin: center; }
.nav-tab { position: relative; z-index: 1; padding: 11px 23px; background: transparent; white-space: nowrap; font-size: 13px; letter-spacing: .005em; border-radius: 30px; color: #ffffffe0; }
.nav-tab[aria-current] { color: #fff; }
.nav-tab:hover { background: #ffffff14; }
.dock { display: flex; align-items: center; gap: 6px; padding: 7px 11px; border-radius: 27px; width: max-content; margin: 28px auto; transform-origin: center; }
.dock-button { display: grid; place-items: center; position: relative; z-index: 1; width: 48px; height: 48px; padding: 0; border-radius: 19px; background: transparent; }
.dock-button .icon { width: 23px; height: 23px; }
.dock-button:hover { color: #e1fff3; }
.card, .floating-panel { padding: 22px; border-radius: 30px; min-width: 0; }
.floating-panel { padding: 26px; border-radius: 32px; }
.floating-panel::before { background: rgba(7,24,20,.66); }
[data-renderer="webgl"] .floating-panel::before { background: none; }
h1, h2 { font-size: 22px; font-weight: 550; letter-spacing: -.025em; margin: 0 0 16px; }
p { color: #ffffffc7; font-size: 16px; line-height: 1.5; }
.small { font-size: 11px; }
.inline-slider { display: grid; grid-template-columns: 24px minmax(0,1fr) 38px; align-items: center; gap: 12px; min-height: 44px; margin-top: 16px; cursor: pointer; }
.inline-slider .icon { width: 21px; height: 21px; justify-self: center; }
output { font-size: 11px; text-align: right; font-variant-numeric: tabular-nums; }
input[type="range"] { --range-fill: 50%; appearance: none; -webkit-appearance: none; width: 100%; min-width: 0; height: 6px; border-radius: 10px; background: linear-gradient(to right, #ffffffdb calc(var(--range-fill) - 9px), #ffffff2b calc(var(--range-fill) + 9px)); cursor: ew-resize; }
input[type="range"]::-webkit-slider-thumb { appearance: none; -webkit-appearance: none; width: 24px; height: 24px; border: 0; background: transparent; }
input[type="range"]::-moz-range-thumb { width: 24px; height: 24px; border: 0; background: transparent; }
.range-control { position: relative; display: block; height: 24px; min-width: 0; width: 100%; touch-action: pan-y; }
.range-control input { position: absolute; z-index: 2; top: 9px; left: 0; margin: 0; height: 6px; width: 100%; touch-action: pan-y; }
.range-control input:focus-visible { outline: none; }
.range-control:has(input:focus-visible) .range-thumb { outline: 2px solid #d5fff0; outline-offset: 4px; }
.filled-range { height: 38px; border-radius: 22px; box-shadow: none; }
.filled-range input { inset: 0; height: 100%; opacity: 0; background: none; border-radius: inherit; }
.filled-range input::-webkit-slider-thumb { width: 36px; height: 38px; }
.filled-range input::-moz-range-thumb { width: 36px; height: 38px; }
.range-fill { display: none; }
.filled-range .range-fill { display: block; position: absolute; inset: 0 auto 0 0; width: min(100%, calc(var(--visual-fill,50%) + 16px)); background: #ffffff38; border-radius: inherit; filter: blur(1px); pointer-events: none; }
[data-renderer="webgl"] .filled-range .range-fill { display: none; }
.filled-range:has(input:focus-visible) { outline: 2px solid #d5fff0; outline-offset: 4px; }
.filled-range:has(input:focus-visible) .range-thumb { outline: none; }
.material-slider { display: block; margin-top: 28px; }
.material-slider > span { display: flex; justify-content: space-between; font-size: 12px; margin-bottom: 13px; }
.switch-row { display: flex; align-items: center; justify-content: space-between; gap: 12px; font-size: 12px; margin-top: 20px; cursor: pointer; }
.liquid-switch, .mini-switch { position: relative; display: block; flex-shrink: 0; width: 52px; height: 30px; touch-action: pan-y; }
.mini-switch { width: 40px; height: 24px; }
.switch-input { position: absolute; z-index: 2; inset: 0; margin: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; touch-action: pan-y; }
.switch-track { position: absolute; inset: 0; border-radius: 20px; background: #ffffff30; }
.switch-track::before { content: ''; position: absolute; inset: 0; border-radius: inherit; background: #96dab899; opacity: var(--switch-progress,0); }
.switch-input:focus-visible ~ .switch-thumb { outline: 2px solid #d5fff0; outline-offset: 4px; }
.water-row { display: flex; align-items: center; justify-content: space-between; margin-top: 26px; }
.water-wave { display: block; width: 110px; height: 30px; overflow: hidden; color: #e3fff5; mask-image: linear-gradient(#000 65%,transparent); }
.water-current { animation: water-flow 5.6s linear infinite; animation-play-state: paused; }
.water-current path { fill: currentColor; fill-opacity: .10; stroke: none; }
.water-current .water-line { fill: none; stroke: currentColor; stroke-opacity: .75; stroke-width: .8; }
.water-current.back { animation-duration: 8.4s; animation-direction: reverse; opacity: .55; }
.water-current.front { animation-duration: 6.8s; opacity: .7; }
.is-playing .water-wave[data-visible="true"] .water-current { animation-play-state: running; }
@keyframes water-flow { to { transform: translateX(-110px); } }
/* Nur das Layout des Prüfgerüsts, keine Materialvorgabe für andere Produkte. */
.example { position: relative; z-index: 3; max-width: 1000px; margin: auto; padding: 44px 28px; }
header { display: flex; align-items: center; justify-content: space-between; gap: 20px; margin-bottom: 80px; }
.example-grid { display: grid; grid-template-columns: minmax(0,1.2fr) minmax(0,1fr); gap: 24px; }
.status { text-align: center; }
@media(max-width:760px) {
  .example { padding: 24px 18px; }
  header { flex-wrap: wrap; margin-bottom: 40px; }
  .top-nav { max-width: 100%; }
  .nav-tab { padding-inline: 12px; }
  .example-grid { grid-template-columns: minmax(0,1fr); }
  .dock { padding: 6px 9px; gap: 3px; }
  .dock-button { width: 46px; height: 46px; }
}
@media(prefers-reduced-motion:reduce) {
  .water-current { animation: none; }
  *, *::before, *::after { transition-duration: .001ms !important; scroll-behavior: auto !important; }
}
@media(prefers-reduced-transparency:reduce) { .glass::before { background: rgba(8,32,27,.8); } }
```

### Referenzdatei: app.js

Dieses kleine Integrationsbeispiel verbindet die unveränderten Materialmodule mit dem Prüfgerüst. Es ist kein Ersatz für die Fachlogik des Zielprojekts.

```js
import { GlassRenderer } from './glass.js';
import { LiquidMotion, LiquidMarker, LiquidToggle, LiquidRange, LiquidSurface } from './motion.js';

const $ = selector => document.querySelector(selector);
const $$ = selector => [...document.querySelectorAll(selector)];
const CRYSTAL = Object.freeze({ refraction: 1, frost: 0, tint: .08, brightness: 1, contrast: 0, focus: 0, motion: true });
let renderer;
const clock = new LiquidMotion(() => renderer?.draw());
const nav = new LiquidMarker($('.nav-selection'), clock, { container: $('.top-nav'), radius: 30 });
const dock = new LiquidMarker($('.dock-selection'), clock, { container: $('.dock'), radius: 21 });
const ranges = new Map();
const toggles = new Map();
let selection = 'one';

function select(key, immediate = false) {
  selection = key;
  $$('[data-key]').forEach(button => {
    if (button.dataset.key === key) button.setAttribute('aria-current', 'page');
    else button.removeAttribute('aria-current');
  });
  nav.follow($(`.nav-tab[data-key="${key}"]`), immediate);
  dock.follow($(`.dock-button[data-key="${key}"]`), immediate);
}

function updateRange(input, immediate = false) {
  const progress = (Number(input.value) - Number(input.min)) / (Number(input.max) - Number(input.min));
  input.style.setProperty('--range-fill', `${progress * 100}%`);
  $(`#${input.id}-value`).textContent = `${input.value}%`;
  ranges.get(input.id).update(immediate);
  if (input.id === 'brightness') {
    if (renderer?.ready) renderer.update({ brightness: Number(input.value) / 100 });
    else $('.wallpaper').style.filter = `brightness(${Number(input.value) / 100})`;
  }
}

$$('input[type="range"]').forEach(input => {
  const large = ['brightness', 'volume'].includes(input.id);
  const wrapper = document.createElement('span');
  wrapper.className = large ? 'range-control filled-range glass' : 'range-control';
  if (large) wrapper.dataset.glass = 'pill';
  const fill = document.createElement('span'); fill.className = 'range-fill'; fill.setAttribute('aria-hidden', 'true');
  const thumb = document.createElement('span'); thumb.className = 'range-thumb liquid-marker glass'; thumb.dataset.glass = 'pill'; thumb.setAttribute('aria-hidden', 'true');
  input.before(wrapper); wrapper.append(fill, input, thumb);
  const marker = new LiquidRange(thumb, clock, input, large);
  ranges.set(input.id, marker);
  input.addEventListener('input', () => updateRange(input));
  input.addEventListener('pointerdown', () => marker.press());
  input.addEventListener('pointerup', () => marker.release());
  input.addEventListener('pointercancel', () => marker.release());
});

function syncToggle(input, immediate = false) {
  const track = input.parentElement;
  const size = track.classList.contains('mini-switch') ? 20 : 26;
  toggles.get(input.id).to({ x: input.checked ? track.clientWidth - 2 - size : 2, y: 2, width: size, height: size }, immediate);
}
$$('.switch-input').forEach(input => {
  const marker = new LiquidToggle(input.parentElement.querySelector('.switch-thumb'), clock);
  toggles.set(input.id, marker);
  input.addEventListener('change', () => syncToggle(input));
  marker.bind(input);
});

for (const [selector, marker] of [['.nav-tab', nav], ['.dock-button', dock]]) {
  $$(selector).forEach(button => {
    button.addEventListener('click', () => select(button.dataset.key));
    button.addEventListener('pointerdown', () => marker.press());
    button.addEventListener('pointerup', () => marker.release());
    button.addEventListener('pointercancel', () => marker.release());
  });
}
const surfaces = $$('.icon-button').map(element => {
  const surface = new LiquidSurface(element, clock);
  element.addEventListener('pointerdown', () => surface.press());
  element.addEventListener('pointerup', () => surface.release());
  element.addEventListener('pointercancel', () => surface.release());
  return surface;
});
const release = () => [nav, dock, ...ranges.values(), ...toggles.values(), ...surfaces].forEach(control => control.release());
window.addEventListener('pointerup', release);
window.addEventListener('pointercancel', release);
window.addEventListener('blur', release);

function setWater(playing) {
  document.body.classList.toggle('is-playing', playing);
  $('#water-toggle').setAttribute('aria-pressed', playing);
  $('#water-toggle').setAttribute('aria-label', playing ? 'Wasserbewegung pausieren' : 'Wasserbewegung starten');
  $('#play-path').setAttribute('d', playing ? 'M6 5h4v14H6zm8 0h4v14h-4z' : 'm6 4 12 8-12 8Z');
}
$('#water-toggle').addEventListener('click', () => setWater(!document.body.classList.contains('is-playing')));
new IntersectionObserver(entries => { $('.water-wave').dataset.visible = entries[0].isIntersecting; }).observe($('.water-wave'));
document.addEventListener('visibilitychange', () => {
  if (document.hidden) setWater(false);
  else renderer?.requestDraw();
});

function layout() {
  select(selection, true);
  $$('input[type="range"]').forEach(input => updateRange(input, true));
  $$('.switch-input').forEach(input => syncToggle(input, true));
}
function fallback() {
  document.documentElement.dataset.renderer = 'css';
  $('#glass-canvas').style.visibility = 'hidden';
  $('#renderer-status').textContent = 'Kristall-Fallback · ohne optische Brechung';
  $('.wallpaper').style.filter = `brightness(${Number($('#brightness').value) / 100})`;
}
layout();
try {
  renderer = new GlassRenderer($('#glass-canvas'));
  renderer.update(CRYSTAL);
  await renderer.loadScene('background');
  $('#renderer-status').textContent = 'Kristall · optische Brechung aktiv';
} catch { fallback(); }
document.addEventListener('glass-fallback', fallback);
document.addEventListener('glass-ready', () => {
  renderer.update({ ...CRYSTAL, brightness: Number($('#brightness').value) / 100 });
  $('.wallpaper').style.filter = '';
  $('#renderer-status').textContent = 'Kristall · optische Brechung aktiv';
});
window.addEventListener('resize', layout);
layout();
// Nur Diagnose des Prüfgerüsts; keine Abhängigkeit der Materialmodule.
window.crystalReference = { renderer, clock, material: CRYSTAL };
```

### Referenzdatei: glass.js

Vollständige, unveränderte optische Implementierung aus der freigegebenen Version 1.3.0. Der Szenenloader erwartet `assets/<name>.jpg` relativ zu dieser Moduldatei. Das Beispiel ruft ihn mit `background` auf. Bei anderer Assetverwaltung ausschließlich diesen Ladevertrag anpassen; der Shader bleibt gleich.

```js
// The photograph is rendered once, then sampled through rounded optical lenses.
// HTML above this canvas provides the accessible, interactive control layer.
const vertexSource = `
attribute vec2 a_position;
void main() { gl_Position = vec4(a_position, 0.0, 1.0); }
`;

const fragmentSource = `
precision highp float;
uniform sampler2D u_image;
uniform vec2 u_resolution;
uniform vec2 u_imageSize;
uniform vec4 u_rect;
uniform vec2 u_pixelRatio;
uniform float u_radius;
uniform float u_mode;
uniform float u_refraction;
uniform float u_frost;
uniform float u_tint;
uniform float u_brightness;
uniform float u_contrast;
uniform float u_focus;
uniform float u_readability;
uniform float u_control;
uniform float u_opacity;
uniform vec4 u_range;
uniform vec4 u_knob;
uniform float u_fill;
uniform vec2 u_pointer;

vec2 imageUV(vec2 p) {
  vec2 uv = p / u_resolution;
  float screenRatio = u_resolution.x / u_resolution.y;
  float imageRatio = u_imageSize.x / u_imageSize.y;
  if (screenRatio > imageRatio) uv.y = (uv.y - 0.5) * imageRatio / screenRatio + 0.5;
  else uv.x = (uv.x - 0.5) * screenRatio / imageRatio + 0.5;
  return clamp(uv, vec2(0.001), vec2(0.999));
}
vec3 scene(vec2 p) {
  vec3 color = texture2D(u_image, imageUV(p)).rgb;
  // A photographic exposure adjustment keeps highlights and reflections intact.
  color *= u_brightness;
  color = mix(color, color * vec3(0.70, 0.78, 0.77), u_focus * 0.6);
  return color;
}
float roundedDistance(vec2 p, vec2 halfSize, float r) {
  vec2 q = abs(p) - halfSize + r;
  return length(max(q, 0.0)) + min(max(q.x, q.y), 0.0) - r;
}
void main() {
  vec2 p = vec2(gl_FragCoord.x / u_pixelRatio.x, u_resolution.y - gl_FragCoord.y / u_pixelRatio.y);
  if (u_mode < 0.5) { gl_FragColor = vec4(scene(p), 1.0); return; }
  vec2 halfSize = u_rect.zw * 0.5;
  vec2 local = p - u_rect.xy - halfSize;
  float radius = min(u_radius, min(halfSize.x, halfSize.y));
  float sdf = roundedDistance(local, halfSize, radius);
  float coverage = 1.0 - smoothstep(-0.8, 0.5, sdf);
  if (coverage <= 0.0) discard;

  float depth = max(-sdf, 0.0);
  float bevelWidth = min(radius * 0.85, 24.0);
  float bevel = 1.0 - smoothstep(0.0, bevelWidth, depth);
  // Analytic SDF normal avoids four additional distance evaluations per pixel.
  vec2 q = abs(local) - halfSize + radius;
  vec2 outside = max(q, 0.0);
  vec2 normal = length(outside) > 0.001
    ? normalize(outside) * sign(local)
    : (q.x > q.y ? vec2(sign(local.x), 0.0) : vec2(0.0, sign(local.y)));
  float lens = sin(bevel * 1.570796);
  // Keep the mapping monotonic through the bevel. A stronger offset folds the
  // photograph back onto itself and looks like several stacked silhouettes.
  vec2 displacement = -normal * lens * u_refraction * bevelWidth * 0.34;
  // A gentle central magnification and stronger peripheral bending create depth.
  vec2 samplePoint = p + displacement - local * (0.012 * u_refraction);
  float filled = 0.0;
  float knob = 0.0;
  float knobDistance = 1000.0;
  if (u_fill >= 0.0) {
    vec2 knobHalf = u_knob.zw * 0.5;
    vec2 knobLocal = p - u_knob.xy - knobHalf;
    knobDistance = roundedDistance(knobLocal, knobHalf, min(knobHalf.x, knobHalf.y));
    knob = 1.0 - smoothstep(-4.0, 7.0, knobDistance);
    float end = min(u_range.x + u_range.z, u_knob.x + u_knob.z);
    vec2 fillHalf = vec2(max(1.0, end - u_range.x), u_range.w) * 0.5;
    float fillDistance = roundedDistance(p - u_range.xy - fillHalf, fillHalf, min(fillHalf.x, fillHalf.y));
    filled = 1.0 - smoothstep(-6.0, 6.0, fillDistance);
    // The knob is a smooth bulge in this same lens, not an opaque second
    // photograph with an abrupt boundary over the filled capsule.
    samplePoint -= knobLocal * (0.055 * u_refraction * knob);
  }
  float blur = u_frost * 15.0;
  vec3 col = scene(samplePoint);
  if (u_frost > 0.001) {
    col *= 0.24;
    col += scene(samplePoint + vec2(blur, 0.0)) * 0.12;
    col += scene(samplePoint - vec2(blur, 0.0)) * 0.12;
    col += scene(samplePoint + vec2(0.0, blur)) * 0.12;
    col += scene(samplePoint - vec2(0.0, blur)) * 0.12;
    col += scene(samplePoint + vec2(blur, blur) * 0.707) * 0.07;
    col += scene(samplePoint - vec2(blur, blur) * 0.707) * 0.07;
    col += scene(samplePoint + vec2(blur, -blur) * 0.707) * 0.07;
    col += scene(samplePoint + vec2(-blur, blur) * 0.707) * 0.07;
  }
  // Sub-pixel chromatic separation is confined to the curved glass edge.
  float dispersion = lens * u_refraction * 0.35;
  col.r = mix(col.r, scene(samplePoint + normal * dispersion).r, bevel * 0.65);
  col.b = mix(col.b, scene(samplePoint - normal * dispersion).b, bevel * 0.65);
  col = mix(col, vec3(0.84, 0.92, 0.89), u_tint * 0.31 + u_frost * 0.10);
  col *= 0.91;
  col = mix(col, vec3(0.94, 0.98, 0.96), u_control * 0.14);
  if (u_fill >= 0.0) {
    col = mix(col, vec3(0.96, 0.99, 0.98), filled * 0.23 + knob * 0.13);
    col += exp(-pow(knobDistance / 4.2, 2.0)) * 0.085;
  }
  col = mix(col, vec3(0.055, 0.15, 0.13), u_contrast * 0.70);
  // Large popovers gain optical density in their interior. The transparent
  // refractive bevel stays clear, while labels retain contrast over bright sky.
  col = mix(col, vec3(0.027, 0.094, 0.078), u_readability * 0.66 * smoothstep(2.0, 18.0, depth));

  vec2 light = normalize(vec2(-0.7, -1.0) + (u_pointer - p) / max(u_resolution, vec2(1.0)) * 1.7);
  float highlight = pow(max(dot(normal, light), 0.0), 4.0);
  float backlight = pow(max(dot(normal, -light), 0.0), 5.0);
  float rim = exp(-depth * 0.8);
  col += vec3(0.95, 1.0, 0.98) * rim * (0.10 + highlight * 0.52 + backlight * 0.12);
  // A wide specular response moves with the pointer without obscuring content.
  vec2 spot = (p - u_pointer) / vec2(300.0, 230.0);
  col += exp(-dot(spot, spot) * 2.0) * 0.045;
  gl_FragColor = vec4(col, coverage * u_opacity);
}
`;

export class GlassRenderer {
  constructor(canvas) {
    this.canvas = canvas;
    this.gl = canvas.getContext('webgl', { alpha: false, antialias: false, depth: false, powerPreference: 'low-power', preserveDrawingBuffer: true });
    if (!this.gl) throw new Error('WebGL unavailable');
    this.settings = { refraction: 1, frost: 0, tint: .08, brightness: 1, contrast: 0, focus: 0, motion: true };
    this.pointer = [window.innerWidth * .2, 0];
    this.imageSize = [1, 1];
    this.pending = false;
    this.frameId = 0;
    this.sceneGeneration = 0;
    this.ready = false;
    this.init();
    this.resize = this.resize.bind(this);
    window.addEventListener('resize', this.resize);
    window.addEventListener('scroll', () => this.requestDraw(), { passive: true, capture: true });
    window.addEventListener('pointermove', event => {
      if (!this.settings.motion || matchMedia('(prefers-reduced-motion: reduce)').matches) return;
      this.pointer = [event.clientX, event.clientY];
      this.requestDraw();
    }, { passive: true });
    this.observer = new ResizeObserver(() => { this.resize(); this.requestDraw(); });
    this.observer.observe(canvas);
    document.querySelectorAll('[data-glass]:not(.liquid-marker)').forEach(el => this.observer.observe(el));
    canvas.addEventListener('webglcontextlost', event => {
      event.preventDefault();
      this.ready = false;
      cancelAnimationFrame(this.frameId);
      this.pending = false;
      document.documentElement.dataset.renderer = 'css';
      canvas.style.visibility = 'hidden';
      document.dispatchEvent(new CustomEvent('glass-fallback'));
    });
    canvas.addEventListener('webglcontextrestored', async () => {
      try { this.init(); await this.loadScene(this.currentScene || 'alpine'); }
      catch { document.documentElement.dataset.renderer = 'css'; }
    });
    this.resize();
  }

  init() {
    const gl = this.gl;
    const compile = (type, source) => {
      const shader = gl.createShader(type);
      gl.shaderSource(shader, source);
      gl.compileShader(shader);
      if (!gl.getShaderParameter(shader, gl.COMPILE_STATUS)) throw new Error(gl.getShaderInfoLog(shader));
      return shader;
    };
    this.program = gl.createProgram();
    const vertex = compile(gl.VERTEX_SHADER, vertexSource);
    const fragment = compile(gl.FRAGMENT_SHADER, fragmentSource);
    gl.attachShader(this.program, vertex);
    gl.attachShader(this.program, fragment);
    gl.linkProgram(this.program);
    gl.deleteShader(vertex);
    gl.deleteShader(fragment);
    if (!gl.getProgramParameter(this.program, gl.LINK_STATUS)) throw new Error(gl.getProgramInfoLog(this.program));
    gl.useProgram(this.program);
    const buffer = gl.createBuffer();
    gl.bindBuffer(gl.ARRAY_BUFFER, buffer);
    gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([-1,-1, 3,-1, -1,3]), gl.STATIC_DRAW);
    const position = gl.getAttribLocation(this.program, 'a_position');
    gl.enableVertexAttribArray(position);
    gl.vertexAttribPointer(position, 2, gl.FLOAT, false, 0, 0);
    this.uniforms = {};
    ['image','resolution','imageSize','rect','pixelRatio','radius','mode','refraction','frost','tint','brightness','contrast','focus','readability','control','opacity','range','knob','fill','pointer'].forEach(name => {
      this.uniforms[name] = gl.getUniformLocation(this.program, `u_${name}`);
    });
    gl.activeTexture(gl.TEXTURE0);
    this.texture = gl.createTexture();
    gl.bindTexture(gl.TEXTURE_2D, this.texture);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);
    gl.uniform1i(this.uniforms.image, 0);
  }

  async loadScene(name) {
    const generation = ++this.sceneGeneration;
    const image = new Image();
    image.src = new URL(`assets/${name}.jpg`, import.meta.url).href;
    await image.decode();
    if (generation !== this.sceneGeneration || this.gl.isContextLost()) return false;
    const gl = this.gl;
    gl.bindTexture(gl.TEXTURE_2D, this.texture);
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGB, gl.RGB, gl.UNSIGNED_BYTE, image);
    this.imageSize = [image.naturalWidth, image.naturalHeight];
    this.currentScene = name;
    this.ready = true;
    document.documentElement.dataset.renderer = 'webgl';
    this.canvas.style.visibility = 'visible';
    document.dispatchEvent(new CustomEvent('glass-ready'));
    this.requestDraw();
    return true;
  }

  resize() {
    this.dpr = Math.min(window.devicePixelRatio || 1, 1.5);
    // innerWidth includes a classic scrollbar; the fixed CSS canvas does not.
    // Measure the actual drawing surface so lenses and HTML share coordinates.
    const bounds = this.canvas.getBoundingClientRect();
    this.width = bounds.width;
    this.height = bounds.height;
    const width = Math.round(this.width * this.dpr);
    const height = Math.round(this.height * this.dpr);
    if (this.canvas.width !== width) this.canvas.width = width;
    if (this.canvas.height !== height) this.canvas.height = height;
    this.pixelRatio = [width / this.width, height / this.height];
    this.gl.viewport(0, 0, this.canvas.width, this.canvas.height);
    this.requestDraw();
  }

  update(settings) {
    Object.assign(this.settings, settings);
    this.requestDraw();
  }

  requestDraw() {
    if (this.pending || !this.ready) return;
    this.pending = true;
    this.frameId = requestAnimationFrame(() => {
      this.pending = false;
      this.draw();
    });
  }

  draw() {
    if (!this.ready || document.hidden || this.gl.isContextLost()) return;
    const gl = this.gl;
    const u = this.uniforms;
    gl.useProgram(this.program);
    gl.uniform2f(u.resolution, this.width, this.height);
    gl.uniform2f(u.imageSize, ...this.imageSize);
    gl.uniform2f(u.pointer, ...this.pointer);
    gl.uniform2f(u.pixelRatio, ...this.pixelRatio);
    for (const key of ['refraction','frost','tint','brightness','contrast','focus']) gl.uniform1f(u[key], this.settings[key]);
    gl.disable(gl.SCISSOR_TEST);
    gl.disable(gl.BLEND);
    gl.uniform1f(u.mode, 0);
    gl.drawArrays(gl.TRIANGLES, 0, 3);
    gl.uniform1f(u.mode, 1);
    gl.enable(gl.SCISSOR_TEST);
    gl.enable(gl.BLEND);
    gl.blendFunc(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA);
    // DOM order is also the painter's order; popovers follow widgets in the DOM.
    document.querySelectorAll('[data-glass]').forEach(element => {
      if (element.classList.contains('range-thumb') && element.closest('.filled-range')) return;
      if (!element.checkVisibility({ checkVisibilityCSS: true, checkOpacity: true })) return;
      const rect = element.getBoundingClientRect();
      if (rect.bottom < 0 || rect.top > this.height || rect.width === 0) return;
      const style = getComputedStyle(element);
      const scale = Math.min(rect.width / Math.max(1, parseFloat(style.width) || element.offsetWidth), rect.height / Math.max(1, parseFloat(style.height) || element.offsetHeight));
      const radius = (parseFloat(style.borderTopLeftRadius) || 24) * scale;
      gl.uniform4f(u.rect, rect.x, rect.y, rect.width, rect.height);
      const popover = element.closest('.floating-panel');
      gl.uniform1f(u.readability, popover ? 1 : 0);
      gl.uniform1f(u.control, element.classList.contains('liquid-marker') || element.classList.contains('range-thumb') ? 1 : 0);
      let opacity = Number(style.opacity);
      if (popover && popover !== element) opacity *= Number(getComputedStyle(popover).opacity);
      const workspace = element.closest('.workspace');
      if (workspace && workspace !== element) opacity *= Number(getComputedStyle(workspace).opacity);
      const intro = element.closest('.intro');
      if (intro && intro !== element) opacity *= Number(getComputedStyle(intro).opacity);
      gl.uniform1f(u.opacity, opacity);
      const range = element.closest('.filled-range');
      if (range) {
        const rangeRect = range.getBoundingClientRect();
        const knobRect = range.querySelector('.range-thumb').getBoundingClientRect();
        gl.uniform4f(u.range, rangeRect.x, rangeRect.y, rangeRect.width, rangeRect.height);
        gl.uniform4f(u.knob, knobRect.x, knobRect.y, knobRect.width, knobRect.height);
        gl.uniform1f(u.fill, Number(range.dataset.fill || 0));
      } else gl.uniform1f(u.fill, -1);
      gl.uniform1f(u.radius, element.dataset.glass === 'circle' ? rect.width / 2 : radius);
      let clip = { left: 0, right: this.width, top: 0, bottom: this.height };
      // The mobile content viewport reserves a separate area for the dock.
      // Scissor the canvas lenses exactly where CSS clips their HTML controls.
      if (matchMedia('(max-width: 760px)').matches && element.closest('.home-view')) clip = document.querySelector('main').getBoundingClientRect();
      const [pixelX, pixelY] = this.pixelRatio;
      const left = Math.max(0, Math.floor(Math.max(rect.left, clip.left) * pixelX) - 1);
      const bottom = Math.max(0, Math.floor((this.height - Math.min(rect.bottom, clip.bottom)) * pixelY));
      const right = Math.min(this.canvas.width, Math.ceil(Math.min(rect.right, clip.right) * pixelX) + 1);
      const top = Math.min(this.canvas.height, Math.ceil((this.height - Math.max(rect.top, clip.top)) * pixelY));
      gl.scissor(left, bottom, Math.max(0, right - left), Math.max(0, top - bottom));
      gl.drawArrays(gl.TRIANGLES, 0, 3);
    });
    gl.disable(gl.SCISSOR_TEST);
  }
}
```

### Referenzdatei: motion.js

Vollständige, unveränderte Bewegungsimplementierung aus derselben Version. Sie setzt echte DOM-Elemente, einen dimensionierten positionierten Elterncontainer und die im Stylesheet festgelegte absolute Markerposition voraus.

```js
// One interruptible spring clock drives both the HTML silhouette and its optical
// lens. Different leading/trailing edge responses make a moving drop stretch.
const clamp = (value, low, high) => Math.min(high, Math.max(low, value));

class Spring {
  constructor(value = 0, stiffness = 460, damping = 30) {
    this.value = this.target = value;
    this.velocity = 0;
    this.stiffness = stiffness;
    this.damping = damping;
  }
  snap(value = this.target) { this.value = this.target = value; this.velocity = 0; }
  step(dt) {
    // Small fixed substeps keep pointer interruptions and slow frames stable.
    const count = Math.max(1, Math.ceil(dt / (1 / 180)));
    const step = dt / count;
    for (let i = 0; i < count; i++) {
      this.velocity += ((this.target - this.value) * this.stiffness - this.velocity * this.damping) * step;
      this.value += this.velocity * step;
    }
    if (Math.abs(this.target - this.value) < .012 && Math.abs(this.velocity) < .06) this.snap();
    return this.value !== this.target || this.velocity !== 0;
  }
}

export class LiquidMotion {
  constructor(render) {
    this.render = render;
    this.active = new Set();
    this.frame = 0;
    this.previous = 0;
    this.until = 0;
    this.preference = matchMedia('(prefers-reduced-motion: reduce)');
    this.preference.addEventListener('change', () => {
      if (this.preference.matches) this.finish();
    });
    document.addEventListener('visibilitychange', () => { if (document.hidden) this.finish(); });
  }
  start(controller) {
    this.active.add(controller);
    if (this.preference.matches || document.hidden) {
      // A preference change can precede its change event. Settle the whole
      // clock, including controllers that were already in flight.
      this.finish();
      return;
    }
    this.schedule();
  }
  animateFor(duration = 0) {
    if (document.hidden) { this.finish(); return; }
    this.until = Math.max(this.until, performance.now() + (this.preference.matches ? 0 : duration));
    this.schedule();
  }
  schedule() {
    if (!this.frame) {
      this.previous = performance.now();
      this.frame = requestAnimationFrame(time => this.tick(time));
    }
  }
  tick(time) {
    this.frame = 0;
    if (this.preference.matches || document.hidden) { this.finish(); return; }
    const dt = clamp((time - this.previous) / 1000, 1 / 240, 1 / 30);
    this.previous = time;
    for (const controller of this.active) if (!controller.step(dt)) this.active.delete(controller);
    // Draw in this very frame: never let the refraction trail the DOM silhouette.
    this.render();
    if (this.active.size || time < this.until) this.frame = requestAnimationFrame(next => this.tick(next));
  }
  finish() {
    cancelAnimationFrame(this.frame);
    this.frame = 0;
    this.until = 0;
    for (const controller of this.active) controller.finish();
    this.active.clear();
    this.render();
  }
}

export class LiquidMarker {
  constructor(element, motion, { container = null, radius = 24, clampEdges = true } = {}) {
    this.element = element;
    this.motion = motion;
    this.container = container;
    this.radius = radius;
    this.clampEdges = clampEdges;
    this.left = new Spring();
    this.right = new Spring();
    this.top = new Spring(0, 460, 31);
    this.height = new Spring(0, 460, 31);
    this.pressure = new Spring(1, 510, 25);
    this.initialized = false;
    this.springs = [this.left, this.right, this.top, this.height, this.pressure];
  }
  follow(target, immediate = false) {
    if (!target || !target.offsetWidth) {
      this.element.hidden = true;
      return;
    }
    this.element.hidden = false;
    const parent = this.element.offsetParent;
    const box = target.getBoundingClientRect();
    const origin = parent.getBoundingClientRect();
    const style = getComputedStyle(parent);
    const scaleX = origin.width / parseFloat(style.width);
    const scaleY = origin.height / parseFloat(style.height);
    this.to({ x: (box.x - origin.x) / scaleX, y: (box.y - origin.y) / scaleY, width: box.width / scaleX, height: box.height / scaleY }, immediate);
  }
  to(rect, immediate = false) {
    this.element.hidden = false;
    if (this.initialized && !immediate && Math.abs(rect.x - this.left.target) < .001 && Math.abs(rect.x + rect.width - this.right.target) < .001 && Math.abs(rect.y - this.top.target) < .001 && Math.abs(rect.height - this.height.target) < .001) return;
    const delta = rect.x - this.left.target;
    this.left.stiffness = delta >= 0 ? 370 : 570;
    this.right.stiffness = delta >= 0 ? 570 : 370;
    this.left.damping = delta >= 0 ? 27 : 31;
    this.right.damping = delta >= 0 ? 31 : 27;
    this.left.target = rect.x;
    this.right.target = rect.x + rect.width;
    this.top.target = rect.y;
    this.height.target = rect.height;
    this.pressure.target = 1;
    if (immediate || !this.initialized) {
      this.initialized = true;
      this.finish();
      this.motion.animateFor(0);
    } else {
      if (Math.abs(delta) > .1) this.pressure.velocity -= Math.min(2.8, Math.max(1.5, Math.abs(delta) / 80));
      this.element.dataset.liquidState = 'moving';
      this.motion.start(this);
    }
  }
  press() {
    if (!this.initialized || this.motion.preference.matches) return;
    this.pressure.target = .87;
    this.element.dataset.liquidState = 'moving';
    this.motion.start(this);
  }
  release() {
    if (!this.initialized) return;
    if (this.pressure.target === 1 && this.pressure.value === 1 && this.pressure.velocity === 0) return;
    this.pressure.target = 1;
    this.motion.start(this);
  }
  step(dt) {
    let moving = false;
    for (const spring of this.springs) moving = spring.step(dt) || moving;
    this.paint();
    if (!moving) this.element.dataset.liquidState = 'idle';
    return moving;
  }
  finish() {
    this.springs.forEach(spring => spring.snap());
    this.paint();
    this.element.dataset.liquidState = 'idle';
  }
  geometry() {
    const targetWidth = this.right.target - this.left.target;
    const speed = Math.abs((this.left.velocity + this.right.velocity) * .5);
    const pressure = clamp(this.pressure.value - speed / 14000, .78, 1.065);
    const width = clamp(this.right.value - this.left.value, targetWidth * .78, targetWidth + 70);
    const height = Math.max(1, this.height.value * pressure);
    let left = (this.left.value + this.right.value - width) * .5;
    if (this.clampEdges && this.container) left = clamp(left, 2, Math.max(2, this.container.clientWidth - width - 2));
    const top = this.top.value + (this.height.value - height) * .5;
    return { left, top, width, height, speed, pressure };
  }
  paint() {
    this.currentGeometry = this.geometry();
    const { left, top, width, height, speed, pressure } = this.currentGeometry;
    this.element.style.transform = `translate3d(${left.toFixed(3)}px,${top.toFixed(3)}px,0)`;
    this.element.style.width = `${width.toFixed(3)}px`;
    this.element.style.height = `${height.toFixed(3)}px`;
    this.element.style.borderRadius = `${Math.min(this.radius, height / 2).toFixed(3)}px`;
    if (this.container) {
      const strain = clamp(speed / 1200 + Math.abs(1 - pressure), 0, 1);
      this.container.style.scale = `${1 + strain * .008} ${1 - strain * .012}`;
    }
  }
}

// A knob expands under the finger, stretches in the direction of travel, and
// returns as one continuous lens. No separate CSS animation can drift from it.
export class LiquidKnob extends LiquidMarker {
  constructor(element, motion, options = {}) {
    super(element, motion, { radius: 40, clampEdges: false, ...options });
    this.engagement = new Spring(0, 440, 29);
    this.springs.push(this.engagement);
    this.dragging = false;
  }
  press() {
    if (!this.initialized || this.motion.preference.matches) return;
    this.engagement.target = 1;
    this.element.dataset.liquidState = 'moving';
    this.motion.start(this);
  }
  release() {
    if (!this.initialized) return;
    this.dragging = false;
    this.engagement.target = 0;
    super.release();
    if (this.engagement.value !== 0 || this.engagement.velocity !== 0) this.motion.start(this);
  }
  geometry() {
    const rect = super.geometry();
    const parent = this.element.parentElement;
    const parentStyle = getComputedStyle(parent);
    const parentWidth = parseFloat(parentStyle.width) || parent.clientWidth;
    const parentHeight = parseFloat(parentStyle.height) || parent.clientHeight;
    const size = this.right.target - this.left.target;
    const expansion = clamp(this.engagement.value, -.10, 1.08);
    const stretch = clamp(rect.speed / 650, 0, .18);
    const width = Math.min(parentWidth - 2, rect.width + size * (.30 * expansion + stretch));
    const height = Math.min(parentHeight - 2, rect.height * (1 + .09 * expansion));
    const left = clamp(rect.left + (rect.width - width) / 2, 1, Math.max(1, parentWidth - width - 1));
    return { ...rect, left, top: (parentHeight - height) / 2, width, height };
  }
}

export class LiquidToggle extends LiquidKnob {
  paint() {
    super.paint();
    const track = this.element.parentElement;
    const size = this.right.target - this.left.target;
    const inset = (track.clientHeight - size) / 2;
    const travel = Math.max(1, track.clientWidth - 2 * inset - size);
    const center = (this.left.value + this.right.value) / 2;
    track.style.setProperty('--switch-progress', clamp((center - inset - size / 2) / travel, 0, 1).toFixed(4));
  }
  bind(input) {
    let gesture = null;
    let suppressClick = false;
    const position = event => {
      const box = input.getBoundingClientRect();
      return clamp((event.clientX - box.left - box.height / 2) / Math.max(1, box.width - box.height), 0, 1);
    };
    input.addEventListener('pointerdown', event => {
      if (event.button !== 0) return;
      gesture = { id: event.pointerId, x: event.clientX, start: input.checked, moved: false };
      input.setPointerCapture(event.pointerId);
      this.press();
    });
    input.addEventListener('pointermove', event => {
      if (!gesture || gesture.id !== event.pointerId) return;
      gesture.moved ||= Math.abs(event.clientX - gesture.x) > 3;
      if (!gesture.moved) return;
      event.preventDefault();
      this.dragging = true;
      this.preview = position(event);
      const size = this.right.target - this.left.target;
      const track = this.element.parentElement;
      const inset = (track.clientHeight - size) / 2;
      const left = inset + (track.clientWidth - inset * 2 - size) * this.preview;
      this.left.snap(left);
      this.right.snap(left + size);
      this.paint();
      this.motion.animateFor(0);
    });
    const finish = (event, cancelled = false) => {
      if (!gesture || gesture.id !== event.pointerId) return;
      if (gesture.moved) {
        input.checked = cancelled ? gesture.start : this.preview >= .5;
        // Use the existing change handler so persistent state and optics agree.
        input.dispatchEvent(new Event('change', { bubbles: true }));
        suppressClick = !cancelled;
        if (suppressClick) setTimeout(() => { suppressClick = false; }, 0);
      }
      gesture = null;
      this.release();
    };
    input.addEventListener('pointerup', event => finish(event));
    input.addEventListener('pointercancel', event => finish(event, true));
    input.addEventListener('click', event => {
      if (!suppressClick) return;
      event.preventDefault();
      suppressClick = false;
    });
  }
}

export class LiquidRange extends LiquidKnob {
  constructor(element, motion, input, large) {
    super(element, motion);
    this.input = input;
    this.large = large;
  }
  update(immediate = false) {
    const parent = this.element.parentElement;
    const width = parent.checkVisibility() ? parseFloat(getComputedStyle(parent).width) || parent.clientWidth : 0;
    if (!width) { this.element.hidden = true; return; }
    const size = this.large ? 32 : 20;
    const percent = (Number(this.input.value) - Number(this.input.min)) / (Number(this.input.max) - Number(this.input.min));
    const x = (width - size - 4) * percent + 2;
    const rect = { x, y: 2, width: size, height: size };
    this.to(rect, immediate);
    if (this.dragging) {
      // Interpolate native value steps with a fast, damped follower. Snapping
      // erased all travel velocity and made the supposedly liquid knob rigid.
      this.left.stiffness = 1000; this.right.stiffness = 1100;
      this.left.damping = this.right.damping = 48;
    }
  }
  press() { this.dragging = true; super.press(); }
  paint() {
    super.paint();
    const track = this.element.parentElement;
    const center = this.currentGeometry.left + this.currentGeometry.width / 2;
    const width = parseFloat(getComputedStyle(track).width) || track.clientWidth;
    const fill = clamp(center / Math.max(1, width), 0, 1);
    track.style.setProperty('--visual-fill', `${fill * 100}%`);
    track.dataset.fill = fill.toFixed(5);
  }
}

export class LiquidSurface {
  constructor(element, motion) {
    this.element = element;
    this.motion = motion;
    this.scale = new Spring(1, 540, 32);
  }
  press() {
    if (this.motion.preference.matches) return;
    this.scale.target = .955;
    this.motion.start(this);
  }
  release() {
    if (this.scale.target === 1 && this.scale.value === 1) return;
    this.scale.target = 1;
    this.motion.start(this);
  }
  step(dt) { const active = this.scale.step(dt); this.paint(); return active; }
  paint() { this.element.style.scale = this.scale.value.toFixed(5); }
  finish() { this.scale.snap(1); this.paint(); }
}
```

## Do's and Don'ts

### Do

- **Do** zuerst die eingebettete Referenz zum Laufen bringen und danach das Framework integrieren.
- **Do** Glasgeometrie und DOM im selben Frame aktualisieren und einrastende Federn vollständig beenden.
- **Do** Pointer-Cancel, Loslassen außerhalb des Controls und schnelle Richtungswechsel prüfen.
- **Do** tatsächliche Canvasmaße, beide effektiven Pixelverhältnisse und sichtbare Scrollbalken berücksichtigen.
- **Do** klare SVGs mit zentrierten gezeichneten Grenzen und echte native Inputs verwenden.
- **Do** die originale Projektlogik, vorhandene Daten und alte Builds erhalten.

### Don't

- **Don't** Kristall mit starkem `backdrop-filter: blur(...)` oder milchigem Weiß ersetzen.
- **Don't** Rahmen, Inset-Shadows und mehrere Shaderkonturen übereinanderlegen.
- **Don't** den Brechungsfaktor erhöhen, bis das Hintergrundbild zurückfaltet und Formen mehrfach erscheinen.
- **Don't** Reglerfüllungen senkrecht abschneiden oder den breiten Knopf als deckende zweite Fotolinse rendern.
- **Don't** Marker bei jedem Klick neu erzeugen, native Dragwerte hart snappen oder alte Animationen erst ausspielen lassen.
- **Don't** DOM-only-Zentrierungstests als Beweis für korrekt gezeichnetes Glas ausgeben.
- **Don't** beliebige HTML-Brechung, native Apple-Identität oder Pixelgleichheit bei anderem Hintergrund versprechen.

### Abnahme durch den Coding Agent

| Prüfung | Erwartung |
| --- | --- |
| Shaderstart | WebGL-Programm kompiliert; keine Laufzeit- oder GL-Fehler |
| Kristall überall | Werte `1 / 0 / .08`, keine komponentenweisen Materialvarianten |
| Desktop mit Scrollbalken | Kurzes Fenster mit echtem klassischem Scrollbalken; Canvas-CSSbreite darf von `innerWidth` abweichen |
| Pixeldichten | DPR `1`, `1.25` und `2`; tatsächliche Rasterkonturen mit HTML-Zentren vergleichen |
| Rasterzentrierung | In der Ursprungsprüfung Abweichung unter `1.1 CSS-px`; auf neutraler Testtextur per `readPixels()` gemessen |
| Navigation | Bruchteilige Button- und Markergrenzen stimmen nach dem Einrasten überein; in der Ursprungsprüfung `< .02px` |
| Übergänge | Sichtbare Zwischenzustände, begrenzte Verformung, leichte Dämpfung, schnelle Umlenkung ohne Duplikate |
| Große Regler | 0/50/100 %, Maus, Touch und Tastatur; weiche runde Füllung; Füllmitte folgt animiertem Knopf |
| Regler-Raster | Gleichmäßige Testtextur verwenden; im inneren Zentrum keine harte Helligkeitskante. Ursprungsgrenze: benachbarter Luminanzschritt `< 13/255`; kein Test über ein kontrastreiches Foto |
| Toggle | Beide Zustände, Halten, Ziehen, Doppelklickfolge, Tastatur, Abbruch und Loslassen außerhalb |
| Reduzierte Bewegung | Vor Start und während laufender Animation umschalten; keine verbleibenden aktiven Federn |
| Mobil / Overflow | Mindestens `390 × 844` und ein schmales Layout; HTML und Canvas werden gleich geclippt |
| Fallback | WebGL abschalten; Inhalte, Formen, Werte, Eingabe und Fokus bleiben nutzbar |
| Optionale Wasserbewegung | Läuft nur bei sichtbarer Wiedergabe; Pause, Hintergrund und Reduced Motion stoppen sie |

Bei Playwright ausdrücklich mit sichtbaren Scrollbalken starten:

```js
const browser = await chromium.launch({
  channel: 'msedge', // Oder die tatsächlich zu prüfende installierte Chromium-Variante.
  headless: true,
  ignoreDefaultArgs: ['--hide-scrollbars']
});
```

Die Freigabe der Ausgangsanwendung basierte auf 56 bestandenen Browserprüfungen und einer separaten Prüfung des fertigen Builds. Diese Zahl ist Herkunftsevidenz der Referenz und kein automatischer Testnachweis für eine neue Integration. Das Zielprojekt muss seine eigene Abnahme durchführen.

### Herkunft und Integrität

Die beiden vollständigen Kernmodule sind direkt aus dem freigegebenen lokalen Quellstand übernommen. Die folgenden SHA-256-Werte beziehen sich auf UTF-8-Text mit LF-Zeilenenden und genau einem abschließenden Zeilenumbruch. HTML, CSS und Bootstrap sind das eigenständige Prüfgerüst dieser Übergabe; die Shader- und Federimplementierung bleibt unverändert.

| Referenzmodul | SHA-256 (normalisierter Text) |
| --- | --- |
| glass.js | `3842d223f81ba5880c25b469b98b25694c8f2ceed2b7db7de4efd893eb988922` |
| motion.js | `8540cbc564751d0942e167442191a0b0f22a0d508c18ee318bbebe4302432adc` |

Diese Datei ist eine portable Materialreferenz neben der bestehenden projektbezogenen `DESIGN.md`. Sie benötigt deren Sidecar oder Layoutbeschreibung nicht. Für die Weitergabe genügt diese Datei plus ein eigenes Hintergrundbild; die Codeblöcke sind vollständig und enthalten keine absichtlich ausgelassenen Implementierungsteile.

Die aus dieser Datei extrahierte Standalone-Referenz wurde am 22. Oktober 2026 separat geprüft: **14 Prüfungen bestanden**, darunter Modulidentität, Syntax, echte Shaderwerte, sichtbare Scrollbalken bei gebrochener Pixeldichte, Navigation, native Regler, beide Toggle-Größen, Wasserbewegung, mobile Darstellung, reduzierte Bewegung und CSS-Fallback. YAML-Struktur, alle sieben Tokenreferenzen und die acht kanonischen Abschnitte wurden zusätzlich validiert. Browsernachweis: installierter Microsoft Edge; Desktop, Mobil und Fallback wurden auch visuell kontrolliert.
