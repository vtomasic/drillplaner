# Handover: Übungsplaner (Fussball-Übungseditor)

## Kontext

Web-Tool zum Zeichnen von Fussball-Trainingsübungen (Jugendtraining, SFV/J+S-Kontext).
Zielworkflow: Übungen zeichnen → als SVG speichern → in Obsidian-Vault ablegen
(SVG rendert direkt in Notizen, Metadaten via Dataview/Frontmatter der umgebenden Notiz).

Nutzer-Präferenzen: minimal lauffähige Systeme zuerst, iterativ erweitern,
kein Over-Engineering, keine unnötigen Dependencies.

## Aktueller Stand

Ein einziges React-Component-File: `src/uebungsplaner.jsx` (Default-Export `Uebungsplaner`).
Kein Backend, kein State-Management-Framework. Einzige Library neben React: `docx`
(Word-Export — handgerolltes OOXML/ZIP wäre unverhältnismässig).

### Setup (Vite) — eingerichtet

Vite-Projekt liegt direkt im Ordner (manuell aufgesetzt, kein Scaffolder):
`package.json`, `vite.config.js`, `index.html`, `src/main.jsx` (mountet die
Komponente direkt, kein App.jsx), `src/uebungsplaner.jsx`.

```bash
npm install   # einmalig
npm run dev   # dann angezeigte URL öffnen (Standard: localhost:5173)
```

## Architektur

### Koordinatensystem
- SVG viewBox `0 0 900 600`, Rand `M = 25`
- Alle Positionen in viewBox-Einheiten (nicht Pixel)
- Pointer → viewBox: `getPoint(e)` via `getScreenCTM().inverse()` + `DOMPoint` —
  NICHT über getBoundingClientRect skalieren: das ignoriert das Letterboxing von
  `preserveAspectRatio` und erzeugt Versatz an den Feldrändern

### State (alles in einer Komponente)
```js
items: Array<Item>        // einheitliche Liste, Reihenfolge = Zeichenreihenfolge (für Undo)
mode: string              // 'select' | 'el:<type>' | 'line:<type>'
field: 'half'|'full'|'penalty'|'blank'
                          // penalty = Strafraum-Ansicht, Tor unten, volle Feldbreite
                          // bis zu den Eckfahnen; UNIFORM 12.5 E/m (850 E = 68 m),
                          // Sichttiefe 44 m — einziges massstabsgetreues Feld
selectedId: number|null   // nur Elemente selektierbar, Linien nicht
draft: Line|null          // Linie während des Ziehens
dragRef: useRef           // {id, dx, dy} beim Verschieben
```

### Item-Typen
```js
// Element
{ kind:'el', id, type:'player'|'gk'|'ball'|'cone'|'pole'|'goal'|'ladder',
  x, y, label, color? }   // label: '' bei player (keine Nummern), 'T' bei gk
                          // color nur bei player — runder Farbpunkt neben dem
                          // Spieler-Button öffnet ein Popup mit 6 runden Swatches
                          // (PLAYER_COLORS); Swatch-Klick setzt Farbe, schliesst
                          // Popup und aktiviert den Spieler-Modus. Unsichtbares
                          // Fixed-Overlay schliesst das Popup bei Klick daneben.
                          // Legacy-Typen 'p1'/'p2' rendern weiterhin (Fallback COL[type]),
                          // alte Nummern-Labels bleiben sichtbar

// Linie (zwei Punkte)
{ kind:'line', id, lineType:'run'|'pass'|'dribble'|'cross'|'shot'|'plain'|'rect',
  x1, y1, x2, y2, color? } // color optional: überschreibt COL[lineType];
                           // Pfeil-Marker je Palettenfarbe vordefiniert (ah-<hex>)

// Freihand
{ kind:'line', id, lineType:'free', points:[{x,y}, ...] }
```

### Linien-Rendering (Trainer-Notation)
| Typ | Darstellung |
|---|---|
| run (Laufweg) | weiss, durchgezogen, Pfeil |
| pass | gelb, gestrichelt, Pfeil |
| dribble | weiss, Zickzack (`zigzagPath()`), Pfeil |
| cross (Flanke) | gelb, gebogen (`curvePath()`, quadratische Bézier, Offset 0.28·len), gestrichelt |
| shot (Schuss) | rot, dick (4.5), Pfeil |
| free (Freihand) | gelb, Midpoint-Smoothing (`freePath()`), kein Pfeil |
| plain (Linie) | weiss, durchgezogen, KEIN Pfeil |
| rect (Rechteck) | gelb, geschlossener Pfad (x1/y1 → x2/y2 als Ecken), kein Pfeil; für Zonen |

Pfeilspitzen: `<marker>`-Defs pro Farbe (`ah-white/yellow/red` + `ah-<hex>` je
Palettenfarbe) — SVG-Marker können Stroke-Farbe nicht zuverlässig erben.
Farbwahl: wiederverwendbare `ColorPicker`-Komponente (runder Punkt + Popup mit den
6 PALETTE-Farben) — beim Spieler neben dem Button, bei Wegen & Formen neben der
Gruppen-Überschrift. `lineColor` startet mit Weiss `#f5f2e8`; die typspezifischen
COL-Farben greifen nur noch als Fallback für Items ohne `color` (Legacy-Dateien).

### Freihand-Details
- Punkt wird nur aufgenommen wenn Distanz zum letzten ≥ 3 viewBox-Einheiten (Pfad schlank halten)
- Koordinaten auf 1 Dezimale gerundet (JSON-Grösse)
- Commit nur bei > 2 Punkten

### Persistenz: JSON = Projektdatei, SVG/PNG/Word = Exporte
**„Übung speichern"** erzeugt `{version:1, type:"uebung", meta, field, items}` als
`.json` — das ist die Projektdatei, konsistent zum Trainings-Format.
**„Übung laden"** akzeptiert Übungs-JSON, Trainings-JSON (Weiche auf `uebungen`)
und Legacy-SVG.

Der SVG-Export enthält den Zustand zusätzlich als JSON in einem `<metadata>`-Element
**im gerenderten SVG** (Legacy-Ladepfad, schadet nicht):

```xml
<metadata id="uebung-data">{"version":1,"field":"half","meta":{...},"items":[...]}</metadata>
```

`meta` = Übungsdaten aus dem rechten Panel: `titel, dauer, mannschaften, material,
ziel, beschreibung, variation` (alles Strings; `beschreibung`/`variation` mehrzeilig).
Beim Laden wird mit `EMPTY_META` gemergt — alte Dateien ohne `meta` laden sauber.
Dateiname beim Export = Slug aus dem Titel (Umlaute → ae/oe/ue), Fallback `uebung`.

- **SVG-Export:** `XMLSerializer` auf das Live-SVG → Download `.svg`
- **Laden von SVG:** `DOMParser` → `getElementById('uebung-data')` → `JSON.parse(textContent)`
- Nach dem Laden: `nextId = max(ids) + 1` (Modul-Level-Counter!)
- **PNG-Export** (2×-Auflösung via Canvas) enthält KEINE Metadaten — nur SVG ist die Projektdatei
- **Word-Export** (`exportDOCX`): echtes `.docx` via `docx`-Library — eine Tabelle pro Übung:
  Titelzeile, gross das Bild (PNG 2×, via `renderPNGBlob()`), darunter nur die
  ausgefüllten Metadaten-Felder. `uebungTable(meta, pngBytes)` ist bewusst eine
  freistehende Funktion → für den geplanten Trainings-Export (mehrere Übungen in
  einem Dokument) direkt wiederverwendbar
- `version: 1` im Format — bei Schema-Änderungen Migrationspfad einbauen

### Interaktion
- Element-Modus: Klick platziert (Modus bleibt aktiv für Mehrfachplatzierung)
- Linien-Modus: Drag von Start zu Ende (min. Länge 15, sonst verworfen)
- Select-Modus: Elemente und Linien draggen; Klick auf Element/Linie in beliebigem
  Nicht-Linien-Modus wechselt zu Select
- Linien-Hit-Testing: unsichtbarer transparenter Stroke (Breite 16) über jedem Pfad;
  Verschieben verschiebt beide Endpunkte bzw. alle Freihand-Punkte um den Drag-Delta
  (Delta relativ zur Startposition, kein kumulatives Runden)
- Endpunkt-Editing: gelbe Handles an beiden Endpunkten (Zwei-Punkt-Linien) bzw.
  allen 4 Ecken (Rechteck; Ecke → [xKey,yKey]-Paar, z.B. oben-rechts = x2/y1);
  Handle-Drag setzt nur diese Koordinaten. Freihand hat keine Handles.
  Hover-Detektion im Select-Modus: `onMove` sucht den nächsten End-/Eckpunkt aller
  Linien im Fangradius 12 E (`handleKeys()`) → Handle erscheint beim Annähern ohne
  vorheriges Selektieren, unsichtbare Grabfläche r=14; Zugriff selektiert die Linie
- Tastatur: `Delete`/`Backspace` löscht Auswahl, `Escape` deselektiert
  (Guard: Tippen in Input/Textarea löst KEIN Löschen aus)
- Pointer-Events + `touchAction:'none'` → funktioniert mit Touch/Apple Pencil
- Undo = letztes Item entfernen (`items.slice(0,-1)`), kein Redo

### Training (mehrere Übungen bündeln)
- State: `training: Array<{meta, field, items}>` (Snapshots via `structuredClone`),
  `trainingTitel`, `activeIdx` (welcher Eintrag im Editor liegt)
- UI im rechten Panel unter den Übungsdaten: „+ Übung hinzufügen" nimmt den aktuellen
  Editor-Stand; Klick auf Eintrag lädt ihn zurück in den Editor (markiert aktiv);
  „Übung N aktualisieren" überschreibt den aktiven Eintrag; ↑/✕ für Reihenfolge/Entfernen
- **Speichern/Laden:** eigene JSON-Datei
  `{version:1, type:"training", titel, uebungen:[{meta, field, items}]}` —
  der normale „Übung laden"-Button erkennt Trainings-Dateien ebenfalls (Weiche auf
  `Array.isArray(data.uebungen)`)
- **Training → Word:** ein Dokument, Trainings-Titel als Überschrift, pro Übung
  `uebungTable()` mit Seitenumbruch dazwischen. Übungen werden offscreen gerendert:
  `uebungSvgString(field, items)` via `renderToStaticMarkup` (react-dom/server) →
  `svgToPng()` — kein Umweg über den Editor-State

## UI-Konventionen
- Dark-Slate-Chrome (#1c2320), Signalgelb-Akzent (#ffd447), ui-monospace für Labels
- Deutsche Beschriftung (Schweizer Nutzer: kein ß — "Fussball", "weiss")
- Statusmeldungen (Laden erfolgreich/fehlgeschlagen) in der Fusszeile, kein alert()
- Inline-Styles (bewusst — kein Tailwind, kein CSS-File; bei Wachstum ggf. umziehen)

## Backlog (priorisiert, mit Nutzer besprochen)

1. **Rotation von Elementen** — Leiter/Minitore quer stellen (z.B. Taste R bei Auswahl
   oder Rotations-Handle; `rotation`-Feld im Item ergänzen → Format-Version beachten)
2. **Textlabels** — freie Beschriftungen für Zonen, Abstände ("10m"), Übungstitel aufs Feld
3. ~~Linien selektier-/löschbar machen~~ — ERLEDIGT (selektierbar, verschiebbar,
   via Entf/Auswahl-löschen entfernbar)
4. **Obsidian-Integration** — Ziel: Übungs-SVG + Markdown-Notiz mit Frontmatter
   (kategorie, altersstufe, spielerzahl, dauer, material) für Dataview-Abfragen;
   evtl. Export-Button "Notiz + SVG" der beides generiert
5. ~~Trainings-Export~~ — ERLEDIGT (siehe Abschnitt Training)
6. **Animation** — Positions-Keyframes speichern und abspielen (grösster Brocken, zuletzt)
7. **PWA/Build** — `npm run build` → statisch auf Mac Mini hosten, dann auch
   auf dem iPhone am Platz nutzbar

## Bekannte Einschränkungen
- Linien: nur Verschieben als Ganzes, Endpunkte nicht einzeln editierbar
- Kein Redo
- `nextId` ist Modul-Level — bei HMR (Vite Hot Reload) kann der Counter zurückspringen;
  Laden setzt ihn korrekt, aber bei Bedarf in useRef umziehen
- Element-Labels nicht editierbar (Auto-Nummerierung only)
- Kein Zoom/Pan der Zeichenfläche
