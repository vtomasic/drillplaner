import { useState, useRef, useEffect } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import {
  Document, Packer, Paragraph, TextRun, ImageRun, PageBreak,
  Table, TableRow, TableCell, WidthType,
} from "docx";

// ── Konstanten ────────────────────────────────────────────────
const VW = 900, VH = 600, M = 25; // viewBox + Rand
const FW = VW - 2 * M, FH = VH - 2 * M;

const PALETTE = ["#2f6fde", "#d64545", "#ffd447", "#2e9e4f", "#f5f2e8", "#222222"];

const EL_DEFS = [
  { key: "gk",   label: "Torwart" },
  { key: "ball", label: "Ball" },
  { key: "cone", label: "Hütchen" },
  { key: "pole", label: "Stange" },
  { key: "goal", label: "Minitor" },
  { key: "ladder", label: "Leiter" },
];

const LINE_DEFS = [
  { key: "run",     label: "Laufweg",  hint: "weiss, durchgezogen" },
  { key: "pass",    label: "Pass",     hint: "gelb, gestrichelt" },
  { key: "dribble", label: "Dribbling", hint: "weiss, Zickzack" },
  { key: "cross",   label: "Flanke",   hint: "gelb, gebogen" },
  { key: "shot",    label: "Schuss",   hint: "rot, dick" },
  { key: "free",    label: "Freihand", hint: "frei zeichnen" },
  { key: "plain",   label: "Linie",    hint: "einfache Linie, ohne Pfeil" },
  { key: "rect",    label: "Rechteck", hint: "Zone markieren (Ecke zu Ecke ziehen)" },
];

const META_FIELDS = [
  ["titel", "Titel"],
  ["ziel", "Ziel"],
  ["dauer", "Dauer"],
  ["beschreibung", "Beschreibung", true],
  ["variation", "Variation", true],
  ["mannschaften", "Mannschaften"],
  ["material", "Material"],
];
const EMPTY_META = Object.fromEntries(META_FIELDS.map(f => [f[0], ""]));

// Eine Übung als Word-Tabelle (wiederverwendbar für späteren Trainings-Export)
function uebungTable(meta, pngBytes) {
  const cell = (children, opts = {}) => new TableCell({ ...opts, children });
  const rows = [
    new TableRow({
      children: [cell([new Paragraph({ children: [new TextRun({ text: meta.titel || "Übung", bold: true, size: 32 })] })], { columnSpan: 2 })],
    }),
    new TableRow({
      children: [cell([new Paragraph({ children: [new ImageRun({ type: "png", data: pngBytes, transformation: { width: 620, height: 413 } })] })], { columnSpan: 2 })],
    }),
    ...META_FIELDS
      .filter(([key]) => key !== "titel" && meta[key].trim())
      .map(([key, label]) => new TableRow({
        children: [
          cell([new Paragraph({ children: [new TextRun({ text: label, bold: true })] })], { width: { size: 22, type: WidthType.PERCENTAGE } }),
          cell(meta[key].split("\n").map(line => new Paragraph(line))),
        ],
      })),
  ];
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows });
}

const COL = {
  p1: "#2f6fde", p2: "#d64545", gk: "#e8a020",
  run: "#f5f2e8", pass: "#ffd447", dribble: "#f5f2e8",
  cross: "#ffd447", shot: "#ff5b4d", free: "#ffd447",
  plain: "#f5f2e8", rect: "#ffd447",
};

let nextId = 1;
const uid = () => nextId++;

const slugify = (s) => s.trim().toLowerCase()
  .replace(/[äöüß]/g, ch => ({ ä: "ae", ö: "oe", ü: "ue", ß: "ss" }[ch]))
  .replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");

// SVG-XML → PNG-Blob (2×-Auflösung via Canvas)
function svgToPng(xml) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(new Blob([xml], { type: "image/svg+xml;charset=utf-8" }));
    const img = new Image();
    img.onload = () => {
      const c = document.createElement("canvas");
      c.width = VW * 2; c.height = VH * 2;
      c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
      URL.revokeObjectURL(url);
      c.toBlob(b => b ? resolve(b) : reject(new Error("PNG-Erzeugung fehlgeschlagen")));
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("SVG-Rendering fehlgeschlagen")); };
    img.src = url;
  });
}

// ── Geometrie-Helfer ──────────────────────────────────────────
function zigzagPath(x1, y1, x2, y2) {
  const dx = x2 - x1, dy = y2 - y1, len = Math.hypot(dx, dy);
  if (len < 12) return `M${x1},${y1} L${x2},${y2}`;
  const ux = dx / len, uy = dy / len, px = -uy, py = ux;
  const amp = 5, step = 11, n = Math.max(2, Math.floor((len - 14) / step));
  let d = `M${x1},${y1}`;
  for (let i = 1; i <= n; i++) {
    const t = i * step, off = i % 2 ? amp : -amp;
    d += ` L${(x1 + ux * t + px * off).toFixed(1)},${(y1 + uy * t + py * off).toFixed(1)}`;
  }
  return d + ` L${x2},${y2}`;
}
function curvePath(x1, y1, x2, y2) {
  const mx = (x1 + x2) / 2, my = (y1 + y2) / 2;
  const dx = x2 - x1, dy = y2 - y1, len = Math.hypot(dx, dy) || 1;
  const cx = mx - (dy / len) * len * 0.28, cy = my + (dx / len) * len * 0.28;
  return `M${x1},${y1} Q${cx},${cy} ${x2},${y2}`;
}
function freePath(pts) {
  if (!pts || pts.length < 2) return "";
  let d = `M${pts[0].x},${pts[0].y}`;
  for (let i = 1; i < pts.length - 1; i++) {
    const mx = (pts[i].x + pts[i + 1].x) / 2, my = (pts[i].y + pts[i + 1].y) / 2;
    d += ` Q${pts[i].x},${pts[i].y} ${mx.toFixed(1)},${my.toFixed(1)}`;
  }
  const last = pts[pts.length - 1];
  return d + ` L${last.x},${last.y}`;
}

// ── Spielfelder ───────────────────────────────────────────────
function PitchStripes() {
  const stripes = [];
  for (let i = 0; i < 10; i++)
    stripes.push(
      <rect key={i} x={(VW / 10) * i} y={0} width={VW / 10} height={VH}
        fill={i % 2 ? "#3d8046" : "#41874b"} />
    );
  return <g>{stripes}</g>;
}

function PitchLines({ field }) {
  const s = { fill: "none", stroke: "#f5f2e8", strokeWidth: 2.5, opacity: 0.9 };
  if (field === "blank") return null;
  if (field === "full") {
    const pbW = 110, pbH = 264, gaW = 40, gaH = 120, cy = VH / 2;
    return (
      <g {...s}>
        <rect x={M} y={M} width={FW} height={FH} />
        <line x1={VW / 2} y1={M} x2={VW / 2} y2={VH - M} />
        <circle cx={VW / 2} cy={cy} r={55} />
        <circle cx={VW / 2} cy={cy} r={3} fill="#f5f2e8" />
        {/* links */}
        <rect x={M} y={cy - pbH / 2} width={pbW} height={pbH} />
        <rect x={M} y={cy - gaH / 2} width={gaW} height={gaH} />
        <circle cx={M + 72} cy={cy} r={2.5} fill="#f5f2e8" />
        <path d={`M ${M + pbW} ${cy - 42} A 48 48 0 0 1 ${M + pbW} ${cy + 42}`} />
        <rect x={M - 9} y={cy - 28} width={9} height={56} strokeWidth={3} />
        {/* rechts */}
        <rect x={VW - M - pbW} y={cy - pbH / 2} width={pbW} height={pbH} />
        <rect x={VW - M - gaW} y={cy - gaH / 2} width={gaW} height={gaH} />
        <circle cx={VW - M - 72} cy={cy} r={2.5} fill="#f5f2e8" />
        <path d={`M ${VW - M - pbW} ${cy - 42} A 48 48 0 0 0 ${VW - M - pbW} ${cy + 42}`} />
        <rect x={VW - M} y={cy - 28} width={9} height={56} strokeWidth={3} />
      </g>
    );
  }
  if (field === "penalty") {
    // Strafraum-Ansicht, Tor unten, volle Feldbreite bis zu den Eckfahnen.
    // UNIFORMER Massstab: 850 E = 68 m → 12.5 E/m; Sichttiefe = 550/12.5 = 44 m.
    // Alle Masse real: Strafraum 40.32×16.5, Torraum 18.32×5.5, Punkt 11 m,
    // Teilkreis r = 9.15 m, Tor 7.32 m, Eckbögen 1 m
    const cx = VW / 2, gl = VH - M;
    const pbW = 504, pbH = 206, gaW = 229, gaH = 69;
    const spotY = gl - 137.5, boxTop = gl - pbH;
    // Teilkreis (r 114.4) schneidet Strafraumkante: Halbsehne √(114.4²−68.5²) ≈ 91.6
    return (
      <g {...s}>
        <line x1={M} y1={gl} x2={VW - M} y2={gl} />
        <line x1={M} y1={M} x2={M} y2={gl} />
        <line x1={VW - M} y1={M} x2={VW - M} y2={gl} />
        {/* Eckbögen */}
        <path d={`M ${M} ${gl - 12.5} A 12.5 12.5 0 0 1 ${M + 12.5} ${gl}`} />
        <path d={`M ${VW - M - 12.5} ${gl} A 12.5 12.5 0 0 1 ${VW - M} ${gl - 12.5}`} />
        <rect x={cx - pbW / 2} y={boxTop} width={pbW} height={pbH} />
        <rect x={cx - gaW / 2} y={gl - gaH} width={gaW} height={gaH} />
        <circle cx={cx} cy={spotY} r={3} fill="#f5f2e8" />
        <path d={`M ${cx - 91.6} ${boxTop} A 114.4 114.4 0 0 1 ${cx + 91.6} ${boxTop}`} />
        <rect x={cx - 46} y={gl} width={92} height={10} strokeWidth={3} />
      </g>
    );
  }
  // Halbfeld: Tor links, Mittellinie rechts
  const pbW = 170, pbH = 340, gaW = 60, gaH = 160, cy = VH / 2;
  return (
    <g {...s}>
      <rect x={M} y={M} width={FW} height={FH} />
      <rect x={M} y={cy - pbH / 2} width={pbW} height={pbH} />
      <rect x={M} y={cy - gaH / 2} width={gaW} height={gaH} />
      <circle cx={M + 112} cy={cy} r={3} fill="#f5f2e8" />
      <path d={`M ${M + pbW} ${cy - 62} A 70 70 0 0 1 ${M + pbW} ${cy + 62}`} />
      <path d={`M ${VW - M} ${cy - 70} A 70 70 0 0 0 ${VW - M} ${cy + 70}`} />
      <rect x={M - 10} y={cy - 34} width={10} height={68} strokeWidth={3} />
    </g>
  );
}

// ── Element-Rendering ─────────────────────────────────────────
function Element({ el, selected, onDown }) {
  const common = { transform: `translate(${el.x},${el.y})`, onPointerDown: onDown, style: { cursor: "grab" } };
  const sel = selected && <circle r={20} fill="none" stroke="#ffd447" strokeWidth={1.5} strokeDasharray="4 4" />;
  switch (el.type) {
    case "player": case "p1": case "p2": case "gk": // p1/p2 = Legacy-Dateien
      return (
        <g {...common}>
          {sel}
          <circle r={12} fill={el.color || COL[el.type]} stroke="#00000055" strokeWidth={1.5} />
          <text y={4.5} textAnchor="middle" fontSize={12} fontWeight={700} fill="#fff"
            style={{ userSelect: "none", fontFamily: "ui-monospace, monospace" }}>{el.label}</text>
        </g>
      );
    case "ball":
      return (
        <g {...common}>{sel}
          <circle r={6} fill="#fff" stroke="#222" strokeWidth={1.5} />
          <path d="M-2.5,-1 L2.5,-1 L1.5,2 L-1.5,2 Z" fill="#222" />
        </g>
      );
    case "cone":
      return (
        <g {...common}>{sel}
          <path d="M0,-9 L8,7 L-8,7 Z" fill="#ff8c1a" stroke="#00000044" strokeWidth={1} />
          <line x1={-5} y1={2} x2={5} y2={2} stroke="#fff" strokeWidth={2} />
        </g>
      );
    case "pole":
      return (
        <g {...common}>{sel}
          <rect x={-2} y={-13} width={4} height={26} rx={2} fill="#e8452f" />
          <rect x={-2} y={-5} width={4} height={5} fill="#fff" />
        </g>
      );
    case "goal":
      return (
        <g {...common}>{sel}
          <path d="M-16,-8 L-16,8 M-16,-8 L16,-8 L16,8" fill="none" stroke="#f5f2e8" strokeWidth={4} />
          <path d="M-16,-8 L16,-8 L16,8 L-16,8" fill="none" stroke="#f5f2e888" strokeWidth={1} strokeDasharray="2 3" />
        </g>
      );
    case "ladder":
      return (
        <g {...common}>{sel}
          <rect x={-24} y={-8} width={48} height={16} fill="none" stroke="#ffd447" strokeWidth={2} />
          {[-12, 0, 12].map(x => <line key={x} x1={x} y1={-8} x2={x} y2={8} stroke="#ffd447" strokeWidth={2} />)}
        </g>
      );
    default: return null;
  }
}

// Editierbare Punkte einer Zwei-Punkt-Linie bzw. eines Rechtecks (Freihand: keine)
function handleKeys(ln) {
  if (ln.kind !== "line" || ln.points) return [];
  return ln.lineType === "rect"
    ? [["x1", "y1"], ["x2", "y1"], ["x2", "y2"], ["x1", "y2"]]
    : [["x1", "y1"], ["x2", "y2"]];
}

function linePathD(ln) {
  switch (ln.lineType) {
    case "dribble": return zigzagPath(ln.x1, ln.y1, ln.x2, ln.y2);
    case "cross":   return curvePath(ln.x1, ln.y1, ln.x2, ln.y2);
    case "free":    return freePath(ln.points);
    case "rect":    return `M${ln.x1},${ln.y1} L${ln.x2},${ln.y1} L${ln.x2},${ln.y2} L${ln.x1},${ln.y2} Z`;
    default:        return `M${ln.x1},${ln.y1} L${ln.x2},${ln.y2}`;
  }
}

function Line({ ln, selected, onDown }) {
  const c = ln.color || COL[ln.lineType];
  const d = linePathD(ln);
  const marker = ["free", "plain", "rect"].includes(ln.lineType) ? undefined
    : ln.color ? `url(#ah-${ln.color.slice(1)})`
    : ln.lineType === "shot" ? "url(#ah-red)"
    : ln.lineType === "pass" || ln.lineType === "cross" ? "url(#ah-yellow)"
    : "url(#ah-white)";
  const width = ln.lineType === "shot" ? 4.5 : ln.lineType === "dribble" ? 2.2 : 2.5;
  const dash = ln.lineType === "pass" ? "8 6" : ln.lineType === "cross" ? "10 5" : undefined;
  return (
    <g>
      {selected && <path d={d} fill="none" stroke="#ffd447" strokeWidth={width + 7} opacity={0.35} strokeLinecap="round" strokeLinejoin="round" />}
      <path d={d} fill="none" stroke={c} strokeWidth={width} strokeDasharray={dash} markerEnd={marker} strokeLinecap="round" strokeLinejoin="round" />
      {onDown && <path d={d} fill="none" stroke="transparent" strokeWidth={16} onPointerDown={onDown} style={{ cursor: "grab" }} />}
    </g>
  );
}

function ArrowDefs() {
  const defs = [
    ["ah-white", COL.run], ["ah-yellow", COL.pass], ["ah-red", COL.shot],
    ...PALETTE.map(c => ["ah-" + c.slice(1), c]), // Marker je Palettenfarbe
  ];
  return (
    <defs>
      {defs.map(([id, c]) => (
        <marker key={id} id={id} markerWidth={9} markerHeight={9} refX={7} refY={4.5} orient="auto">
          <path d="M0,0 L9,4.5 L0,9 Z" fill={c} />
        </marker>
      ))}
    </defs>
  );
}

// Runder Farbpunkt mit Popup (Spieler- und Linienfarbe)
function ColorPicker({ color, onPick, title }) {
  const [open, setOpen] = useState(false);
  const dot = (bg, active) => ({
    width: 24, height: 24, borderRadius: "50%", background: bg, padding: 0,
    border: active ? "2px solid #ffd447" : "1px solid #3a423d", cursor: "pointer",
  });
  const pick = (c) => { onPick(c); setOpen(false); };
  return (
    <div style={{ position: "relative", flexShrink: 0 }}>
      <button title={title} onClick={() => setOpen(o => !o)}
        style={{ ...dot(color, open), width: 26, height: 26, display: "block" }} />
      {open && (
        <>
          {/* Klick daneben schliesst das Popup */}
          <div style={{ position: "fixed", inset: 0, zIndex: 5 }} onClick={() => setOpen(false)} />
          <div style={{
            position: "absolute", top: "100%", right: 0, marginTop: 4, zIndex: 6,
            background: "#242b27", border: "1px solid #3a423d", borderRadius: 8,
            padding: 8, display: "grid", gridTemplateColumns: "repeat(3, 24px)", gap: 8,
            boxShadow: "0 4px 16px #0008",
          }}>
            {PALETTE.map(c => (
              <button key={c} title="Farbe" onClick={() => pick(c)} style={dot(c, color === c)} />
            ))}
          </div>
        </>
      )}
    </div>
  );
}

// Übung offscreen als SVG-String rendern (für Word-Export von Trainings-Übungen,
// die nicht im Editor geladen sind)
function uebungSvgString(field, items) {
  return renderToStaticMarkup(
    <svg xmlns="http://www.w3.org/2000/svg" viewBox={`0 0 ${VW} ${VH}`}>
      <ArrowDefs />
      <PitchStripes />
      <PitchLines field={field} />
      {items.filter(i => i.kind === "line").map(ln => <Line key={ln.id} ln={ln} />)}
      {items.filter(i => i.kind === "el").map(el => <Element key={el.id} el={el} />)}
    </svg>
  );
}

// ── Haupt-Komponente ──────────────────────────────────────────
export default function Uebungsplaner() {
  const [items, setItems] = useState([]);
  const [mode, setMode] = useState("select");
  const [field, setField] = useState("half");
  const [playerColor, setPlayerColor] = useState("#2f6fde");
  const [lineColor, setLineColor] = useState("#f5f2e8");
  const [selectedId, setSelectedId] = useState(null);
  const [draft, setDraft] = useState(null);
  const [hoverHandle, setHoverHandle] = useState(null); // {id, xk, yk} nahe dem Cursor
  const [meta, setMeta] = useState(EMPTY_META);
  const [training, setTraining] = useState([]); // Array<{meta, field, items}>
  const [trainingTitel, setTrainingTitel] = useState("");
  const [activeIdx, setActiveIdx] = useState(null); // welcher Trainings-Eintrag im Editor liegt
  const trainFileRef = useRef(null);
  const svgRef = useRef(null);
  const dragRef = useRef(null);
  const fileRef = useRef(null);
  const [msg, setMsg] = useState("");

  const getPoint = (e) => {
    // getScreenCTM berücksichtigt preserveAspectRatio (Letterboxing),
    // simple rect-Skalierung tut das nicht → Versatz an den Feldrändern
    const p = new DOMPoint(e.clientX, e.clientY)
      .matrixTransform(svgRef.current.getScreenCTM().inverse());
    return { x: p.x, y: p.y };
  };

  const onCanvasDown = (e) => {
    const p = getPoint(e);
    if (mode.startsWith("el:")) {
      const type = mode.slice(3);
      const label = type === "gk" ? "T" : "";
      const item = { kind: "el", id: uid(), type, x: p.x, y: p.y, label };
      if (type === "player") item.color = playerColor;
      setItems(it => [...it, item]);
    } else if (mode.startsWith("line:")) {
      const lt = mode.slice(5);
      const col = lineColor ? { color: lineColor } : {};
      setDraft(lt === "free"
        ? { lineType: "free", ...col, points: [{ x: +p.x.toFixed(1), y: +p.y.toFixed(1) }] }
        : { lineType: lt, ...col, x1: p.x, y1: p.y, x2: p.x, y2: p.y });
    } else {
      setSelectedId(null);
    }
  };

  const onElementDown = (el) => (e) => {
    if (!mode.startsWith("line:")) {
      e.stopPropagation();
      const p = getPoint(e);
      setSelectedId(el.id);
      setMode("select");
      dragRef.current = { id: el.id, dx: p.x - el.x, dy: p.y - el.y };
    }
  };

  const onLineDown = (ln) => (e) => {
    if (!mode.startsWith("line:")) {
      e.stopPropagation();
      const p = getPoint(e);
      setSelectedId(ln.id);
      setMode("select");
      dragRef.current = { id: ln.id, line: true, sx: p.x, sy: p.y, orig: ln };
    }
  };

  const onMove = (e) => {
    if (dragRef.current?.endpoint) {
      const p = getPoint(e), { id, endpoint: [xk, yk] } = dragRef.current;
      setItems(it => it.map(i => i.id === id ? { ...i, [xk]: p.x, [yk]: p.y } : i));
    } else if (dragRef.current?.line) {
      const p = getPoint(e), { id, sx, sy, orig } = dragRef.current;
      const dx = p.x - sx, dy = p.y - sy;
      setItems(it => it.map(i => i.id !== id ? i : orig.points
        ? { ...i, points: orig.points.map(pt => ({ x: +(pt.x + dx).toFixed(1), y: +(pt.y + dy).toFixed(1) })) }
        : { ...i, x1: orig.x1 + dx, y1: orig.y1 + dy, x2: orig.x2 + dx, y2: orig.y2 + dy }));
    } else if (dragRef.current) {
      const p = getPoint(e), { id, dx, dy } = dragRef.current;
      setItems(it => it.map(i => i.id === id ? { ...i, x: p.x - dx, y: p.y - dy } : i));
    } else if (draft) {
      const p = getPoint(e);
      if (draft.points) {
        setDraft(d => {
          const last = d.points[d.points.length - 1];
          if (Math.hypot(p.x - last.x, p.y - last.y) < 3) return d;
          return { ...d, points: [...d.points, { x: +p.x.toFixed(1), y: +p.y.toFixed(1) }] };
        });
      } else {
        setDraft(d => ({ ...d, x2: p.x, y2: p.y }));
      }
    } else if (mode === "select") {
      // End-/Eckpunkte automatisch erkennen, sobald der Cursor in die Nähe kommt
      const p = getPoint(e);
      let best = null, bestD = 12; // Fangradius in viewBox-Einheiten
      for (const i of items) {
        for (const [xk, yk] of handleKeys(i)) {
          const d = Math.hypot(p.x - i[xk], p.y - i[yk]);
          if (d < bestD) { bestD = d; best = { id: i.id, xk, yk }; }
        }
      }
      setHoverHandle(h =>
        h?.id === best?.id && h?.xk === best?.xk && h?.yk === best?.yk ? h : best);
    }
  };

  const onUp = () => {
    dragRef.current = null;
    if (draft) {
      const ok = draft.points
        ? draft.points.length > 2
        : Math.hypot(draft.x2 - draft.x1, draft.y2 - draft.y1) > 15;
      if (ok) setItems(it => [...it, { kind: "line", id: uid(), ...draft }]);
      setDraft(null);
    }
  };

  const deleteSelected = () => {
    if (selectedId != null) { setItems(it => it.filter(i => i.id !== selectedId)); setSelectedId(null); }
  };
  const undo = () => setItems(it => it.slice(0, -1));
  const clearAll = () => { setItems([]); setSelectedId(null); };

  useEffect(() => {
    const h = (e) => {
      const t = e.target.tagName;
      if (t === "INPUT" || t === "TEXTAREA") return; // Tippen in Feldern löscht keine Auswahl
      if (e.key === "Delete" || e.key === "Backspace") deleteSelected();
      if (e.key === "Escape") { setSelectedId(null); setMode("select"); }
    };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  });

  const download = (blob, name) => {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  };

  const fileBase = () => slugify(meta.titel) || "uebung";

  const saveUebung = () => {
    const data = { version: 1, type: "uebung", ...snapshot() };
    download(new Blob([JSON.stringify(data)], { type: "application/json" }), fileBase() + ".json");
    setMsg(`Übung gespeichert: ${fileBase()}.json`);
  };

  const exportSVG = () => {
    const xml = new XMLSerializer().serializeToString(svgRef.current);
    download(new Blob([xml], { type: "image/svg+xml" }), fileBase() + ".svg");
  };

  const renderPNGBlob = () => svgToPng(new XMLSerializer().serializeToString(svgRef.current));

  const exportPNG = async () => {
    try { download(await renderPNGBlob(), fileBase() + ".png"); }
    catch (err) { setMsg("PNG-Export fehlgeschlagen: " + err.message); }
  };

  // ── Training ──
  const snapshot = () => structuredClone({ meta, field, items });

  const addToTraining = () => {
    setTraining(t => [...t, snapshot()]);
    setActiveIdx(training.length);
    setMsg(`Übung ${training.length + 1} zum Training hinzugefügt`);
  };

  const updateActive = () => {
    setTraining(t => t.map((u, i) => i === activeIdx ? snapshot() : u));
    setMsg(`Übung ${activeIdx + 1} im Training aktualisiert`);
  };

  const loadEntry = (i) => {
    const u = structuredClone(training[i]);
    setItems(u.items);
    setField(u.field);
    setMeta({ ...EMPTY_META, ...u.meta });
    setSelectedId(null);
    setActiveIdx(i);
    nextId = Math.max(0, ...u.items.map(x => x.id || 0)) + 1;
  };

  const removeEntry = (i) => {
    setTraining(t => t.filter((_, j) => j !== i));
    setActiveIdx(a => a === i ? null : a > i ? a - 1 : a);
  };

  const moveUp = (i) => {
    if (i === 0) return;
    setTraining(t => { const c = [...t]; [c[i - 1], c[i]] = [c[i], c[i - 1]]; return c; });
    setActiveIdx(a => a === i ? i - 1 : a === i - 1 ? i : a);
  };

  const trainingFileBase = () => slugify(trainingTitel) || "training";

  const saveTraining = () => {
    const data = { version: 1, type: "training", titel: trainingTitel, uebungen: training };
    download(new Blob([JSON.stringify(data)], { type: "application/json" }), trainingFileBase() + ".json");
    setMsg(`Training gespeichert: ${trainingFileBase()}.json (${training.length} Übungen)`);
  };

  const loadTrainingData = (data, name) => {
    setTraining(data.uebungen.map(u => ({
      meta: { ...EMPTY_META, ...(u.meta || {}) },
      field: u.field || "half",
      items: Array.isArray(u.items) ? u.items : [],
    })));
    setTrainingTitel(data.titel || "");
    setActiveIdx(null);
    setMsg(`Training geladen: ${name} (${data.uebungen.length} Übungen)`);
  };

  const onLoadTrainingFile = async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      const data = JSON.parse(await f.text());
      if (!Array.isArray(data.uebungen)) throw new Error("kein Training (uebungen fehlt)");
      loadTrainingData(data, f.name);
    } catch (err) {
      setMsg("Training laden fehlgeschlagen: " + err.message);
    }
    e.target.value = "";
  };

  const exportTrainingDOCX = async () => {
    try {
      const children = [new Paragraph({
        children: [new TextRun({ text: trainingTitel || "Training", bold: true, size: 40 })],
        spacing: { after: 300 },
      })];
      for (let i = 0; i < training.length; i++) {
        const u = training[i];
        const png = new Uint8Array(await (await svgToPng(uebungSvgString(u.field, u.items))).arrayBuffer());
        if (i > 0) children.push(new Paragraph({ children: [new PageBreak()] }));
        children.push(uebungTable(u.meta, png));
      }
      const doc = new Document({ sections: [{ children }] });
      download(await Packer.toBlob(doc), trainingFileBase() + ".docx");
      setMsg(`Word-Export erstellt: ${trainingFileBase()}.docx (${training.length} Übungen)`);
    } catch (err) {
      setMsg("Trainings-Export fehlgeschlagen: " + err.message);
    }
  };

  const exportDOCX = async () => {
    try {
      const png = new Uint8Array(await (await renderPNGBlob()).arrayBuffer());
      const doc = new Document({ sections: [{ children: [uebungTable(meta, png)] }] });
      download(await Packer.toBlob(doc), fileBase() + ".docx");
      setMsg(`Word-Export erstellt: ${fileBase()}.docx`);
    } catch (err) {
      setMsg("Word-Export fehlgeschlagen: " + err.message);
    }
  };

  const onLoadFile = async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      const text = await f.text();
      let data;
      if (text.trim().startsWith("<")) {
        const doc = new DOMParser().parseFromString(text, "image/svg+xml");
        const meta = doc.getElementById("uebung-data");
        if (!meta) throw new Error("keine Übungsdaten im SVG gefunden");
        data = JSON.parse(meta.textContent);
      } else {
        data = JSON.parse(text);
      }
      if (Array.isArray(data.uebungen)) { // Trainings-Datei erkannt
        loadTrainingData(data, f.name);
        e.target.value = "";
        return;
      }
      if (!Array.isArray(data.items)) throw new Error("ungültiges Format");
      setItems(data.items);
      setField(data.field || "half");
      setMeta({ ...EMPTY_META, ...(data.meta || {}) });
      setSelectedId(null);
      nextId = Math.max(0, ...data.items.map(i => i.id || 0)) + 1;
      setMsg(`Geladen: ${f.name} (${data.items.length} Objekte)`);
    } catch (err) {
      setMsg(`Laden fehlgeschlagen: ${err.message}`);
    }
    e.target.value = "";
  };

  // ── UI-Stile ──
  const btn = (active) => ({
    display: "block", width: "100%", textAlign: "left", padding: "7px 10px",
    marginBottom: 4, borderRadius: 6, cursor: "pointer",
    border: active ? "1px solid #ffd447" : "1px solid #3a423d",
    background: active ? "#ffd44718" : "#242b27",
    color: active ? "#ffd447" : "#d8d5ca",
    fontFamily: "ui-monospace, SFMono-Regular, monospace", fontSize: 12.5,
  });
  const groupLabel = {
    fontSize: 10.5, letterSpacing: 1.5, textTransform: "uppercase",
    color: "#8a948c", margin: "14px 0 6px", fontFamily: "ui-monospace, monospace",
  };
  const fieldLabel = {
    display: "block", fontSize: 10.5, color: "#8a948c", marginBottom: 2,
    fontFamily: "ui-monospace, monospace",
  };
  const inp = {
    width: "100%", boxSizing: "border-box", padding: "5px 8px", marginBottom: 8,
    borderRadius: 6, border: "1px solid #3a423d", background: "#242b27",
    color: "#d8d5ca", fontFamily: "ui-monospace, monospace", fontSize: 12,
    outline: "none", resize: "vertical",
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", background: "#1c2320", color: "#d8d5ca", fontFamily: "system-ui, sans-serif" }}>
      {/* Kopfzeile */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 16px", borderBottom: "1px solid #313a34" }}>
        <div style={{ fontFamily: "ui-monospace, monospace", fontWeight: 700, fontSize: 15, letterSpacing: 0.5 }}>
          ÜBUNGSPLANER<span style={{ color: "#ffd447" }}>_</span>
        </div>
        <div style={{ flex: 1 }} />
        <input ref={fileRef} type="file" accept=".svg,.json" style={{ display: "none" }} onChange={onLoadFile} />
        <button style={{ ...btn(false), width: "auto", marginBottom: 0 }} onClick={() => fileRef.current.click()}>Übung laden</button>
        <button style={{ ...btn(false), width: "auto", marginBottom: 0 }} onClick={saveUebung}>Übung speichern</button>
        <button style={{ ...btn(false), width: "auto", marginBottom: 0 }} onClick={exportSVG}>SVG exportieren</button>
        <button style={{ ...btn(false), width: "auto", marginBottom: 0 }} onClick={exportPNG}>PNG exportieren</button>
        <button style={{ ...btn(false), width: "auto", marginBottom: 0 }} onClick={exportDOCX}>Word exportieren</button>
      </div>

      <div style={{ display: "flex", flex: 1, minHeight: 0 }}>
        {/* Werkzeugleiste */}
        <div style={{ width: 168, padding: "8px 12px", overflowY: "auto", borderRight: "1px solid #313a34", flexShrink: 0 }}>
          <div style={groupLabel}>Modus</div>
          <button style={btn(mode === "select")} onClick={() => setMode("select")}>Auswählen / Verschieben</button>

          <div style={groupLabel}>Elemente</div>
          <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 4 }}>
            <button style={{ ...btn(mode === "el:player"), flex: 1, minWidth: 0, marginBottom: 0 }} onClick={() => setMode("el:player")}>Spieler</button>
            <ColorPicker color={playerColor} title="Spielerfarbe wählen"
              onPick={(c) => { setPlayerColor(c); setMode("el:player"); }} />
          </div>
          {EL_DEFS.map(d => (
            <button key={d.key} style={btn(mode === "el:" + d.key)} onClick={() => setMode("el:" + d.key)}>{d.label}</button>
          ))}

          <div style={{ ...groupLabel, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <span>Wege & Formen</span>
            <ColorPicker color={lineColor} title="Linienfarbe wählen"
              onPick={setLineColor} />
          </div>
          {LINE_DEFS.map(d => (
            <button key={d.key} style={btn(mode === "line:" + d.key)} onClick={() => setMode("line:" + d.key)} title={d.hint}>{d.label}</button>
          ))}

          <div style={groupLabel}>Feld</div>
          {[["half", "Halbfeld"], ["full", "Ganzes Feld"], ["penalty", "Strafraum"], ["blank", "Freie Fläche"]].map(([k, l]) => (
            <button key={k} style={btn(field === k)} onClick={() => setField(k)}>{l}</button>
          ))}

          <div style={groupLabel}>Aktionen</div>
          <button style={btn(false)} onClick={deleteSelected} disabled={selectedId == null}>Auswahl löschen</button>
          <button style={btn(false)} onClick={undo}>Rückgängig</button>
          <button style={btn(false)} onClick={clearAll}>Alles löschen</button>
        </div>

        {/* Zeichenfläche */}
        <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", padding: 14, minWidth: 0 }}>
          <svg
            ref={svgRef}
            viewBox={`0 0 ${VW} ${VH}`}
            style={{ width: "100%", maxHeight: "100%", borderRadius: 8, touchAction: "none", cursor: mode === "select" ? "default" : "crosshair", boxShadow: "0 4px 24px #0006" }}
            xmlns="http://www.w3.org/2000/svg"
            onPointerDown={onCanvasDown}
            onPointerMove={onMove}
            onPointerUp={onUp}
            onPointerLeave={onUp}
          >
            <ArrowDefs />
            <metadata id="uebung-data">{JSON.stringify({ version: 1, field, meta, items })}</metadata>
            <PitchStripes />
            <PitchLines field={field} />
            {items.filter(i => i.kind === "line").map(ln => (
              <Line key={ln.id} ln={ln} selected={ln.id === selectedId} onDown={onLineDown(ln)} />
            ))}
            {draft && <Line ln={draft} />}
            {items.filter(i => i.kind === "el").map(el => (
              <Element key={el.id} el={el} selected={el.id === selectedId} onDown={onElementDown(el)} />
            ))}
            {/* Endpunkt-/Eck-Handles der selektierten Linie (Freihand hat keine) */}
            {(() => {
              const sel = items.find(i => i.id === selectedId && i.kind === "line");
              if (!sel) return null;
              return handleKeys(sel).map(([xk, yk]) => (
                <circle key={xk + yk} cx={sel[xk]} cy={sel[yk]} r={6}
                  fill="#ffd447" stroke="#1c2320" strokeWidth={1.5}
                  style={{ cursor: "grab" }}
                  onPointerDown={(e) => {
                    e.stopPropagation();
                    dragRef.current = { id: sel.id, endpoint: [xk, yk] };
                  }} />
              ));
            })()}
            {/* Hover-Handle: erscheint beim Annähern an einen End-/Eckpunkt,
                grosszügige unsichtbare Grabfläche (r 14) */}
            {mode === "select" && hoverHandle && (() => {
              const ln = items.find(i => i.id === hoverHandle.id);
              if (!ln) return null;
              const hx = ln[hoverHandle.xk], hy = ln[hoverHandle.yk];
              const grab = (e) => {
                e.stopPropagation();
                setSelectedId(ln.id);
                dragRef.current = { id: ln.id, endpoint: [hoverHandle.xk, hoverHandle.yk] };
              };
              return (
                <g>
                  <circle cx={hx} cy={hy} r={7} fill="#ffd447" stroke="#1c2320" strokeWidth={1.5} />
                  <circle cx={hx} cy={hy} r={14} fill="transparent" style={{ cursor: "grab" }} onPointerDown={grab} />
                </g>
              );
            })()}
          </svg>
        </div>

        {/* Übungsdaten */}
        <div style={{ width: 210, padding: "8px 12px", overflowY: "auto", borderLeft: "1px solid #313a34", flexShrink: 0 }}>
          <div style={groupLabel}>Übungsdaten</div>
          {META_FIELDS.map(([key, label, multi]) => (
            <div key={key}>
              <label htmlFor={"meta-" + key} style={fieldLabel}>{label}</label>
              {multi ? (
                <textarea id={"meta-" + key} rows={4} style={inp} value={meta[key]}
                  onChange={e => setMeta(m => ({ ...m, [key]: e.target.value }))} />
              ) : (
                <input id={"meta-" + key} type="text" style={inp} value={meta[key]}
                  onChange={e => setMeta(m => ({ ...m, [key]: e.target.value }))} />
              )}
            </div>
          ))}

          <div style={groupLabel}>Training</div>
          <label htmlFor="training-titel" style={fieldLabel}>Trainings-Titel</label>
          <input id="training-titel" type="text" style={inp} value={trainingTitel}
            onChange={e => setTrainingTitel(e.target.value)} />
          <button style={btn(false)} onClick={addToTraining}>+ Übung hinzufügen</button>
          {activeIdx != null && activeIdx < training.length && (
            <button style={btn(false)} onClick={updateActive}>Übung {activeIdx + 1} aktualisieren</button>
          )}
          {training.map((u, i) => (
            <div key={i} style={{ display: "flex", gap: 4 }}>
              <button
                style={{ ...btn(i === activeIdx), flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}
                onClick={() => loadEntry(i)} title={u.meta.titel || "Übung"}>
                {i + 1} · {u.meta.titel || "Übung"}
              </button>
              <button style={{ ...btn(false), width: 26, padding: "7px 0", textAlign: "center", flexShrink: 0 }}
                onClick={() => moveUp(i)} disabled={i === 0}>↑</button>
              <button style={{ ...btn(false), width: 26, padding: "7px 0", textAlign: "center", flexShrink: 0 }}
                onClick={() => removeEntry(i)} title="Aus Training entfernen">✕</button>
            </div>
          ))}
          {training.length > 0 && (
            <>
              <button style={btn(false)} onClick={saveTraining}>Training speichern</button>
              <button style={btn(false)} onClick={exportTrainingDOCX}>Training → Word</button>
            </>
          )}
          <button style={btn(false)} onClick={() => trainFileRef.current.click()}>Training laden</button>
          <input ref={trainFileRef} type="file" accept=".json" style={{ display: "none" }} onChange={onLoadTrainingFile} />
        </div>
      </div>

      {/* Fusszeile */}
      <div style={{ padding: "6px 16px", fontSize: 11.5, color: msg.startsWith("Laden fehl") ? "#ff8c7a" : "#8a948c", borderTop: "1px solid #313a34", fontFamily: "ui-monospace, monospace" }}>
        {msg || "Element wählen → aufs Feld tippen · Weg wählen → ziehen · Auswählen-Modus: Elemente & Wege verschieben, Entf löscht"}
      </div>
    </div>
  );
}
