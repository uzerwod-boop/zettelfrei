import { useEffect, useMemo, useRef, useState } from "react";
import { S, canRedo, canUndo, closeFile, commit, emit, issues, mark, openExample, openFile, persist, redo, saveExcel, selectSheet, setPane, sheet, toast, undo, unsaved, useStore, type Pane } from "../store";
import { confirmBox, openMenu, promptBox } from "./Overlays";
import { HIER, IS_APP, SHOW_EXAMPLES, de, fmtVal, openExternal, rawOf } from "../util";
import { GUIDE_URL } from "../config";
import { openFeedback } from "./Feedback";
import { Icon, Mark } from "./Icons";

/* ---------- Leiste links ---------- */
const NAV: { id: Pane; label: string; icon: string }[] = [
  { id: "cmd", label: "Assistent", icon: "chat" }, { id: "pdf", label: "Belege", icon: "file" }, { id: "docs", label: "Dokumente", icon: "doc" },
  { id: "check", label: "Prüfen", icon: "shield" }, { id: "hist", label: "Verlauf", icon: "history" }];
export function Rail() {
  const s = useStore(); const n = useMemo(() => issues().length, [s.rev]);
  return (
    <nav className="rail" role="tablist" aria-label="Bereiche" aria-orientation="vertical">
      <div className="rail-logo" title="Zettelfrei"><Mark dark /></div>
      {NAV.map(x => (
        <button key={x.id} role="tab" id={"t-" + x.id} aria-selected={s.pane === x.id} onClick={() => setPane(x.id)}>
          <Icon name={x.icon} /><span>{x.label}</span>
          {x.id === "check" && n > 0 && <span className="n">{n}</span>}
        </button>))}
      <div className="rail-foot">
        <button type="button" id="t-help" onClick={() => openExternal(GUIDE_URL)} title="Erste Schritte (Anleitung)"><Icon name="help" /><span>Hilfe</span></button>
        <button type="button" id="t-feedback" onClick={openFeedback} title="Feedback senden"><Icon name="feedback" /><span>Feedback</span></button>
      </div>
    </nav>);
}

/* ---------- Kopf ---------- */
export function TopBar() {
  const s = useStore(); const input = useRef<HTMLInputElement>(null);
  const last = s.history[s.history.length - 1];
  const examples = Object.keys(window.ZF_EXAMPLES || {});
  return (
    <header className="top">
      <div className="fileinfo">
        <h1 className="vh">Zettelfrei</h1>
        <span className="fname" title="Geöffnete Datei">{s.model ? s.model.name : "Keine Datei geöffnet"}</span>
        {s.model && <button type="button" className="iconbtn closefile" title="Datei schließen" aria-label="Datei schließen" onClick={askClose}><Icon name="close" size={16} /></button>}
        {last && <span className="saved">· {unsaved() ? "Änderungen nur hier gesichert" : "gespeichert"} {new Date(last.t).toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" })}</span>}
      </div>
      <div className="filebar">
        {s.model && <div className="undo" role="group" aria-label="Rückgängig und Wiederholen">
          <button type="button" className="iconbtn" disabled={!canUndo()} onClick={undo} title={`Rückgängig: ${canUndo() ? last.label : "–"} (${MAC ? "⌘" : "Strg+"}Z)`} aria-label="Rückgängig"><Icon name="undo" size={18} /></button>
          <button type="button" className="iconbtn" disabled={!canRedo()} onClick={redo} title={`Wiederholen (${MAC ? "⌘⇧Z" : "Strg+Y"})`} aria-label="Wiederholen"><Icon name="redo" size={18} /></button>
        </div>}
        {SHOW_EXAMPLES && <select aria-label="Beispiel öffnen" value="" onChange={e => { if (e.target.value) openExample(e.target.value); }}>
          <option value="">Beispiel öffnen…</option>
          {examples.map(n => <option key={n} value={n}>{n.replace(/_/g, " ").replace(/\.xlsx$/, "")}</option>)}
        </select>}
        <button className="btn quiet" onClick={() => input.current?.click()}>Excel öffnen</button>
        <input ref={input} type="file" hidden accept=".xlsx,.xlsm,.xls,.csv,.ods" onChange={e => { const f = e.target.files?.[0]; e.target.value = ""; if (f) openFile(f); }} />
        <button className="btn" onClick={saveExcel} disabled={!s.model}>Als Excel speichern</button>
      </div>
    </header>);
}

async function askClose() {
  if (unsaved() && !(await confirmBox("Datei schließen?", "Ihre Änderungen sind noch nicht als Excel-Datei gespeichert. Ohne Speichern gehen sie verloren.", "Schließen ohne Speichern", true))) return;
  closeFile();
}
/* ---------- Kennzahlen ---------- */
export function Kpis() {
  const s = useStore(); const sh = sheet();
  const data = useMemo(() => {
    if (!sh || sh.hr < 0) return null;
    const calc = new ZF.Calc(s.model), comp = ZF.computedCols(sh), numc = sh.headers.map((_: any, k: number) => comp[k] || ZF.numericCol(sh, k));
    const pick = (re: RegExp) => sh.headers.findIndex((h: string, k: number) => re.test(h) && !/preis|satz|%/i.test(h) && numc[k]);
    let mk = pick(/wert|lohn|betrag|brutto|summe|gesamt/i); if (mk < 0) mk = pick(/netto|€|eur/i);
    let third: [string, string];
    if (mk >= 0) { let t = 0; ZF.eachData(sh, (_: any, i: number) => { const v = calc.value(sh, ZF.dataRowNo(sh, i), sh.c0 + mk); if (typeof v === "number") t += v; });
      third = ["Summe " + sh.headers[mk].replace(/\s*€\s*$/, ""), de(t, 2) + " €"]; }
    else third = ["Berechnete Spalten", String(comp.filter(Boolean).length)];
    const n = issues().filter((x: any) => x.blatt === sh.name).length;
    return { rows: ZF.dataCount(sh), cols: sh.ncol, third, n };
  }, [s.rev, s.cur]);
  if (!s.model) return null;
  if (!data) return <div className="kpis"><p className="note">In diesem Blatt wurde keine Tabelle mit Überschriften erkannt. Es wird beim Speichern unverändert übernommen.</p></div>;
  return (
    <div className="kpis" aria-live="polite">
      <div className="kpi"><span>Zeilen</span><b>{data.rows}</b></div>
      <div className="kpi"><span>Spalten</span><b>{data.cols}</b></div>
      <div className="kpi"><span>{data.third[0]}</span><b>{data.third[1]}</b></div>
      {data.n ? <button type="button" className="kpi warn" onClick={() => setPane("check")}><span>Zum Prüfen</span><b>{data.n} Hinweis{data.n === 1 ? "" : "e"}</b></button>
        : <div className="kpi ok"><span>Zum Prüfen</span><b>alles ok</b></div>}
    </div>);
}

/* ---------- Tabelle wie in Excel: auswählen, Bereiche, tippen, kopieren/einfügen, Rechtsklick ---------- */
type Sel = { r: number; k: number } | null;
const KIND: Record<string, string> = { leer: "Leerzeile", summe: "Zwischensumme", abschnitt: "Abschnitt", notiz: "Notiz" };
const MAC = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);
const KEY = (k: string) => (MAC ? "⌘" : "Strg+") + k;
// Text aus Excel/Numbers/Google Sheets: Tabs und Zeilen, Felder evtl. in Anführungszeichen
function parseTSV(t: string): string[][] {
  const out: string[][] = []; let row: string[] = [], f = "", q = false;
  for (let i = 0; i < t.length; i++) { const c = t[i];
    if (q) { if (c === '"') { if (t[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; continue; }
    if (c === '"' && f === "") { q = true; continue; }
    if (c === "\t") { row.push(f); f = ""; continue; }
    if (c === "\n" || c === "\r") { if (c === "\r" && t[i + 1] === "\n") i++; row.push(f); out.push(row); row = []; f = ""; continue; }
    f += c; }
  if (f !== "" || row.length) { row.push(f); out.push(row); }
  return out;
}
let dragging = false; if (typeof window !== "undefined") window.addEventListener("mouseup", () => { dragging = false; });

export function Sheet() {
  const s = useStore(); const sh = sheet();
  const [sel, setSel] = useState<Sel>(null), [anchor, setAnchor] = useState<Sel>(null), [edit, setEdit] = useState<string | null>(null), [showAll, setShowAll] = useState(false);
  const editDone = useRef(false);
  const wrap = useRef<HTMLDivElement>(null), inp = useRef<HTMLInputElement>(null), drag = useRef(0);
  const [over, setOver] = useState(false);
  useEffect(() => { setSel(null); setAnchor(null); setEdit(null); setShowAll(false); }, [s.cur, s.model]);
  // Sprung von außen (Hinweis, Chat): Zeile zeigen und markieren
  useEffect(() => { const f = s.focus; if (!f || !sh || f.sheet !== sh.name) return;
    const tr = wrap.current?.querySelector(`tr[data-row="${f.r}"]`) as HTMLElement | null; tr?.scrollIntoView({ block: "center", behavior: "smooth" });
    if (f.k != null) { setSel({ r: f.r, k: f.k }); setAnchor(null); setTimeout(() => wrap.current?.focus(), 30); } }, [s.focus?.n]);
  useEffect(() => { if (edit != null) inp.current?.focus(); }, [edit != null]);
  // Strg/⌘+Z und Strg/⌘+Y überall, außer beim Tippen in Feldern
  useEffect(() => { const h = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement; if (!(e.metaKey || e.ctrlKey) || /^(INPUT|TEXTAREA|SELECT)$/.test(t?.tagName) || t?.isContentEditable || document.querySelector("dialog[open]")) return;
      const k = e.key.toLowerCase(); if (k === "z" && !e.shiftKey) { e.preventDefault(); undo(); } else if (k === "y" || (k === "z" && e.shiftKey)) { e.preventDefault(); redo(); } };
    window.addEventListener("keydown", h); return () => window.removeEventListener("keydown", h); }, []);

  const view = useMemo(() => {
    if (!sh) return null;
    const calc = new ZF.Calc(s.model), flag: Record<string, string> = {};
    issues().filter((x: any) => x.blatt === sh.name).forEach((x: any) => { const key = x.zeile + ":" + (x.k != null ? x.k : ZF.findCol(sh, x.spalte)); flag[key] = (flag[key] ? flag[key] + "; " : "") + x.problem; });
    const comp = sh.hr >= 0 ? ZF.computedCols(sh) : [], numc = sh.hr >= 0 ? sh.headers.map((_: any, k: number) => comp[k] || ZF.numericCol(sh, k)) : [];
    return { calc, flag, numc, comp };
  }, [s.rev, s.cur]);
  // Datei auf die Tabelle ziehen: Excel/CSV öffnen, alles andere als Beleg lesen
  const dz = {
    onDragOver: (e: React.DragEvent) => { if ([...e.dataTransfer.items].some(i => i.kind === "file")) { e.preventDefault(); drag.current++; setOver(true); } },
    onDragLeave: () => setOver(false),
    onDrop: (e: React.DragEvent) => { const f = e.dataTransfer.files[0]; setOver(false); if (!f) return; e.preventDefault();
      if (/\.(xlsx|xlsm|xls|csv|ods)$/i.test(f.name) || !S.model) openFile(f); else { setPane("pdf"); window.dispatchEvent(new CustomEvent("zf-beleg", { detail: f })); } },
  };
  if (!s.model || !sh || !view) return <section className="sheetwrap" aria-label="Tabelle" {...dz}><div className={"sheet empty-state" + (over ? " over" : "")}><Empty /></div></section>;
  const { calc, flag, numc, comp } = view;
  if (edit == null) editDone.current = false;
  const C = sh.hr < 0 ? Math.max(1, ...sh.top.map((r: any[]) => (r || []).length)) : sh.c0 + sh.ncol;
  const firstRow = sh.hr + 2, lastData = sh.hr + 1 + sh.rows.length;
  const editable = (r: number) => { if (sh.hr < 0 || r <= sh.hr + 1) return false; const i = r - sh.hr - 2; return !(i < sh.rows.length && sh.kind && sh.kind[i] === "leer"); };
  const lastRow = sh.hr < 0 ? sh.top.length : sh.hr + 1 + sh.rows.length + sh.foot.length;
  const selCell = sel ? ZF.cellAt(sh, sel.r, sh.c0 + sel.k) : null;
  const a: Sel = anchor || sel; const rng = sel && a ? { r1: Math.min(a.r, sel.r), r2: Math.max(a.r, sel.r), k1: Math.min(a.k, sel.k), k2: Math.max(a.k, sel.k) } : null;
  const multi = !!rng && (rng.r1 !== rng.r2 || rng.k1 !== rng.k2);
  const inRng = (r: number, k: number) => !!rng && r >= rng.r1 && r <= rng.r2 && k >= rng.k1 && k <= rng.k2;
  const dataRowsIn = (r1: number, r2: number) => { const out: number[] = []; for (let r = r1; r <= r2; r++) if (r >= firstRow && r <= lastData) out.push(r); return out; };
  const done = (label: string, rows?: number[]) => { S.fresh = {}; S.hits = {}; if (rows?.length) mark("fresh", sh.name, rows); commit(label); if (rows?.length) fadeFresh(); };
  const cellText = (r: number, k: number) => { const v = calc.value(sh, r, sh.c0 + k);
    if (v == null) return ""; if (typeof v === "number") { if (sh.timeCol && sh.timeCol[k]) return ZF.fmtTime(v); if (sh.dateCol[k] && v > 20000) return ZF.serialToDate(v); return String(Math.round(v * 1e6) / 1e6).replace(".", ","); }
    return String(v); };

  const commitEdit = (move?: [number, number]) => {
    if (!sel || edit == null || editDone.current) return; editDone.current = true;
    const before = rawOf(selCell);
    if (edit.trim() !== before) {
      ZF.setCell(sh, sel.r, sel.k, edit.trim(), { formula: true });
      if (sh.timeCol && sh.timeCol[sel.k]) { const tc = ZF.cellAt(sh, sel.r, sh.c0 + sel.k); if (tc && typeof tc.v === "number") tc.tm = 1; }
      commit(`${ZF.colLetter(sh.c0 + sel.k)}${sel.r} geändert: „${edit.length > 30 ? edit.slice(0, 30) + "…" : edit.trim()}“`);
    }
    setEdit(null); if (move) moveSel(move[0], move[1], false); wrap.current?.focus();
  };
  const moveSel = (dr: number, dk: number, extend: boolean) => {
    if (!extend) setAnchor(null); else if (!anchor) setAnchor(sel);
    setSel(p => { if (!p) return p; let r = p.r + dr; const k = Math.max(0, Math.min(sh.ncol - 1, p.k + dk));
      while (r > sh.hr + 1 && r <= lastRow && !editable(r) && dr) r += dr;
      if (r <= sh.hr + 1 || r > lastRow) r = p.r; return { r, k }; }); };
  const clearRange = () => { if (!rng) return; let n = 0;
    for (let r = rng.r1; r <= rng.r2; r++) for (let k = rng.k1; k <= rng.k2; k++) { if (!editable(r) || comp[k]) continue; const c = ZF.cellAt(sh, r, sh.c0 + k); if (rawOf(c)) { ZF.setCell(sh, r, k, "", { formula: true }); n++; } }
    if (n) done(multi ? `${n} Zellen geleert` : `${ZF.colLetter(sh.c0 + rng.k1)}${rng.r1} geleert`); };
  const copyText = () => { if (!rng) return ""; const lines: string[] = [];
    for (let r = rng.r1; r <= rng.r2; r++) { const cells: string[] = []; for (let k = rng.k1; k <= rng.k2; k++) { const t = cellText(r, k); cells.push(/[\t\n"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t); } lines.push(cells.join("\t")); }
    return lines.join("\n"); };
  const pasteText = (text: string) => {
    if (!sel || !text) return; let data = parseTSV(text.replace(/\n$/, "")); if (!data.length) return;
    const r0 = rng ? rng.r1 : sel.r, k0 = rng ? rng.k1 : sel.k;
    // ein einzelner Wert in einen markierten Bereich: überall eintragen (wie Excel)
    if (data.length === 1 && data[0].length === 1 && multi && rng) data = Array.from({ length: rng.r2 - rng.r1 + 1 }, () => Array.from({ length: rng.k2 - rng.k1 + 1 }, () => data[0][0]));
    let r = r0, n = 0, skipComp = 0, skipCols = 0; const touched: number[] = [];
    for (const line of data) {
      while (r <= lastData && !editable(r)) r++;
      if (r > sh.hr + 1 + sh.rows.length) r = ZF.appendRow(sh, [], S.model);
      line.forEach((val, j) => { const k = k0 + j; if (k >= sh.ncol) { skipCols++; return; } if (comp[k]) { skipComp++; return; }
        ZF.setCell(sh, r, k, val.trim(), { formula: true }); n++; });
      touched.push(r); r++;
    }
    if (n) { done(`${n} Zelle${n === 1 ? "" : "n"} eingefügt`, touched); setAnchor({ r: touched[0], k: k0 }); setSel({ r: touched[touched.length - 1], k: Math.min(sh.ncol - 1, k0 + Math.max(...data.map(l => l.length)) - 1) }); }
    if (skipComp || skipCols) toast(`${n} Zellen eingefügt${skipComp ? ` · ${skipComp} in berechneten Spalten übersprungen` : ""}${skipCols ? ` · ${skipCols} rechts außerhalb der Tabelle` : ""}`);
  };
  const onKey = (e: React.KeyboardEvent) => {
    if (!sel || edit != null) return;
    const mod = e.metaKey || e.ctrlKey;
    if (mod && e.key.toLowerCase() === "a") { e.preventDefault(); setAnchor({ r: firstRow, k: 0 }); setSel({ r: lastData, k: sh.ncol - 1 }); return; }
    if (mod) return; // kopieren/einfügen/rückgängig über eigene Ereignisse
    const m: Record<string, [number, number]> = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Enter: [1, 0], Tab: [0, e.shiftKey ? -1 : 1] };
    if (m[e.key]) { e.preventDefault(); moveSel(...m[e.key], e.shiftKey && e.key.startsWith("Arrow")); return; }
    if (e.key === "Escape") { setAnchor(null); return; }
    if (e.key === "F2") { e.preventDefault(); setAnchor(null); setEdit(rawOf(selCell)); return; }
    if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); clearRange(); return; }
    if (e.key.length === 1 && !e.altKey) { e.preventDefault(); setAnchor(null); setEdit(e.key); }
  };
  const pick = (r: number, k: number, e: React.MouseEvent) => {
    if (edit != null) commitEdit();
    if (e.button === 2 && inRng(r, k)) return;            // Rechtsklick in den markierten Bereich: Auswahl behalten
    if (e.shiftKey && sel) { if (!anchor) setAnchor(sel); setSel({ r, k }); return; }
    setAnchor(null); setSel({ r, k }); dragging = e.button === 0;
  };
  // ----- Rechtsklick-Menüs -----
  const insertRows = (at: number, n: number) => { const rows: number[] = []; for (let i = 0; i < n; i++) rows.push(ZF.insertRowAt(sh, at + i, null, S.model));
    done(n === 1 ? `Zeile ${at} eingefügt` : `${n} Zeilen ab ${at} eingefügt`, rows); setAnchor(null); setSel({ r: rows[0], k: Math.max(0, comp.findIndex((c: boolean) => !c)) }); setTimeout(() => wrap.current?.focus(), 30); };
  const deleteRows = (rows: number[]) => { if (!rows.length) return; [...rows].sort((x, y) => y - x).forEach(r => ZF.deleteRow(sh, r, S.model));
    setAnchor(null); setSel(null); done(rows.length === 1 ? `Zeile ${rows[0]} gelöscht` : `${rows.length} Zeilen gelöscht`); toast(`${rows.length === 1 ? "Zeile" : rows.length + " Zeilen"} gelöscht · ${KEY("Z")} holt ${rows.length === 1 ? "sie" : "sie"} zurück`); };
  const rowMenu = (e: React.MouseEvent, r: number) => {
    const rows = rng && r >= rng.r1 && r <= rng.r2 ? dataRowsIn(rng.r1, rng.r2) : [r], n = Math.max(1, rows.length);
    openMenu(e, [
      { label: n === 1 ? "Zeile darüber einfügen" : `${n} Zeilen darüber einfügen`, onClick: () => insertRows(rows[0], n) },
      { label: n === 1 ? "Zeile darunter einfügen" : `${n} Zeilen darunter einfügen`, onClick: () => insertRows(rows[rows.length - 1] + 1, n) },
      "sep",
      { label: "Kopieren", hint: KEY("C"), onClick: () => navigator.clipboard?.writeText(copyText()).then(() => toast("Kopiert"), () => toast(`Bitte ${KEY("C")} drücken`)) },
      { label: "Einfügen", hint: KEY("V"), onClick: () => navigator.clipboard?.readText ? navigator.clipboard.readText().then(pasteText, () => toast(`Bitte ${KEY("V")} drücken`)) : toast(`Bitte ${KEY("V")} drücken`) },
      { label: "Inhalt löschen", hint: "Entf", onClick: clearRange },
      "sep",
      { label: n === 1 ? `Zeile ${rows[0]} löschen` : `${n} Zeilen löschen`, danger: true, onClick: () => deleteRows(rows) },
    ]); };
  const colMenu = (e: React.MouseEvent, k: number) => openMenu(e, [
    { label: "Aufsteigend sortieren (A → Z, 0 → 9)", onClick: () => sortBy(k, false) },
    { label: "Absteigend sortieren (Z → A, 9 → 0)", onClick: () => sortBy(k, true) },
    "sep",
    { label: "Spalte umbenennen …", onClick: () => renameCol(k) },
    { label: "Neue Spalte am Ende …", onClick: addCol },
    { label: `Spalte „${sh.headers[k]}“ löschen`, danger: true, disabled: k !== sh.ncol - 1, hint: k !== sh.ncol - 1 ? "nur letzte" : undefined, onClick: () => delCol() },
  ]);
  const sortBy = (k: number, desc: boolean) => { const r = ZF.sortRows(S.model, sh, k, desc); if (r.error) { toast(r.error); return; } setSel(null); setAnchor(null); done(`Nach „${sh.headers[k]}“ ${desc ? "absteigend" : "aufsteigend"} sortiert`); };
  const renameCol = async (k: number) => { const n = await promptBox("Spalte umbenennen", sh.headers[k]); if (n == null) return; const r = ZF.renameColumn(sh, k, n); if (r.error) { toast(r.error); return; } done(`Spalte umbenannt: „${n.trim()}“`); };
  const addCol = async () => { const n = await promptBox("Neue Spalte", "", "Anlegen", "Name der Spalte, z. B. „Bemerkung“ oder „Rechnung am“."); if (!n || !n.trim()) return;
    if (sh.headers.some((h: string) => h.toLowerCase() === n.trim().toLowerCase())) { toast(`Eine Spalte „${n.trim()}“ gibt es schon.`); return; }
    const k = ZF.addColumn(sh, n.trim()); done(`Neue Spalte „${n.trim()}“`); setAnchor(null); setSel({ r: firstRow, k }); };
  const delCol = async () => { const k = sh.ncol - 1; if (!(await confirmBox(`Spalte „${sh.headers[k]}“ löschen?`, "Die Inhalte dieser Spalte werden entfernt. Mit Rückgängig holen Sie sie zurück.", "Löschen", true))) return;
    const r = ZF.deleteLastColumn(S.model, sh); if (r.error) { toast(r.error); return; } setSel(null); setAnchor(null); done(`Spalte „${sh.headers[k] || ""}“ gelöscht`.replace("„“ ", "")); };
  const selectCol = (k: number, e: React.MouseEvent) => { if (sh.hr < 0 || k < 0 || k >= sh.ncol) return; if (e.shiftKey && sel) { setAnchor(a || { r: firstRow, k: sel.k }); setSel({ r: lastData, k }); } else { setAnchor({ r: firstRow, k }); setSel({ r: lastData, k }); } wrap.current?.focus(); };
  const selectRow = (r: number, e: React.MouseEvent) => { if (!editable(r)) return; if (edit != null) commitEdit();
    if (e.shiftKey && sel) { setAnchor(a ? { r: a.r, k: 0 } : { r: sel.r, k: 0 }); setSel({ r, k: sh.ncol - 1 }); } else { setAnchor({ r, k: 0 }); setSel({ r, k: sh.ncol - 1 }); } dragging = false; wrap.current?.focus(); };

  const cellTd = (row: any[], r: number, k: number, canEdit: boolean) => {
    const cell = row[k], v = calc.value(sh, r, sh.c0 + k), fl = flag[r + ":" + k], isSel = sel && sel.r === r && sel.k === k, inR = multi && inRng(r, k);
    const cls = [(typeof v === "number" || (numc[k] && (v == null || v === ""))) ? "num" : "", cell?.f ? "fxc" : "", fl ? "flag" : "", isSel ? "sel" : "", inR ? "inrange" : ""].join(" ").trim();
    return (
      <td key={k} className={cls} title={fl || undefined} data-r={r} data-k={k}
        onMouseDown={canEdit ? e => pick(r, k, e) : undefined}
        onMouseEnter={canEdit ? e => { if (dragging && (e.buttons & 1) && sel) { if (!anchor) setAnchor(sel); setSel({ r, k }); } } : undefined}
        onContextMenu={canEdit ? e => rowMenu(e, r) : undefined}
        onDoubleClick={canEdit ? () => { setAnchor(null); setSel({ r, k }); setEdit(rawOf(cell)); } : undefined}>
        {isSel && edit != null
          ? <input ref={inp} className="cellin" value={edit} aria-label={`Zelle ${ZF.colLetter(sh.c0 + k)}${r}`} onChange={e => setEdit(e.target.value)}
              onKeyDown={e => { if (e.key === "Enter") { e.preventDefault(); commitEdit([1, 0]); } else if (e.key === "Tab") { e.preventDefault(); commitEdit([0, e.shiftKey ? -1 : 1]); } else if (e.key === "Escape") { setEdit(null); wrap.current?.focus(); } }}
              onBlur={() => commitEdit()} />
          : fmtVal(sh, k, v, cell)}
      </td>);
  };
  const fresh = s.fresh[sh.name] || {}, hits = s.hits[sh.name] || {};
  const limit = showAll ? sh.rows.length : Math.min(sh.rows.length, 400);
  const pre = (n: number) => Array.from({ length: n }, (_, i) => <td key={"p" + i} />);
  // Summe des markierten Bereichs (wie die Statusleiste in Excel)
  let rsum = 0, rcnt = 0, rnum = 0;
  if (multi && rng) for (let r = rng.r1; r <= rng.r2; r++) for (let k = rng.k1; k <= rng.k2; k++) { const v = calc.value(sh, r, sh.c0 + k); if (v == null || v === "") continue; rcnt++; if (typeof v === "number" && !(sh.timeCol && sh.timeCol[k]) && !(sh.dateCol[k] && v > 20000)) { rsum += v; rnum++; } }
  const tabMenu = (e: React.MouseEvent, i: number) => { const x = s.model.sheets[i];
    openMenu(e, [{ label: "Umbenennen …", onClick: () => renameSheetUI(i) }, { label: "Neues Blatt", onClick: addSheet }, "sep",
      { label: `Blatt „${x.name}“ löschen …`, danger: true, disabled: s.model.sheets.length < 2, onClick: () => deleteSheetUI(i) }]); };

  return (
    <section className="sheetwrap" aria-label="Tabelle"
      {...dz}>
      <div className={"sheet" + (over ? " over" : "")}>
        <div className="fxbar" aria-live="polite">
          <span className="addr">{sel ? (multi && rng ? `${ZF.colLetter(sh.c0 + rng.k1)}${rng.r1}:${ZF.colLetter(sh.c0 + rng.k2)}${rng.r2}` : ZF.colLetter(sh.c0 + sel.k) + sel.r) : "–"}</span><span className="fx">fx</span>
          <span className="val">{multi ? <span className="rsum">{rnum ? <>Summe <b>{de(Math.round(rsum * 100) / 100)}</b> · </> : null}Anzahl <b>{rcnt}</b>{rnum > 1 ? <> · Mittelwert <b>{de(Math.round(rsum / rnum * 100) / 100)}</b></> : null}</span> : sel ? (edit != null ? edit : rawOf(selCell)) : ""}</span>
        </div>
        <div className="scroll" ref={wrap} tabIndex={0} onKeyDown={onKey} aria-label="Tabellenbereich"
          onCopy={e => { if (edit != null || !rng) return; e.preventDefault(); e.clipboardData.setData("text/plain", copyText()); toast(multi ? `${(rng.r2 - rng.r1 + 1) * (rng.k2 - rng.k1 + 1)} Zellen kopiert` : "Kopiert"); }}
          onCut={e => { if (edit != null || !rng) return; e.preventDefault(); e.clipboardData.setData("text/plain", copyText()); clearRange(); }}
          onPaste={e => { if (edit != null) return; e.preventDefault(); pasteText(e.clipboardData.getData("text/plain")); }}>
          <table id="grid">
            <thead><tr><th className="rn" />{Array.from({ length: C }, (_, c) => <th key={c} className={rng && c - sh.c0 >= rng.k1 && c - sh.c0 <= rng.k2 ? "on" : ""} onMouseDown={e => selectCol(c - sh.c0, e)} onContextMenu={e => { if (sh.hr >= 0 && c - sh.c0 >= 0 && c - sh.c0 < sh.ncol) colMenu(e, c - sh.c0); }}>{ZF.colLetter(c)}</th>)}<th /></tr></thead>
            <tbody>
              {sh.top.map((r: any[], ri: number) => {
                const filled = (r || []).map((c: any, i: number) => ZF.blank(c) ? -1 : i).filter((i: number) => i >= 0);
                if (sh.hr >= 0 && filled.length <= 1) { const ci = filled.length ? filled[0] : 0, v0 = calc.value(sh, ri + 1, ci);
                  return <tr key={"t" + ri} className="top"><th className="rn">{ri + 1}</th>{ci > 0 && <td colSpan={ci} />}<td colSpan={C - ci + 1} className={ri === 0 ? "big" : ""}>{fmtVal(sh, -1, v0, (r || [])[ci])}</td></tr>; }
                return <tr key={"t" + ri} className="top"><th className="rn">{ri + 1}</th>{Array.from({ length: C }, (_, c) => <td key={c}>{fmtVal(sh, -1, calc.value(sh, ri + 1, c), (r || [])[c])}</td>)}<td /></tr>;
              })}
              {sh.hr >= 0 && <tr className="hdr"><th className="rn">{sh.hr + 1}</th>{pre(sh.c0)}{sh.headers.map((h: string, k: number) => <td key={k} title="Doppelklick: umbenennen · Rechtsklick: sortieren und mehr" onDoubleClick={() => renameCol(k)} onContextMenu={e => colMenu(e, k)}>{h}</td>)}<td className="addcol"><button type="button" title="Neue Spalte" aria-label="Neue Spalte" onClick={addCol}>+</button></td></tr>}
              {sh.hr >= 0 && sh.rows.slice(0, limit).map((row: any[], i: number) => {
                const r = ZF.dataRowNo(sh, i), kd = sh.kind && sh.kind[i], rowOn = rng && r >= rng.r1 && r <= rng.r2 && rng.k1 === 0 && rng.k2 === sh.ncol - 1;
                if (kd) return <tr key={r} className={"sep sep-" + kd} data-row={r} title={KIND[kd] + " – wird nicht als Datenzeile gezählt"}><th className="rn" onMouseDown={e => selectRow(r, e)} onContextMenu={e => rowMenu(e, r)}>{r}</th>{pre(sh.c0)}{row.map((_: any, k: number) => cellTd(row, r, k, kd !== "leer"))}<td /></tr>;
                return (
                  <tr key={r} data-row={r} className={(fresh[r] ? "fresh" : hits[r] ? "hit" : "") + (rowOn ? " rowon" : "")}>
                    <th className="rn" onMouseDown={e => selectRow(r, e)} onContextMenu={e => rowMenu(e, r)}>{r}</th>{pre(sh.c0)}{row.map((_: any, k: number) => cellTd(row, r, k, true))}
                    <td className="del"><button type="button" aria-label={`Zeile ${r} löschen`} title="Zeile löschen" onClick={() => deleteRows([r])}>×</button></td>
                  </tr>);
              })}
              {sh.hr >= 0 && limit < sh.rows.length && <tr className="add"><th className="rn" /><td colSpan={C + 1}><button type="button" onClick={() => setShowAll(true)}>Alle {sh.rows.length} Zeilen anzeigen</button></td></tr>}
              {sh.hr >= 0 && <tr className="add"><th className="rn" /><td colSpan={C + 1}><button type="button" onClick={() => {
                const r = ZF.appendRow(sh, [], S.model); done(`Leere Zeile ${r} hinzugefügt`, [r]);
                const k = sh.headers.findIndex((_: any, kk: number) => !ZF.computedCols(sh)[kk]); setAnchor(null); setSel({ r, k: Math.max(0, k) }); setTimeout(() => wrap.current?.focus(), 30); }}>+ Zeile hinzufügen</button></td></tr>}
              {sh.hr >= 0 && sh.foot.map((row: any[], j: number) => { const r = ZF.footRowNo(sh, j); return <tr key={"f" + r} className="foot" data-row={r}><th className="rn">{r}</th>{pre(sh.c0)}{row.map((_: any, k: number) => cellTd(row, r, k, true))}<td /></tr>; })}
            </tbody>
          </table>
        </div>
        <div className="sheettabs" role="tablist" aria-label="Blätter">
          {s.model.sheets.map((x: any, i: number) => <button key={x.name + i} role="tab" aria-selected={i === s.cur} title={x.name + " · Doppelklick: umbenennen · Rechtsklick: mehr"} onClick={() => selectSheet(i)} onDoubleClick={() => renameSheetUI(i)} onContextMenu={e => tabMenu(e, i)}>{x.name}</button>)}
          <button type="button" className="addtab" title="Neues Blatt" aria-label="Neues Blatt" onClick={addSheet}>+</button>
        </div>
      </div>
      <p className="foot">Klicken und tippen wie in Excel · Bereich mit Maus oder Umschalt markieren · {KEY("C")} / {KEY("V")} kopieren und einfügen (auch aus Excel) · {KEY("Z")} rückgängig · Rechtsklick für Zeilen, Spalten und Blätter. Formeln, Farben und Formate der Originaldatei bleiben beim Speichern erhalten; die Datei bleibt {HIER}.</p>
    </section>);
}
/* ---------- Blätter: neu, umbenennen, löschen ---------- */
export function addSheet() { const sh = ZF.newSheet(S.model); S.cur = S.model.sheets.indexOf(sh); commit(`Neues Blatt „${sh.name}“`); toast(`Blatt „${sh.name}“ angelegt – Doppelklick auf eine Überschrift benennt die Spalte um`); }
export async function renameSheetUI(i: number) { const sh = S.model.sheets[i]; const n = await promptBox("Blatt umbenennen", sh.name, "Umbenennen", "Formeln in anderen Blättern werden automatisch angepasst.");
  if (n == null) return; const r = ZF.renameSheet(S.model, sh, n); if (r.error) { toast(r.error); return; } commit(`Blatt umbenannt: „${sh.name}“`); }
export async function deleteSheetUI(i: number) { const sh = S.model.sheets[i], refs = ZF.sheetRefCount(S.model, sh);
  if (!(await confirmBox(`Blatt „${sh.name}“ löschen?`, (refs ? `Achtung: ${refs} Formel${refs === 1 ? "" : "n"} in anderen Blättern beziehen sich auf dieses Blatt und zeigen danach einen Fehler. ` : "") + `Mit ${KEY("Z")} holen Sie das Blatt zurück.`, "Löschen", true))) return;
  const r = ZF.deleteSheet(S.model, sh); if (r.error) { toast(r.error); return; } S.cur = Math.max(0, Math.min(i, S.model.sheets.length - 1)); commit(`Blatt „${sh.name}“ gelöscht`); }

function Empty() {
  const ex = SHOW_EXAMPLES ? Object.keys(window.ZF_EXAMPLES || {}) : [], inp = useRef<HTMLInputElement>(null);
  return (
    <div className="empty">
      <Mark />
      <h2>Excel bleibt. Das Tippen fällt weg.</h2>
      <p>Öffnen Sie Ihre Excel-Datei oder ziehen Sie sie hierher. Sie bleibt {IS_APP ? "auf Ihrem Rechner" : "in Ihrem Browser"}.</p>
      <div className="row center"><button className="btn" type="button" id="emptyOpen" onClick={() => inp.current?.click()}>Excel-Datei öffnen</button></div>
      <input ref={inp} type="file" hidden accept=".xlsx,.xlsm,.xls,.csv,.ods" onChange={e => { const f = e.target.files?.[0]; e.target.value = ""; if (f) openFile(f); }} />
      {ex.length > 0 && <><p className="hint">Oder mit einem Beispiel ausprobieren:</p>
        <div className="row center">{ex.slice(0, 5).map(n => <button key={n} className="btn quiet small" onClick={() => openExample(n)}>{n.replace(/_/g, " ").replace(/\.xlsx$/, "")}</button>)}</div></>}
    </div>);
}
export function Toast() { const s = useStore(); return s.toast ? <div className="toast" role="status">{s.toast}</div> : null; }
// Zeilen-Markierung nach 7 Sekunden ausblenden
let fadeT: any; export function fadeFresh() { clearTimeout(fadeT); fadeT = setTimeout(() => { S.fresh = {}; emit(); }, 7000); }
export { persist };
