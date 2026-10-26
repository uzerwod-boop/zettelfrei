import { useEffect, useMemo, useRef, useState } from "react";
import { S, commit, emit, issues, mark, openExample, openFile, persist, saveExcel, selectSheet, setPane, sheet, toast, useStore, type Pane } from "../store";
import { de, fmtVal, rawOf } from "../util";
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
        {last && <span className="saved">· gesichert {new Date(last.t).toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" })}</span>}
      </div>
      <div className="filebar">
        <select aria-label="Beispiel öffnen" value="" onChange={e => { if (e.target.value) openExample(e.target.value); }}>
          <option value="">Beispiel öffnen…</option>
          {examples.map(n => <option key={n} value={n}>{n.replace(/_/g, " ").replace(/\.xlsx$/, "")}</option>)}
        </select>
        <button className="btn quiet" onClick={() => input.current?.click()}>Excel öffnen</button>
        <input ref={input} type="file" hidden accept=".xlsx,.xlsm,.xls,.csv,.ods" onChange={e => { const f = e.target.files?.[0]; e.target.value = ""; if (f) openFile(f); }} />
        <button className="btn" onClick={saveExcel} disabled={!s.model}>Als Excel speichern</button>
      </div>
    </header>);
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

/* ---------- Tabelle wie in Excel: auswählen, tippen, Enter, Tab, Pfeile, Esc ---------- */
type Sel = { r: number; k: number } | null;
const KIND: Record<string, string> = { leer: "Leerzeile", summe: "Zwischensumme", abschnitt: "Abschnitt", notiz: "Notiz" };
export function Sheet() {
  const s = useStore(); const sh = sheet();
  const [sel, setSel] = useState<Sel>(null), [edit, setEdit] = useState<string | null>(null), [showAll, setShowAll] = useState(false);
  const wrap = useRef<HTMLDivElement>(null), inp = useRef<HTMLInputElement>(null), drag = useRef(0);
  const [over, setOver] = useState(false);
  useEffect(() => { setSel(null); setEdit(null); setShowAll(false); }, [s.cur, s.model]);
  // Sprung von außen (Hinweis, Chat): Zeile zeigen und markieren
  useEffect(() => { const f = s.focus; if (!f || !sh || f.sheet !== sh.name) return;
    const tr = wrap.current?.querySelector(`tr[data-row="${f.r}"]`) as HTMLElement | null; tr?.scrollIntoView({ block: "center", behavior: "smooth" });
    if (f.k != null) { setSel({ r: f.r, k: f.k }); setTimeout(() => wrap.current?.focus(), 30); } }, [s.focus?.n]);
  useEffect(() => { if (edit != null) inp.current?.focus(); }, [edit != null]);

  const view = useMemo(() => {
    if (!sh) return null;
    const calc = new ZF.Calc(s.model), flag: Record<string, string> = {};
    issues().filter((x: any) => x.blatt === sh.name).forEach((x: any) => { const key = x.zeile + ":" + (x.k != null ? x.k : ZF.findCol(sh, x.spalte)); flag[key] = (flag[key] ? flag[key] + "; " : "") + x.problem; });
    const comp = sh.hr >= 0 ? ZF.computedCols(sh) : [], numc = sh.hr >= 0 ? sh.headers.map((_: any, k: number) => comp[k] || ZF.numericCol(sh, k)) : [];
    return { calc, flag, numc };
  }, [s.rev, s.cur]);
  if (!s.model || !sh || !view) return <section className="sheetwrap"><div className="sheet empty-state"><Empty /></div></section>;
  const { calc, flag, numc } = view;
  const C = sh.hr < 0 ? Math.max(1, ...sh.top.map((r: any[]) => (r || []).length)) : sh.c0 + sh.ncol;
  const editable = (r: number) => { if (sh.hr < 0 || r <= sh.hr + 1) return false; const i = r - sh.hr - 2; return !(i < sh.rows.length && sh.kind && sh.kind[i] === "leer"); };
  const lastRow = sh.hr < 0 ? sh.top.length : sh.hr + 1 + sh.rows.length + sh.foot.length;
  const selCell = sel ? ZF.cellAt(sh, sel.r, sh.c0 + sel.k) : null;

  const commitEdit = (move?: [number, number]) => {
    if (!sel || edit == null) return;
    const before = rawOf(selCell);
    if (edit.trim() !== before) {
      ZF.setCell(sh, sel.r, sel.k, edit.trim(), { formula: true });
      if (sh.timeCol && sh.timeCol[sel.k]) { const tc = ZF.cellAt(sh, sel.r, sh.c0 + sel.k); if (tc && typeof tc.v === "number") tc.tm = 1; }
      commit(`${ZF.colLetter(sh.c0 + sel.k)}${sel.r} geändert: „${edit.length > 30 ? edit.slice(0, 30) + "…" : edit.trim()}“`);
    }
    setEdit(null); if (move) moveSel(move[0], move[1]); wrap.current?.focus();
  };
  const moveSel = (dr: number, dk: number) => setSel(p => {
    if (!p) return p; let r = p.r + dr, k = Math.max(0, Math.min(sh.ncol - 1, p.k + dk));
    while (r > sh.hr + 1 && r <= lastRow && !editable(r) && dr) r += dr;
    if (r <= sh.hr + 1 || r > lastRow) r = p.r; return { r, k }; });
  const onKey = (e: React.KeyboardEvent) => {
    if (!sel || edit != null) return;
    const m: Record<string, [number, number]> = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Enter: [1, 0], Tab: [0, e.shiftKey ? -1 : 1] };
    if (m[e.key]) { e.preventDefault(); moveSel(...m[e.key]); return; }
    if (e.key === "F2") { e.preventDefault(); setEdit(rawOf(selCell)); return; }
    if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); if (rawOf(selCell)) { ZF.setCell(sh, sel.r, sel.k, "", { formula: true }); commit(`${ZF.colLetter(sh.c0 + sel.k)}${sel.r} geleert`); } return; }
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); setEdit(e.key); }
  };
  const cellTd = (row: any[], r: number, k: number, canEdit: boolean) => {
    const cell = row[k], v = calc.value(sh, r, sh.c0 + k), fl = flag[r + ":" + k], isSel = sel && sel.r === r && sel.k === k;
    const cls = [(typeof v === "number" || (numc[k] && (v == null || v === ""))) ? "num" : "", cell?.f ? "fxc" : "", fl ? "flag" : "", isSel ? "sel" : ""].join(" ").trim();
    return (
      <td key={k} className={cls} title={fl || undefined} data-r={r} data-k={k}
        onMouseDown={canEdit ? () => { if (edit != null) commitEdit(); setSel({ r, k }); } : undefined}
        onDoubleClick={canEdit ? () => { setSel({ r, k }); setEdit(rawOf(cell)); } : undefined}>
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

  return (
    <section className="sheetwrap" aria-label="Tabelle"
      onDragOver={e => { if ([...e.dataTransfer.items].some(i => i.kind === "file")) { e.preventDefault(); drag.current++; setOver(true); } }}
      onDragLeave={() => setOver(false)}
      onDrop={e => { const f = e.dataTransfer.files[0]; setOver(false); if (!f) return; e.preventDefault();
        if (/\.(xlsx|xlsm|xls|csv|ods)$/i.test(f.name)) openFile(f); else { setPane("pdf"); window.dispatchEvent(new CustomEvent("zf-beleg", { detail: f })); } }}>
      <div className={"sheet" + (over ? " over" : "")}>
        <div className="fxbar" aria-live="polite">
          <span className="addr">{sel ? ZF.colLetter(sh.c0 + sel.k) + sel.r : "–"}</span><span className="fx">fx</span>
          <span className="val">{sel ? (edit != null ? edit : rawOf(selCell)) : ""}</span>
        </div>
        <div className="scroll" ref={wrap} tabIndex={0} onKeyDown={onKey} aria-label="Tabellenbereich">
          <table id="grid">
            <thead><tr><th className="rn" />{Array.from({ length: C }, (_, c) => <th key={c}>{ZF.colLetter(c)}</th>)}<th /></tr></thead>
            <tbody>
              {sh.top.map((r: any[], ri: number) => {
                const filled = (r || []).map((c: any, i: number) => ZF.blank(c) ? -1 : i).filter((i: number) => i >= 0);
                if (sh.hr >= 0 && filled.length <= 1) { const ci = filled.length ? filled[0] : 0, v0 = calc.value(sh, ri + 1, ci);
                  return <tr key={"t" + ri} className="top"><th className="rn">{ri + 1}</th>{ci > 0 && <td colSpan={ci} />}<td colSpan={C - ci + 1} className={ri === 0 ? "big" : ""}>{fmtVal(sh, -1, v0, (r || [])[ci])}</td></tr>; }
                return <tr key={"t" + ri} className="top"><th className="rn">{ri + 1}</th>{Array.from({ length: C }, (_, c) => <td key={c}>{fmtVal(sh, -1, calc.value(sh, ri + 1, c), (r || [])[c])}</td>)}<td /></tr>;
              })}
              {sh.hr >= 0 && <tr className="hdr"><th className="rn">{sh.hr + 1}</th>{pre(sh.c0)}{sh.headers.map((h: string, k: number) => <td key={k}>{h}</td>)}<td /></tr>}
              {sh.hr >= 0 && sh.rows.slice(0, limit).map((row: any[], i: number) => {
                const r = ZF.dataRowNo(sh, i), kd = sh.kind && sh.kind[i];
                if (kd) return <tr key={r} className={"sep sep-" + kd} data-row={r} title={KIND[kd] + " – wird nicht als Datenzeile gezählt"}><th className="rn">{r}</th>{pre(sh.c0)}{row.map((_: any, k: number) => cellTd(row, r, k, kd !== "leer"))}<td /></tr>;
                return (
                  <tr key={r} data-row={r} className={fresh[r] ? "fresh" : hits[r] ? "hit" : ""}>
                    <th className="rn">{r}</th>{pre(sh.c0)}{row.map((_: any, k: number) => cellTd(row, r, k, true))}
                    <td className="del"><button type="button" aria-label={`Zeile ${r} löschen`} title="Zeile löschen"
                      onClick={() => { ZF.deleteRow(sh, r, S.model); S.fresh = {}; S.hits = {}; setSel(null); commit(`Zeile ${r} gelöscht`); toast(`Zeile ${r} gelöscht. Im Verlauf zurückholen.`); }}>×</button></td>
                  </tr>);
              })}
              {sh.hr >= 0 && limit < sh.rows.length && <tr className="add"><th className="rn" /><td colSpan={C + 1}><button type="button" onClick={() => setShowAll(true)}>Alle {sh.rows.length} Zeilen anzeigen</button></td></tr>}
              {sh.hr >= 0 && <tr className="add"><th className="rn" /><td colSpan={C + 1}><button type="button" onClick={() => {
                const r = ZF.appendRow(sh, [], S.model); mark("fresh", sh.name, [r]); commit(`Leere Zeile ${r} hinzugefügt`);
                const k = sh.headers.findIndex((_: any, kk: number) => !ZF.computedCols(sh)[kk]); setSel({ r, k: Math.max(0, k) }); setTimeout(() => wrap.current?.focus(), 30); }}>+ Zeile hinzufügen</button></td></tr>}
              {sh.hr >= 0 && sh.foot.map((row: any[], j: number) => { const r = ZF.footRowNo(sh, j); return <tr key={"f" + r} className="foot" data-row={r}><th className="rn">{r}</th>{pre(sh.c0)}{row.map((_: any, k: number) => cellTd(row, r, k, true))}<td /></tr>; })}
            </tbody>
          </table>
        </div>
        <div className="sheettabs" role="tablist" aria-label="Blätter">
          {s.model.sheets.map((x: any, i: number) => <button key={x.name} role="tab" aria-selected={i === s.cur} title={x.name} onClick={() => selectSheet(i)}>{x.name}</button>)}
        </div>
      </div>
      <p className="foot">Zelle anklicken und lostippen, Enter speichert und springt nach unten, Esc bricht ab, mit „=“ beginnt eine Formel. Beim Speichern bleiben Werte, Formeln, Farben und Zahlenformate der Originaldatei erhalten. Die Datei bleibt in diesem Browser; an die KI geht nur, was Sie ausführen.</p>
    </section>);
}

function Empty() {
  const ex = Object.keys(window.ZF_EXAMPLES || {});
  return (
    <div className="empty">
      <Mark />
      <h2>Excel bleibt. Das Tippen fällt weg.</h2>
      <p>Öffnen Sie Ihre Excel-Datei oder ziehen Sie sie hierher. Sie bleibt in Ihrem Browser.</p>
      <div className="row center">{ex.slice(0, 5).map(n => <button key={n} className="btn quiet small" onClick={() => openExample(n)}>{n.replace(/_/g, " ").replace(/\.xlsx$/, "")}</button>)}</div>
    </div>);
}
export function Toast() { const s = useStore(); return s.toast ? <div className="toast" role="status">{s.toast}</div> : null; }
// Zeilen-Markierung nach 7 Sekunden ausblenden
let fadeT: any; export function fadeFresh() { clearTimeout(fadeT); fadeT = setTimeout(() => { S.fresh = {}; emit(); }, 7000); }
export { persist };
