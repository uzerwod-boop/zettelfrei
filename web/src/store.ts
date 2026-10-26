// Zustand der Anwendung. Das Modell wird von den Kernmodulen (ZF) direkt verändert;
// nach jeder Änderung erhöht emit() die Revision und React zeichnet neu.
import { useSyncExternalStore } from "react";
import { b64, firma, rawOf, stamp, toB64 } from "./util";

export type Pane = "cmd" | "pdf" | "docs" | "check" | "hist";
export interface HistEntry { t: number; label: string; snap: string }
export interface ProtEntry { zeit: string; wer: string; aktion: string; details: string[] }
export interface Msg { id: number; text: string; status: "busy" | "done" | "error"; res?: any; results?: any; plan?: any[]; planState?: "open" | "done" | "dropped"; error?: string; noAi?: boolean }

const KEY = "zf3", KEY_ORIG = "zf3-orig";
export const S = {
  model: null as any, cur: 0, history: [] as HistEntry[], protokoll: [] as ProtEntry[], origBytes: null as Uint8Array | null,
  fresh: {} as Record<string, Record<number, 1>>, hits: {} as Record<string, Record<number, 1>>, aiIssues: [] as any[],
  pane: "cmd" as Pane, msgs: [] as Msg[], toast: "", rev: 0, focus: null as null | { sheet: string; r: number; k: number | null; n: number },
  protSheet: true, ai: { on: false, label: "KI wird verbunden…" },
};
const subs = new Set<() => void>();
export function emit() { S.rev++; subs.forEach(f => f()); }
export function useStore() { useSyncExternalStore(cb => { subs.add(cb); return () => subs.delete(cb); }, () => S.rev); return S; }

let toastTimer: any;
export function toast(m: string) { S.toast = m; emit(); clearTimeout(toastTimer); toastTimer = setTimeout(() => { S.toast = ""; emit(); }, 3200); }
export const sheet = () => S.model ? S.model.sheets[S.cur] : null;
export function setPane(p: Pane) { S.pane = p; emit(); }
export function selectSheet(i: number) { S.cur = i; persist(); emit(); }
export function mark(store: "fresh" | "hits", sheetName: string, rows: number[]) { S[store] = { [sheetName]: Object.fromEntries(rows.map(r => [r, 1])) } as any; }
export function goTo(sheetName: string, row: number, k?: number | null) {
  const i = S.model.sheets.findIndex((s: any) => s.name === sheetName); if (i < 0) return;
  S.cur = i; S.focus = { sheet: sheetName, r: row, k: k ?? null, n: Date.now() }; emit();
}

/* ---------- Speichern im Browser ---------- */
let persistWarned = false;
export function persist() {
  if (!S.model) return;
  for (const [h, p] of [[15, 2000], [3, 500], [0, 100]] as const) {
    try { localStorage.setItem(KEY, JSON.stringify({ model: S.model, history: h ? S.history.slice(-h) : [], protokoll: S.protokoll.slice(-p), cur: S.cur })); return; }
    catch (e: any) { if (!/quota/i.test(String(e?.name) + String(e?.message))) return; }
  }
  if (!persistWarned) { persistWarned = true; toast("Browser-Speicher voll: Änderungen bitte mit „Speichern“ als Excel sichern."); }
}
function persistOrig() { try { if (S.origBytes && S.origBytes.length < 3e6) localStorage.setItem(KEY_ORIG, toB64(S.origBytes)); else localStorage.removeItem(KEY_ORIG); } catch { /* voll */ } }
export function restore() {
  try {
    const s = JSON.parse(localStorage.getItem(KEY) || "null");
    if (s?.model?.sheets) { S.model = s.model; S.history = s.history || []; S.protokoll = s.protokoll || []; S.cur = Math.min(s.cur || 0, S.model.sheets.length - 1);
      const o = localStorage.getItem(KEY_ORIG); if (o) S.origBytes = b64(o); return true; }
  } catch { /* defekt */ }
  return false;
}

/* ---------- Verlauf und Änderungsprotokoll ---------- */
function diffDetails(prevJson: string, now: any) {
  const out: string[] = []; let prev: any; try { prev = JSON.parse(prevJson); } catch { return out; }
  now.sheets.forEach((sh: any) => {
    const ps = prev.sheets.find((x: any) => x.name === sh.name); if (!ps || sh.hr < 0) return;
    if (ps.rows.length !== sh.rows.length) out.push(`${sh.name}: ${sh.rows.length > ps.rows.length ? "+" : ""}${sh.rows.length - ps.rows.length} Zeile(n)`);
    const n = Math.min(ps.rows.length, sh.rows.length);
    for (let i = 0; i < n && out.length < 25; i++) for (let k = 0; k < sh.ncol; k++) {
      const a = rawOf(ps.rows[i][k]), b = rawOf(sh.rows[i][k]);
      if (a !== b && !(sh.rows[i][k] && sh.rows[i][k].f)) out.push(`${sh.name}!${ZF.colLetter(sh.c0 + k)}${ZF.dataRowNo(sh, i)}: „${a || "leer"}“ → „${b || "leer"}“`);
    }
    if (ps.ncol !== sh.ncol) out.push(`${sh.name}: neue Spalte „${sh.headers.slice(ps.ncol).join("“, „")}“`);
  });
  return out;
}
export function commit(label: string) {
  const prev = S.history.length ? S.history[S.history.length - 1].snap : null;
  S.protokoll.push({ zeit: stamp(), wer: firma().bearbeiter || "", aktion: label, details: prev ? diffDetails(prev, S.model) : [] });
  S.history.push({ t: Date.now(), label, snap: JSON.stringify(S.model) }); if (S.history.length > 40) S.history = S.history.slice(-40);
  persist(); emit();
}
export function restoreVersion(i: number) {
  const h = S.history[i]; if (!h) return;
  S.model = JSON.parse(h.snap); S.cur = Math.min(S.cur, S.model.sheets.length - 1); S.fresh = {}; S.hits = {};
  commit("Zurückgeholt: Stand „" + h.label + "“"); toast("Früherer Stand zurückgeholt");
}

/* ---------- Datei öffnen / speichern ---------- */
function readWorkbook(data: Uint8Array, name: string): Promise<any> {
  try {
    if (/\.(csv|txt)$/i.test(name)) return Promise.resolve(ZF.loadCSV(XLSX, new TextDecoder("utf-8").decode(data), name));
    if (/\.(xlsx|xlsm)$/i.test(name) && window.ExcelJS) return ZF.loadWorkbookExcelJS(XLSX, ExcelJS, data, name);
    return Promise.resolve(ZF.loadWorkbook(XLSX, data, name));
  } catch (e) { return Promise.reject(e); }
}
export function openData(data: Uint8Array, name: string) {
  if (!window.XLSX) { toast("Der Excel-Leser konnte nicht geladen werden. Bitte Seite neu laden."); return; }
  if (data.length > 25 * 1024 * 1024) { toast("Die Datei ist größer als 25 MB – bitte nur die benötigten Blätter speichern."); return; }
  readWorkbook(data, name).then(m => {
    if (!m.sheets.length) { toast("Die Datei ist leer."); return; }
    S.model = m; S.cur = Math.max(0, m.sheets.findIndex((s: any) => s.hr >= 0)); S.fresh = {}; S.hits = {}; S.aiIssues = []; S.msgs = [];
    S.history = [{ t: Date.now(), label: "Geöffnet: " + name, snap: JSON.stringify(m) }];
    S.origBytes = /\.xlsx$|\.xlsm$/i.test(name) ? new Uint8Array(data) : null; persistOrig();
    S.protokoll.push({ zeit: stamp(), wer: firma().bearbeiter || "", aktion: "Datei geöffnet: " + name, details: [] });
    persist(); emit();
    const parsed = m.sheets.filter((s: any) => s.hr >= 0);
    toast(parsed.length ? `${name} geöffnet · ${parsed.map((s: any) => `${ZF.dataCount(s)} Zeilen in „${s.name}“`).join(", ")}` : "Keine Tabelle mit Überschriften gefunden.");
  }, () => toast("Diese Datei lässt sich nicht lesen."));
}
export function openExample(name: string) { const ex = window.ZF_EXAMPLES || {}; if (ex[name]) openData(b64(ex[name]), name); }
export function openFile(f: File) { f.arrayBuffer().then(b => openData(new Uint8Array(b), f.name)); }

let downloads: any = null;
if (window.claude?.use) window.claude.use("downloads").then((d: any) => { downloads = d; }, () => {});
export function saveFile(name: string, data: BlobPart | Blob, mime?: string): Promise<void> {
  const blob = data instanceof Blob ? data : new Blob([data], { type: mime || "application/octet-stream" });
  const tauri = (window as any).__TAURI__;
  if (tauri?.core) // Desktop-Programm: „Speichern unter“-Dialog des Systems
    return blob.arrayBuffer().then(b => tauri.core.invoke("save_file", { name, data: toB64(new Uint8Array(b)) }))
      .then((p: string | null) => { if (p) toast(name + " gespeichert"); }, () => toast("Speichern fehlgeschlagen."));
  if (downloads) return downloads.save({ filename: name, data: blob }).then(() => toast(name + " gespeichert"), (e: any) => { if (e?.code !== "declined") toast("Speichern ist hier nicht möglich."); });
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000); toast(name + " gespeichert"); return Promise.resolve();
}
export function saveExcel() {
  if (!S.model || !window.XLSX) return;
  let name = S.model.name.replace(/\.(xls|csv|ods|xlsm)$/i, ".xlsx"); if (!/\.xlsx$/i.test(name)) name += ".xlsx";
  const plain = () => { try { return Promise.resolve(ZF.buildWorkbook(XLSX, S.model)); } catch (e) { return Promise.reject(e); } };
  const styled = (window.ExcelJS && S.origBytes && typeof ZFX !== "undefined")
    ? ZFX.save(ExcelJS, ZF, S.origBytes, S.model, { protokoll: S.protSheet ? S.protokoll : null }).catch(plain) : plain();
  styled.then((buf: any) => saveFile(name, buf, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), () => toast("Speichern fehlgeschlagen."));
}

/* ---------- Prüfen ---------- */
export function issues() { return S.model ? ZF.checks(S.model, { today: new Date() }).concat(S.aiIssues) : []; }

/* ---------- KI-Status ---------- */
AI.onChange((st: any) => { S.ai = { on: !!st.on, label: st.label }; emit(); });
