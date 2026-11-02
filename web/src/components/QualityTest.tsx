// Dialog „KI-Qualitätstest“: 17 Fragen mit bekannten Antworten, Ergebnis je Frage, Zeit und Kosten, Vergleich mit früheren Läufen.
import { useEffect, useRef, useState, useSyncExternalStore } from "react";
import { CASES, runEval, type CaseResult } from "../eval";
import { toast } from "../store";

let open = false; const subs = new Set<() => void>();
export function openQualityTest() { open = true; subs.forEach(f => f()); }
function close() { open = false; subs.forEach(f => f()); }

const KEY = "zf-ki-test";
type Run = Awaited<ReturnType<typeof runEval>>;
const loadRuns = (): Run[] => { try { return JSON.parse(localStorage.getItem(KEY) || "[]"); } catch { return []; } };
const saveRuns = (r: Run[]) => { try { localStorage.setItem(KEY, JSON.stringify(r.slice(0, 12))); } catch { /* voll */ } };
const ICON: Record<string, string> = { ok: "✓", guard: "●", fail: "✗", error: "!" };
const LABEL: Record<string, string> = { ok: "korrekt", guard: "abgefangen", fail: "falsch", error: "Fehler" };
const score = (r: CaseResult[]) => ({ ok: r.filter(x => x.verdict === "ok").length, guard: r.filter(x => x.verdict === "guard").length, bad: r.filter(x => x.verdict === "fail" || x.verdict === "error").length });
const dur = (ms: number) => { const s = Math.round(ms / 1000); return s >= 60 ? `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")} min` : `${s} s`; };
const usd = (u: Run["usage"]) => u.costKnown && u.requests ? `${u.cost.toLocaleString("de-DE", { minimumFractionDigits: 3, maximumFractionDigits: 3 })} $` : "–";

export function QualityTest() {
  const isOpen = useSyncExternalStore(cb => { subs.add(cb); return () => subs.delete(cb); }, () => open);
  const dlg = useRef<HTMLDialogElement>(null), ctl = useRef<AbortController | null>(null);
  const [rows, setRows] = useState<CaseResult[]>([]), [run, setRun] = useState<Run | null>(null), [busy, setBusy] = useState(false), [runs, setRuns] = useState<Run[]>(loadRuns);
  useEffect(() => { const d = dlg.current; if (!d) return; if (isOpen && !d.open) d.showModal(); if (!isOpen && d.open) d.close(); }, [isOpen]);
  const start = async () => {
    if (!AI.active()) { toast("Zuerst die KI einrichten."); return; }
    setRows([]); setRun(null); setBusy(true); ctl.current = new AbortController();
    try { const r = await runEval(c => setRows(x => [...x, c].sort((a, b) => +a.id - +b.id)), ctl.current.signal);
      if (!ctl.current.signal.aborted) { setRun(r); const all = [r, ...loadRuns()]; saveRuns(all); setRuns(all.slice(0, 12)); } }
    finally { setBusy(false); ctl.current = null; }
  };
  const stop = () => ctl.current?.abort();
  const copy = () => { if (!run) return; const s = score(run.results);
    const txt = [`Zettelfrei KI-Qualitätstest · ${new Date(run.date).toLocaleString("de-DE")} · Modell: ${run.model}`,
      `${s.ok} korrekt · ${s.guard} abgefangen · ${s.bad} falsch · ${dur(run.ms)} · Kosten ${usd(run.usage)} · ${run.usage.requests} Anfragen`, "",
      ...run.results.map(r => `${ICON[r.verdict]} ${r.id}. ${r.frage} — ${LABEL[r.verdict]}: ${r.note}\n   Antwort: ${r.answer || "–"}`)].join("\n");
    navigator.clipboard.writeText(txt).then(() => toast("Ergebnis kopiert"), () => toast("Kopieren nicht möglich")); };
  const s = score(rows);
  return (
    <dialog ref={dlg} id="qualitytest" className="wide" onClose={() => { stop(); close(); }}>
      <div className="qt">
        <h2>KI-Qualitätstest</h2>
        <p className="note">17 Fragen an eine Testdatei mit bekannten Antworten – Auswertungen, mehrdeutige Namen, Änderungen, Anweisungen in Zellen. Ihre geöffnete Datei bleibt unberührt.
          Kostet je nach Modell etwa 1–10 Cent. <b>✓ korrekt</b> = Modell richtig · <b>● abgefangen</b> = Modell falsch, Zettelfrei warnt oder verhindert · <b>✗ falsch</b>.</p>
        <div className="row">
          {busy ? <button className="btn quiet" type="button" onClick={stop}>Abbrechen</button> : <button className="btn" type="button" id="qtStart" onClick={start}>{run ? "Noch einmal testen" : "Test starten"}</button>}
          {run && !busy && <button className="btn quiet" type="button" onClick={copy}>Ergebnis kopieren</button>}
          <span className="hint">{AI.status().label}</span>
          {(busy || rows.length > 0) && <span className="qtscore"><b className="ok">{s.ok} ✓</b> <b className="guard">{s.guard} ●</b> <b className="bad">{s.bad} ✗</b> · {rows.length}/{CASES.length}</span>}
        </div>
        {run && <p className="qtsum">{dur(run.ms)} · Kosten {usd(run.usage)} · {run.usage.requests} Anfragen · {(run.usage.input / 1000).toFixed(0)}k Tokens rein, {(run.usage.output / 1000).toFixed(1)}k raus</p>}
        <ol className="qtlist">{CASES.map(c => { const r = rows.find(x => x.id === c.id);
          return <li key={c.id} className={r ? r.verdict : busy ? "wait" : ""} title={r?.answer || ""}>
            <span className="ic">{r ? ICON[r.verdict] : busy ? "…" : "·"}</span><span className="q">{c.frage}</span><span className="n">{r ? r.note : ""}</span><span className="t">{r ? dur(r.ms) : ""}</span></li>; })}</ol>
        {runs.length > 0 && <><h3>Frühere Läufe</h3>
          <table className="qtruns"><thead><tr><th>Datum</th><th>Modell</th><th>✓</th><th>●</th><th>✗</th><th>Zeit</th><th>Kosten</th></tr></thead>
            <tbody>{runs.map((r, i) => { const x = score(r.results); return <tr key={i}><td>{new Date(r.date).toLocaleString("de-DE", { dateStyle: "short", timeStyle: "short" })}</td><td>{r.model}</td><td>{x.ok}</td><td>{x.guard}</td><td>{x.bad}</td><td>{dur(r.ms)}</td><td>{usd(r.usage)}</td></tr>; })}</tbody></table></>}
        <div className="row end"><button className="btn quiet" type="button" onClick={() => { stop(); close(); }}>Schließen</button></div>
      </div>
    </dialog>);
}
