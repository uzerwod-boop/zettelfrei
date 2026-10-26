import { useEffect, useMemo, useRef, useState } from "react";
import { S, commit, emit, goTo, mark, setPane, toast, useStore, type Msg } from "../store";
import { TODAY, WEEKDAY, de, fmtVal } from "../util";
import { fadeFresh } from "./Workspace";
import { Icon, Mic } from "./Icons";
import { openSettings } from "./Settings";

export const RULES = 'Antworte NUR mit JSON in genau dieser Form:\n{"aenderungen":[ {"typ":"neu","blatt":"…","werte":{"Spaltenname":Wert}} | {"typ":"aendern","blatt":"…","zeile":Excel-Zeilennummer,"werte":{"Spaltenname":Wert}} | {"typ":"loeschen","blatt":"…","zeile":Excel-Zeilennummer} ],\n "abfragen":[{"id":"q1","blatt":"…","spalte":"…","art":"summe|anzahl|mittelwert|min|max|liste","filter":[{"spalte":"…","op":"=|!=|enthält|>|<|>=|<=|leer|nicht leer|in|nicht in","wert":"…"}],"gruppieren":"Spalte oder Datumsspalte:tag|woche|monat|jahr (optional)","sortierung":"absteigend|aufsteigend|zeit (optional)","limit":Zahl (optional)} | {"id":"q3","ausdruck":"q1-q2 oder q1/q2*100","prozent":true|false,"euro":true|false}],\n "antwort":"kurzer deutscher Text, Ergebnisse von Abfragen als {q1} einsetzen",\n "rueckfrage":null}\nRegeln:\n- Rechne Summen, Anzahlen und Mittelwerte NIE selbst. Lege dafür immer eine Abfrage an und setze {id} in die Antwort ein. Für Listen von Namen/Artikeln nutze art "liste".\n- „Je/pro/nach …“ (z. B. Stunden je Baustelle, Umsatz pro Monat, Top 5 Kunden): eine Abfrage mit "gruppieren" statt vieler Einzelabfragen. Zeiträume: "Datum:monat" usw.\n- Unterschiede, Anteile, Prozent: zuerst die Werte als Abfragen, dann eine Abfrage mit "ausdruck" (nur q-IDs, Zahlen, + - * / und Klammern).\n- Verknüpfen von Blättern: erst eine Liste (art "liste") aus Blatt A, dann in Blatt B filtern mit {"op":"in","wert":"{q1}"}.\n- Zu- und Abgänge (z. B. Bestand) als {"plus": Zahl} angeben, z. B. {"Bestand":{"plus":-3}}.\n- Spalten mit „berechnet“ niemals setzen, die Tabelle rechnet sie.\n- Spaltennamen und Blattnamen exakt wie angegeben. Datum als TT.MM.JJJJ, Zahlen als Zahl.\n- Zeilennummern sind Excel-Zeilennummern aus der Spalte „Zeile“.\n- Ändere nur, was verlangt ist. Erfinde keine Werte. Ist unklar, welche Zeile gemeint ist, oder fehlen Pflichtangaben, stelle eine kurze rueckfrage und ändere nichts.\n- Leere Listen weglassen ist erlaubt.';
export const context = () => `Heute ist ${TODAY} (${WEEKDAY}). Du arbeitest in der Excel-Datei eines kleinen deutschen Handwerksbetriebs.\n\n` + ZF.describe(S.model, 180000);
export function aiError(e: any): string {
  const c = e?.code;
  const m: Record<string, string> = {
    cancelled: "", no_provider: "Keine KI eingerichtet. Oben im Assistenten auf den KI-Status klicken.", auth: "Der API-Schlüssel wurde abgelehnt. Bitte in den KI-Einstellungen prüfen.",
    credits: "Beim KI-Anbieter ist kein Guthaben mehr. Bitte aufladen (bei OpenRouter unter openrouter.ai/credits).", model: "Dieses Modell kennt der Anbieter nicht. Modellnamen in den KI-Einstellungen prüfen.",
    network: "Keine Verbindung zur KI. Internet, Server oder lokales Modell prüfen.", rate_limited: "Gerade zu viele Anfragen. Bitte in einer Minute noch einmal.",
    image_rejected: "Diese Datei lässt sich nicht lesen. Bitte ein anderes Foto oder PDF wählen.", images_unavailable: "Bilder kann die KI hier nicht lesen. Ein PDF mit Text funktioniert.",
    prompt_too_large: "Die Datei ist für eine Anfrage zu groß. Bitte ein kleineres Blatt öffnen.", invalid_json: "Die Antwort war unvollständig. Bitte noch einmal versuchen oder einfacher formulieren." };
  if (c === "http") return `Die KI hat einen Fehler gemeldet (${e.message}). Modellname in den Einstellungen prüfen.`;
  if (["not_granted", "sampling_disabled", "not_declared", "capability_disabled"].includes(c)) return "Die KI wurde für diese Seite nicht freigegeben. Tabelle bearbeiten, öffnen und speichern funktioniert weiter.";
  return c in m ? m[c] : "Die KI hat nicht geantwortet. Bitte noch einmal versuchen.";
}

let msgId = 1, ctl: AbortController | null = null;
export function send(text: string) {
  text = text.trim(); if (!text || !S.model) return;
  const msg: Msg = { id: msgId++, text, status: "busy" }; S.msgs = [...S.msgs, msg]; emit();
  const done = (patch: Partial<Msg>) => { Object.assign(msg, patch); S.msgs = [...S.msgs]; emit(); };
  if (!AI.active()) { done({ status: "error", noAi: true }); return; }
  ctl = new AbortController();
  AI.json(context() + `\n\nAnweisung oder Frage des Nutzers: """${text}"""\n\n` + RULES, { signal: ctl.signal }).then((res: any) => {
    if (!res || typeof res !== "object") { done({ status: "error", error: "Keine verwertbare Antwort." }); return; }
    const results = ZF.runQueries(S.model, res.abfragen || []), plan = ZF.planOps(S.model, res.aenderungen || []);
    done({ status: "done", res, results, plan, planState: plan.length ? "open" : undefined });
  }, (e: any) => done({ status: "error", error: aiError(e) || "Abgebrochen." })).finally(() => { ctl = null; });
}

function fmtQuery(r: any) {
  if (!r || r.error) return "(?)";
  if (r.groups) return r.groups.slice(0, 8).map((g: any) => `${g.key}: ${de(g.value, r.money ? 2 : null)}${r.money ? " €" : ""}`).join(" · ") + (r.groups.length > 8 ? " …" : "");
  if (r.expr && r.percent) return de(r.value, 1) + " %";
  if (r.expr && r.money) return de(r.value, 2) + " €";
  if (r.list) { const l = r.list.map((v: any) => typeof v === "number" ? de(v) : String(v)); return !l.length ? "keine" : l.length > 15 ? l.slice(0, 15).join(", ") + ` und ${l.length - 15} weitere` : l.join(", "); }
  if (r.value == null) return "–";
  return de(r.value, r.money ? 2 : null);
}

function addResultSheet(r: any) {
  const base = ("Auswertung " + r.by).replace(/[\\/?*[\]:]/g, " ").slice(0, 28); let name = base, n = 2;
  while (S.model.sheets.some((s: any) => s.name.toLowerCase() === name.toLowerCase())) name = base.slice(0, 26) + " " + (n++);
  const rows = r.groups.map((g: any) => [{ v: typeof g.key === "number" ? g.key : String(g.key) }, { v: g.value }]), sumF = r.art === "summe" || r.art === "anzahl";
  S.model.sheets.push({ name, hr: 0, top: [], headers: [r.by, r.of], rows, kind: rows.map(() => ""), foot: sumF ? [[{ v: "Gesamt" }, { f: `SUM(B2:B${rows.length + 1})` }]] : [],
    c0: 0, ncol: 2, widths: [28, 16], merges: [], colFmt: [null, r.money ? '#,##0.00 "€"' : "#,##0.##"], dateCol: [false, false], timeCol: [false, false], extra: {}, isNew: true, orig: { data: [], foot: [] } });
  S.cur = S.model.sheets.length - 1; commit(`Auswertung als Blatt „${name}“`); toast(`Blatt „${name}“ angelegt – wird beim Speichern in die Excel-Datei übernommen`);
  return name;
}

function GroupCard({ r }: { r: any }) {
  const [saved, setSaved] = useState<string | null>(null);
  const fmt = (v: any) => v == null ? "–" : de(v, r.money ? 2 : null) + (r.money ? " €" : "");
  const shown = r.groups.slice(0, 12), max = Math.max(1e-9, ...shown.map((g: any) => Math.abs(g.value || 0)));
  const total = r.groups.reduce((a: number, g: any) => a + (g.value || 0), 0);
  return (
    <div className="gcard">
      <div className="gh"><span>{r.of} je {r.by}</span><span className="hint">{r.groups.length} Gruppe{r.groups.length === 1 ? "" : "n"}</span></div>
      <div className="gchart" role="img" aria-label={`Balkendiagramm ${r.of} je ${r.by}`}>
        {shown.map((g: any, i: number) => (
          <div className="gbar" key={String(g.key)} style={{ animationDelay: i * 40 + "ms" }}>
            <span className="gk" title={String(g.key)}>{String(g.key)}</span>
            <span className="gtrack"><span className="gfill" style={{ width: Math.max(1, Math.round(Math.abs(g.value || 0) / max * 100)) + "%" }} /></span>
            <b>{fmt(g.value)}</b>
          </div>))}
      </div>
      {r.groups.length > 12 && <p className="hint">Diagramm zeigt die ersten 12 von {r.groups.length}.</p>}
      {(r.art === "summe" || r.art === "anzahl") && <div className="gsum"><span>Gesamt</span><b>{fmt(Math.round(total * 100) / 100)}</b></div>}
      <div className="row"><button type="button" className="btn quiet small" disabled={!!saved} onClick={() => setSaved(addResultSheet(r))}>{saved ? `Blatt „${saved}“ angelegt` : "Als Blatt in die Excel-Datei"}</button></div>
    </div>);
}

function fv(v: any, sheetName: string, k: number) {
  if (v == null || v === "") return "leer";
  const sh = S.model.sheets.find((x: any) => x.name === sheetName);
  if (sh && k >= 0 && typeof v === "number") return fmtVal(sh, k, v, null);
  return typeof v === "number" ? de(v) : String(v);
}
export function PlanCard({ plan, state, label, onDone }: { plan: any[]; state?: string; label: string; onDone: (s: "done" | "dropped") => void }) {
  const [on, setOn] = useState(() => plan.map(p => !p.error));
  const ok = plan.filter(p => !p.error);
  if (state === "done") return <div className="plan done"><div className="ph"><span>✓ {ok.length} Änderung{ok.length === 1 ? "" : "en"} übernommen</span></div></div>;
  if (state === "dropped") return <div className="plan done"><div className="ph"><span className="hint">Verworfen</span></div></div>;
  const apply = () => {
    const chosen = plan.filter((p, i) => on[i] && !p.error); if (!chosen.length) { onDone("dropped"); return; }
    const touched = ZF.applyPlan(S.model, chosen); S.fresh = {}; S.hits = {};
    touched.forEach((t: any) => { S.fresh[t.sheet] = S.fresh[t.sheet] || {}; S.fresh[t.sheet][t.row] = 1; });
    commit(label); onDone("done"); toast(`${chosen.length} Änderung${chosen.length === 1 ? "" : "en"} übernommen`);
    const f = touched[0]; if (f) goTo(f.sheet, f.row); fadeFresh();
  };
  return (
    <div className="plan">
      <div className="ph"><span>{ok.length} Änderung{ok.length === 1 ? "" : "en"} vorgeschlagen</span>
        <span className="row"><button className="btn small" data-ok onClick={apply}>Übernehmen</button><button className="btn quiet small" onClick={() => onDone("dropped")}>Verwerfen</button></span></div>
      <ul>{plan.map((p, i) => p.error ? <li key={i}><span /><span className="err">{p.error}</span></li> : (
        <li key={i}>
          <input type="checkbox" checked={on[i]} onChange={e => setOn(o => o.map((x, j) => j === i ? e.target.checked : x))} aria-label={p.typ === "neu" ? `Neue Zeile in „${p.sheet}“` : `Zeile ${p.row}`} />
          <div><div className="what">{p.typ === "neu" ? `Neue Zeile in „${p.sheet}“` : p.typ === "loeschen" ? `Zeile ${p.row} löschen` : `Zeile ${p.row} ändern`}</div>
            {p.typ === "loeschen" ? <div className="hint">{p.summary}</div> :
              <dl className="diff">{p.changes.map((c: any, j: number) => c.error ? [<dt key={"t" + j}>{c.name}</dt>, <dd key={"d" + j} className="err">{c.error}</dd>] :
                [<dt key={"t" + j}>{c.name}</dt>, <dd key={"d" + j}>{p.typ === "neu" ? <ins>{fv(c.nv, p.sheet, c.k)}</ins> : <><del>{fv(c.old, p.sheet, c.k)}</del> → <ins>{fv(c.nv, p.sheet, c.k)}</ins></>}</dd>])}</dl>}
          </div>
        </li>))}</ul>
    </div>);
}

export function Bot({ m }: { m: Msg }) {
  if (m.status === "busy") return <div className="bot"><span className="busy">Liest die Tabelle…</span></div>;
  if (m.noAi) return <div className="bot"><p className="hint">Dafür braucht Zettelfrei eine KI. <button type="button" className="linkbtn" onClick={openSettings}>KI einrichten</button></p></div>;
  if (m.status === "error") return <div className="bot"><p className="hint warnc">{m.error}</p></div>;
  const res = m.res, results = m.results || {}; let txt = res.antwort || ""; const groups: any[] = [], hit: Record<string, number[]> = {};
  (res.abfragen || []).forEach((q: any) => { const r = results[q.id]; if (!r) return; txt = txt.split("{" + q.id + "}").join(fmtQuery(r));
    if (r.groups) groups.push(r);
    if (r.rows?.length && q.blatt) { const sh = (S.model.sheets.find((s: any) => s.name.toLowerCase() === String(q.blatt).toLowerCase()) || S.model.sheets[S.cur]).name;
      hit[sh] = [...new Set([...(hit[sh] || []), ...r.rows])]; } });
  const names = Object.keys(hit), count = names.reduce((a, n) => a + hit[n].length, 0);
  const label = `Assistent: „${m.text.length > 50 ? m.text.slice(0, 50) + "…" : m.text}“`;
  return (
    <div className="bot">
      {txt && <div className="answer">{txt}{names.length > 0 && <span className="rows">Berechnet aus {count} Zeile{count === 1 ? "" : "n"} · <button type="button" onClick={() => {
        S.hits = {}; names.forEach(n => { S.hits[n] = Object.fromEntries(hit[n].map(r => [r, 1])) as any; }); S.fresh = {}; goTo(names[0], hit[names[0]][0]); }}>in der Tabelle zeigen</button></span>}</div>}
      {groups.map((r, i) => <GroupCard key={i} r={r} />)}
      {res.rueckfrage && <div className="ask">{res.rueckfrage}</div>}
      {m.plan && m.plan.length > 0 && <PlanCard plan={m.plan} state={m.planState} label={label} onDone={st => { m.planState = st; S.msgs = [...S.msgs]; emit(); }} />}
      {!txt && !res.rueckfrage && !(m.plan && m.plan.length) && !groups.length && <p className="hint">Dazu gab es nichts zu tun. Versuchen Sie es mit einer genaueren Anweisung.</p>}
    </div>);
}

function examples(model: any): string[] {
  if (!model) return [];
  const heads = model.sheets.map((s: any) => s.headers.join(" ")).join(" ").toLowerCase();
  if (/mitarbeiter/.test(heads) && /stunden/.test(heads)) return ["Wie viele Stunden je Baustelle?", "Mehmet heute 7 bis 16 Uhr bei Weber, 30 Min. Pause", "Wer hatte im Oktober die meisten Stunden?", "Was kostet uns Bad Schneider bisher an Lohn?"];
  if (/beginn/.test(heads) && /ende/.test(heads)) return ["Kevin heute 7:00 bis 15:30, 30 Min. Pause, Küche Yilmaz", "Wer hatte diese Woche die meisten Stunden?", "Stunden je Mitarbeiter"];
  if (/lieferant/.test(heads) && /fällig/.test(heads)) return ["Welche Rechnungen sind diese Woche fällig?", "Wie viel haben wir bei Baustoff Kaiser offen?", "Summe je Lieferant"];
  if (/bestand/.test(heads)) return ["Heute 3 Sack Fliesenkleber C2TE verbraucht", "Welche Artikel liegen unter dem Mindestbestand?", "Lagerwert je Lieferant"];
  if (/status/.test(heads) && /kunde/.test(heads)) return ["Becker hat heute bezahlt", "Umsatz pro Monat", "Wie viel ist noch offen?"];
  return ["Wie viele Zeilen hat die Tabelle?", "Welche Zeilen sind unvollständig?"];
}

export function Assistant() {
  const s = useStore(); const [text, setText] = useState(""); const thread = useRef<HTMLDivElement>(null); const ta = useRef<HTMLTextAreaElement>(null);
  const [rec, setRec] = useState<any>(null);
  const ex = useMemo(() => examples(s.model), [s.model]);
  const busy = s.msgs.some(m => m.status === "busy");
  useEffect(() => { const t = thread.current; if (t) t.scrollTop = t.scrollHeight; }, [s.msgs.length, busy]);
  const SR: any = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
  const go = () => { if (!text.trim()) return; send(text); setText(""); };
  const mic = () => {
    if (rec) { rec.stop(); return; }
    const r = new SR(); r.lang = "de-DE"; r.interimResults = true; r.continuous = false; const base = text;
    r.onresult = (ev: any) => { let t = ""; for (let i = 0; i < ev.results.length; i++) t += ev.results[i][0].transcript; setText((base ? base + " " : "") + t); };
    r.onerror = (ev: any) => { if (ev.error === "not-allowed") toast("Mikrofon wurde nicht freigegeben."); };
    r.onend = () => setRec(null); r.start(); setRec(r);
  };
  return (
    <div className="pane chat" id="p-cmd">
      <div className="panehead"><h2>Assistent</h2>
        <button type="button" className={"chip click " + (s.ai.on ? "on" : "off")} id="aiChip" title="KI-Einstellungen" onClick={openSettings}><span className="dot" /><span>{s.ai.label}</span></button></div>
      <div className="thread" ref={thread}>
        {!s.msgs.length && <div className="welcome"><p><b>Sagen Sie einfach, was passiert ist.</b></p>
          <p className="hint">Eintragen, ändern, fragen, auswerten. Summen rechnet die Tabelle selbst, nicht die KI. Jede Änderung sehen Sie vorher als „alt → neu“.</p></div>}
        <div className="msgs" id="cmdOut">
          {s.msgs.map(m => <div className="msg" key={m.id}><div className="me">{m.text}</div><Bot m={m} /></div>)}
        </div>
      </div>
      <div className="composer">
        <div className="examples">{ex.map(t => <button type="button" key={t} onClick={() => { setText(t); ta.current?.focus(); }}>{t}</button>)}</div>
        <div className="cmdwrap">
          <label htmlFor="cmdText" className="vh">Nachricht an den Assistenten</label>
          <textarea id="cmdText" ref={ta} rows={2} value={text} placeholder={s.model ? "Was ist passiert? z. B. „Becker hat heute bezahlt“" : "Zuerst eine Excel-Datei öffnen"} disabled={!s.model}
            onChange={e => setText(e.target.value)} onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing) { e.preventDefault(); go(); } }} />
          {SR && <button type="button" className={"mic" + (rec ? " on" : "")} aria-pressed={!!rec} aria-label="Sprechen" title="Sprechen" onClick={mic}><Mic /></button>}
          {busy ? <button className="btn send" aria-label="Stopp" title="Stopp" onClick={() => ctl?.abort()}><Icon name="stop" size={18} /></button>
            : <button className="btn send" id="cmdBtn" aria-label="Ausführen" title="Ausführen (Enter)" onClick={go} disabled={!s.model}><Icon name="send" size={19} /></button>}
        </div>
      </div>
    </div>);
}
export { mark, setPane };
