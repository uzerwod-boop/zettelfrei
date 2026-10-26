import { useEffect, useMemo, useRef, useState } from "react";
import { S, commit, emit, goTo, issues, restoreVersion, saveFile, setPane, toast, useStore } from "../store";
import { TODAY, b64, de, firma } from "../util";
import { Bot, PlanCard, RULES, aiError, context } from "./Assistant";
import { openSettings } from "./Settings";
import { DIALOGS } from "../legacyDialogs";

/* ---------- Prüfen ---------- */
export function Pruefen() {
  const s = useStore(); const all = useMemo(() => issues(), [s.rev]); const [busy, setBusy] = useState(false);
  const aiCheck = () => {
    if (!AI.active()) { openSettings(); return; } setBusy(true);
    const prompt = context() + "\n\nPrüfe die Datei fachlich wie eine erfahrene Bürokraft im Handwerk. Finde höchstens 12 echte Auffälligkeiten, die Geld, Termine oder Ordnung betreffen: z. B. Bestand unter Mindestbestand, fertige Aufträge ohne Rechnung, Rechnungen lange unbezahlt, unplausible Stunden, Tippfehler in Namen, fehlende Zuordnungen, falscher MwSt-Satz, Bauleistungen an andere Baufirmen, bei denen § 13b UStG (Steuerschuldnerschaft des Leistungsempfängers) zu prüfen wäre. Formuliere solche Punkte als Hinweis zum Prüfen, nicht als Rechtsauskunft. Keine Kleinigkeiten zur Formatierung.\nAntworte NUR mit JSON: {\"probleme\":[{\"blatt\":\"…\",\"zeile\":Excel-Zeilennummer,\"spalte\":\"Spaltenname\",\"problem\":\"kurz, deutsch\"}]}";
    AI.json(prompt).then((res: any) => {
      S.aiIssues = ((res && res.probleme) || []).filter((p: any) => p && p.blatt && p.zeile).map((p: any) => ({ blatt: p.blatt, zeile: +p.zeile, spalte: p.spalte, problem: String(p.problem || ""), ki: 1 }));
      emit(); toast(S.aiIssues.length ? S.aiIssues.length + " fachliche Hinweise gefunden" : "Fachlich nichts Auffälliges gefunden");
    }, (e: any) => { const m = aiError(e); if (m) toast(m); }).finally(() => setBusy(false));
  };
  return (
    <div className="pane" id="p-check">
      <div className="panehead"><h2>Was stimmt nicht?</h2>{all.length > 0 && <span className="count">{all.length}</span>}</div>
      <p className="hint">Leere Pflichtfelder, doppelte Nummern, Zahlen als Text, Arbeitszeitgesetz und Formelfehler findet die Tabelle selbst. Fachliches prüft die KI.</p>
      <div className="row"><button className="btn ghost" id="aiCheckBtn" disabled={busy || !s.model} onClick={aiCheck}>Fachlich prüfen lassen</button>{busy && <span className="busy">KI prüft die Datei…</span>}</div>
      {!s.model ? <p className="hint">Zuerst eine Excel-Datei öffnen.</p> :
        <ul className="issues" id="issues">{all.length ? all.map((x: any, i: number) => {
          const sh = s.model.sheets.find((q: any) => q.name === x.blatt), k = x.k != null ? x.k : (sh ? ZF.findCol(sh, x.spalte) : -1);
          return <li key={i}><button type="button" onClick={() => goTo(x.blatt, x.zeile, k >= 0 ? k : null)}>
            <span className="where">{x.blatt} · {sh && k >= 0 ? ZF.colLetter(sh.c0 + k) + x.zeile : "Z. " + x.zeile}</span>
            <span>{x.problem}{x.ki ? <span className="ki">KI</span> : null}</span></button></li>; })
          : <li className="hint" style={{ padding: "8px 0" }}>Keine Auffälligkeiten gefunden.</li>}</ul>}
    </div>);
}

/* ---------- Verlauf und Änderungsprotokoll ---------- */
export function Verlauf() {
  const s = useStore(); const hm = (t: number) => new Date(t).toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" });
  return (
    <div className="pane" id="p-hist">
      <div className="panehead"><h2>Verlauf</h2></div>
      <p className="hint">Jede Änderung ist ein Stand. Ein Klick holt einen früheren zurück.</p>
      <ul className="histlist" id="hist">{s.history.map((h, i) => ({ h, i })).reverse().map(({ h, i }, j) => (
        <li key={i}><time>{hm(h.t)}</time><span>{h.label}</span>{j === 0 ? <span className="hint">aktuell</span> : <button className="btn quiet small" onClick={() => restoreVersion(i)}>Zurückholen</button>}</li>))}</ul>
      <h2 className="sub">Änderungsprotokoll</h2>
      <p className="hint">Wer hat wann was geändert, Zelle für Zelle. Das Protokoll wird nie gelöscht, auch nicht beim Zurückholen. {s.protokoll.length} Einträge</p>
      <label className="check"><input type="checkbox" id="protSheet" checked={s.protSheet} onChange={e => { S.protSheet = e.target.checked; emit(); }} /> Beim Speichern als eigenes Blatt „Änderungsprotokoll“ anhängen</label>
      <div className="row"><button className="btn quiet small" type="button" onClick={() => saveFile("Aenderungsprotokoll.csv", ZFD.csv([["Zeitpunkt", "Bearbeiter", "Aktion", "Details"]].concat(S.protokoll.map(e => [e.zeit, e.wer, e.aktion, (e.details || []).join("; ")]))), "text/csv")}>Protokoll als CSV</button></div>
      <ul className="histlist prot" id="protList">{s.protokoll.slice(-30).reverse().map((e, i) => (
        <li key={i}><time>{e.zeit.slice(0, 16)}</time><span>{e.aktion}{e.details?.length ? <small>{e.details.slice(0, 4).join(" · ")}{e.details.length > 4 ? " …" : ""}</small> : null}</span>{e.wer ? <span className="hint">{e.wer}</span> : <span />}</li>))}</ul>
    </div>);
}

/* ---------- Belege: E-Rechnung exakt, sonst KI ---------- */
const SAMPLE_INVOICE = "Fliesen-Großhandel Weber KG · Industriestr. 14 · 04129 Leipzig\nRECHNUNG Nr. 88213        Datum: 15.10.2026\nKunde: Fliesen Krause, Kommission: Küche Yilmaz\nPos  Artikel                                Menge   Einh.   Einzelpreis   Gesamt\n1    Feinsteinzeug 60x60 anthrazit          18,5    m²      32,90 €       608,65 €\n2    Abschlussprofil Alu 2,5 m              8       Stk     9,80 €        78,40 €\n3    Flexkleber S1 25 kg                    6       Sack    17,50 €       105,00 €\n4    Fugenmörtel weiß 5 kg                  4       Eimer   11,40 €       45,60 €\nNetto 837,65 €   MwSt 19 % 159,15 €   Brutto 996,80 €\nZahlbar bis 29.10.2026 ohne Abzug.";
type BelegState = { busy?: string; inv?: any; plan?: any[]; planState?: "open" | "done" | "dropped"; label?: string; msg?: any; error?: string; info?: string };
export function Belege() {
  const s = useStore(); const [st, setSt] = useState<BelegState>({}); const [target, setTarget] = useState<number>(-1); const [over, setOver] = useState(false);
  const inp = useRef<HTMLInputElement>(null);
  const tables = (s.model?.sheets || []).map((x: any, i: number) => ({ x, i })).filter((o: any) => o.x.hr >= 0);
  const tgt = () => S.model.sheets[target >= 0 && S.model.sheets[target] ? target : S.cur];
  useEffect(() => { setTarget(-1); }, [s.model]);
  const showInv = (inv: any) => {
    const t = tgt(), plan = ZF.planOps(S.model, ZFE.toOps(ZF, S.model, t, inv)).filter((p: any) => p.error || p.typ !== "neu" || p.changes.length);
    setSt({ inv, plan, planState: plan.length ? "open" : undefined, label: `E-Rechnung übernommen: ${inv.seller.name} ${inv.number}`,
      info: plan.length ? undefined : `Das Blatt „${t.name}“ passt nicht zu einer Rechnung. Wählen Sie oben ein Lager- oder Rechnungsblatt.` });
  };
  const readXml = (text: string) => { try { showInv(ZFE.read(text)); } catch (e: any) { setSt({ error: `Diese XML-Datei ist keine lesbare E-Rechnung (${e.message}).` }); } };
  const extract = (text: string, images: Blob[], name: string) => {
    if (!AI.active()) { setSt({ error: "", msg: { noAi: true } }); return; }
    const t = tgt(); setSt({ busy: "KI liest das Dokument…" });
    const prompt = context() + `\n\nDer Nutzer hat ein Dokument hochgeladen (${name})${images.length ? ", siehe Bild" : ""}. Übertrage den Inhalt in das Blatt „${t.name}“.\n` +
      "Überlege, was das Dokument für dieses Blatt bedeutet: z. B. bei einer Lieferantenrechnung im Lager den Bestand passender vorhandener Artikel per {\"plus\": Menge} erhöhen und unbekannte Artikel neu anlegen; bei einer Kundenbestellung einen neuen Auftrag anlegen; bei einem Stundenzettel neue Zeilen pro Person und Tag. Ordne Kommission oder Baustelle einem vorhandenen Projekt zu, wenn eindeutig.\n" +
      "In \"antwort\" in einem Satz zusammenfassen, was im Dokument steht und was du überträgst. Unleserliches nicht raten, sondern in rueckfrage nennen.\n" + (text ? `Text des Dokuments:\n"""${text}"""\n\n` : "") + RULES;
    AI.json(prompt, { images }).then((res: any) => {
      const results = ZF.runQueries(S.model, res.abfragen || []), plan = ZF.planOps(S.model, res.aenderungen || []);
      setSt({ msg: { id: 0, text: name, status: "done", res, results, plan, planState: plan.length ? "open" : undefined, label: "Dokument übernommen: " + name } });
    }, (e: any) => setSt({ error: aiError(e) }));
  };
  const handle = async (f: File) => {
    if (!S.model) { toast("Zuerst eine Excel-Datei öffnen."); return; }
    setSt({});
    if (/\.xml$/i.test(f.name) || /xml/.test(f.type)) { readXml(await f.text()); return; }
    if (f.type.startsWith("image/")) { extract("", [f], f.name); return; }
    if (!/pdf$/i.test(f.type) && !/\.pdf$/i.test(f.name)) { toast("Bitte ein PDF, Foto oder eine E-Rechnung (XML) wählen."); return; }
    setSt({ busy: "PDF wird geöffnet…" });
    try {
      const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await f.arrayBuffer()), isEvalSupported: false }).promise;
      let att = null; try { att = await ZFE.fromPdf(pdf); } catch { /* keine E-Rechnung */ }
      if (att) { readXml(att.xml); return; }
      let text = "";
      for (let p = 1; p <= Math.min(pdf.numPages, 6); p++) { const pg = await pdf.getPage(p), tc = await pg.getTextContent(); let lastY: number | null = null;
        tc.items.forEach((it: any) => { const y = it.transform[5]; text += (lastY !== null && Math.abs(y - lastY) > 3) ? "\n" : (lastY !== null ? " " : ""); text += it.str; lastY = y; }); text += "\n\n"; }
      if (text.replace(/\s/g, "").length > 40) { extract(text.slice(0, 30000), [], f.name); return; }
      setSt({ busy: "Gescanntes PDF, Seiten werden als Bild gelesen…" }); const imgs: Blob[] = [];
      for (let p = 1; p <= Math.min(pdf.numPages, 3); p++) { const pg = await pdf.getPage(p), vp = pg.getViewport({ scale: 1.6 }), c = document.createElement("canvas");
        c.width = vp.width; c.height = vp.height; await pg.render({ canvasContext: c.getContext("2d"), viewport: vp }).promise; imgs.push(await new Promise<Blob>(r => c.toBlob(b => r(b!), "image/jpeg", 0.85))); }
      extract("", imgs, f.name);
    } catch { setSt({ error: "Diese Datei lässt sich nicht öffnen. Ist es ein gültiges PDF?" }); }
  };
  useEffect(() => { const h = (e: any) => handle(e.detail); window.addEventListener("zf-beleg", h); return () => window.removeEventListener("zf-beleg", h); });
  const docs = window.ZF_DOCS || {};
  const sample = (n: string) => handle(new File([b64(docs[n])], n, { type: /\.pdf$/i.test(n) ? "application/pdf" : "application/xml" }));
  const inv = st.inv;
  return (
    <div className="pane" id="p-pdf">
      <div className="panehead"><h2>Belege übernehmen</h2></div>
      <p className="hint">E-Rechnungen (XRechnung, ZUGFeRD) werden exakt und ohne KI gelesen. Andere PDFs, Scans und Fotos liest die KI. Sie sehen vorher, welche Zeilen sich ändern.</p>
      <label className="fl">In welches Blatt?
        <select id="pdfTarget" value={target >= 0 ? target : s.cur} onChange={e => setTarget(+e.target.value)} disabled={!s.model}>
          {tables.length ? tables.map((o: any) => <option key={o.i} value={o.i}>{o.x.name}</option>) : <option value={0}>Zuerst eine Excel-Datei öffnen</option>}</select></label>
      <label className={"drop" + (over ? " over" : "")} htmlFor="pdfInput"
        onDragOver={e => { e.preventDefault(); e.stopPropagation(); setOver(true); }} onDragLeave={() => setOver(false)}
        onDrop={e => { e.preventDefault(); e.stopPropagation(); setOver(false); const f = e.dataTransfer.files[0]; if (f) handle(f); }}>
        <strong>Datei hierher ziehen oder antippen</strong>E-Rechnung (XML, ZUGFeRD-PDF), PDF, JPG oder PNG</label>
      <input ref={inp} type="file" id="pdfInput" hidden accept="application/pdf,image/png,image/jpeg,image/webp,.xml,application/xml,text/xml" onChange={e => { const f = e.target.files?.[0]; e.target.value = ""; if (f) handle(f); }} />
      <div className="row">
        {Object.keys(docs).map(n => <button key={n} className="btn quiet small" type="button" data-sampledoc={n} onClick={() => sample(n)}>Beispiel: {/\.pdf$/i.test(n) ? "ZUGFeRD-PDF" : "XRechnung"}</button>)}
        <button className="btn quiet small" type="button" onClick={() => extract(SAMPLE_INVOICE, [], "Beispiel-Rechnung Weber KG")}>Beispiel: Rechnungstext (KI)</button></div>
      {st.busy && <span className="busy">{st.busy}</span>}
      <div id="pdfOut" className="stack">
        {st.error && <p className="hint warnc">{st.error}</p>}
        {inv && <div className="einv">
          <div className="einv-h"><span className="badge">E-Rechnung</span><span>{inv.format} · exakt gelesen, ohne KI</span></div>
          <dl className="kv"><dt>{inv.type}</dt><dd>Nr. {inv.number} vom {inv.date}</dd><dt>Von</dt><dd>{inv.seller.name}</dd>
            {(inv.buyerRef || inv.project) && <><dt>Kommission</dt><dd>{inv.project || inv.buyerRef}</dd></>}
            <dt>Betrag</dt><dd>{de(inv.net, 2)} € netto · {de(inv.gross, 2)} € brutto</dd>{inv.due && <><dt>Fällig</dt><dd>{inv.due}</dd></>}</dl>
          <ul className="einv-lines">{inv.lines.map((l: any, i: number) => <li key={i}><span>{l.name}</span><span>{de(l.qty, l.qty % 1 ? 2 : 0)} {l.unit} × {de(l.price, 2)} €</span></li>)}</ul>
        </div>}
        {st.info && <p className="hint">{st.info}</p>}
        {inv && st.plan && st.plan.length > 0 && <PlanCard plan={st.plan} state={st.planState} label={st.label!} onDone={ps => setSt(o => ({ ...o, planState: ps }))} />}
        {st.msg && (st.msg.noAi ? <p className="hint">Dafür braucht Zettelfrei eine KI. <button type="button" className="linkbtn" onClick={openSettings}>KI einrichten</button></p> : <Bot m={st.msg} />)}
      </div>
    </div>);
}

/* ---------- Dokumente: bewährtes Modul src/workflows.js, eingebettet ---------- */
let docsLoaded = false, docsKey = "";
const dkey = () => S.protokoll.length + ":" + S.cur + ":" + (S.model?.name || "");
const docsRender = () => { const D = (window as any).ZFDOC; if (D && S.model) { D.render(); docsKey = dkey(); } };
export function Dokumente({ active }: { active: boolean }) {
  const s = useStore();
  useEffect(() => {
    (window as any).ZFAPP = { model: () => S.model, commit: (l: string) => { commit(l); docsRender(); }, saveFile, toast, firma, today: () => TODAY,
      esc: (x: any) => String(x == null ? "" : x).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;" } as any)[c]), de,
      selectTab: (id: string) => setPane((id.replace("p-", "") as any)), goTo, renderAll: () => { emit(); docsRender(); },
      mark: (rows: number[], sh: string) => { S.fresh = { [sh]: Object.fromEntries(rows.map(r => [r, 1])) } as any; } };
    if (!docsLoaded) { docsLoaded = true; const sc = document.createElement("script"); sc.src = "./lib/workflows.js"; sc.onload = docsRender; document.body.appendChild(sc); }
  }, []);
  useEffect(() => { if (active && docsKey !== dkey()) docsRender(); }, [active, s.rev]);
  return (
    <div className="pane" id="p-docs" hidden={!active}>
      <div className="panehead"><h2>Dokumente</h2></div>
      {!s.model && <p className="hint">Zuerst eine Excel-Datei öffnen.</p>}
      <div id="docs" className="docs" />
      <div dangerouslySetInnerHTML={{ __html: DIALOGS }} />
    </div>);
}
