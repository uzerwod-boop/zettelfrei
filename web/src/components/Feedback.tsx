// Feedback aus dem Pilotbetrieb: kurzer Text + Version, System und KI-Modell. Keine Tabellendaten.
import { useEffect, useRef, useState, useSyncExternalStore } from "react";
import { FEEDBACK_EMAIL, REPO } from "../config";
import { S } from "../store";
import { IS_APP, openExternal } from "../util";

let open = false; const subs = new Set<() => void>();
export function openFeedback() { open = true; subs.forEach(f => f()); }
function close() { open = false; subs.forEach(f => f()); }

function sysInfo() {
  const ua = navigator.userAgent, os = /Mac OS X/.test(ua) ? "macOS" : /Windows/.test(ua) ? "Windows" : /Android/.test(ua) ? "Android" : /iPhone|iPad/.test(ua) ? "iOS" : /Linux/.test(ua) ? "Linux" : "anderes System";
  return `Zettelfrei ${__APP_VERSION__} · ${IS_APP ? "Programm" : "Browser"} · ${os} · KI: ${AI.status().label}`;
}
export function Feedback() {
  const isOpen = useSyncExternalStore(cb => { subs.add(cb); return () => subs.delete(cb); }, () => open);
  const dlg = useRef<HTMLDialogElement>(null);
  const [kind, setKind] = useState("Problem"), [text, setText] = useState(""), [withQ, setWithQ] = useState(false);
  useEffect(() => { const d = dlg.current; if (!d) return; if (isOpen && !d.open) { setText(""); setWithQ(false); d.showModal(); } if (!isOpen && d.open) d.close(); }, [isOpen]);
  const lastQ = [...S.msgs].reverse().find(m => m.text)?.text || "";
  const body = () => `${text.trim()}\n\n—\n${sysInfo()}${withQ && lastQ ? `\nLetzte Frage an den Assistenten: ${lastQ}` : ""}`;
  const subject = () => `Zettelfrei-Feedback: ${kind}${text.trim() ? " – " + text.trim().split("\n")[0].slice(0, 60) : ""}`;
  const mail = () => { openExternal(`mailto:${FEEDBACK_EMAIL}?subject=${encodeURIComponent(subject())}&body=${encodeURIComponent(body())}`); close(); };
  const github = () => { openExternal(`${REPO}/issues/new?title=${encodeURIComponent(subject())}&body=${encodeURIComponent(body())}&labels=${encodeURIComponent(kind === "Problem" ? "fehler" : "idee")}`); close(); };
  return (
    <dialog ref={dlg} id="feedback" onClose={close}>
      <form method="dialog" onSubmit={e => e.preventDefault()}>
        <h2>Feedback</h2>
        <p className="note">Was hat nicht funktioniert, was fehlt Ihnen? Jede Rückmeldung hilft. Ihre Tabelle wird nicht mitgeschickt.</p>
        <div className="row seg">{["Problem", "Idee", "Lob"].map(k => <button key={k} type="button" className={"btn quiet small" + (kind === k ? " on" : "")} onClick={() => setKind(k)}>{k}</button>)}</div>
        <label>Ihre Nachricht<textarea id="fbText" rows={6} value={text} onChange={e => setText(e.target.value)} placeholder={kind === "Problem" ? "Was haben Sie gemacht, was ist passiert, was hätten Sie erwartet?" : kind === "Idee" ? "Was würde Ihnen im Alltag Zeit sparen?" : "Was gefällt Ihnen?"} /></label>
        {lastQ && <label className="check"><input type="checkbox" checked={withQ} onChange={e => setWithQ(e.target.checked)} /> Letzte Frage an den Assistenten mitschicken („{lastQ.length > 60 ? lastQ.slice(0, 60) + "…" : lastQ}“)</label>}
        <p className="note">Wird mitgeschickt: {sysInfo()}</p>
        <div className="row">
          {FEEDBACK_EMAIL ? <button className="btn" type="button" disabled={!text.trim()} onClick={mail}>Per E-Mail senden</button> : null}
          <button className={FEEDBACK_EMAIL ? "btn quiet" : "btn"} type="button" disabled={!text.trim()} onClick={github}>{FEEDBACK_EMAIL ? "Über GitHub" : "Senden (GitHub-Formular)"}</button>
          <button className="btn quiet" type="button" onClick={close}>Abbrechen</button>
        </div>
      </form>
    </dialog>);
}
