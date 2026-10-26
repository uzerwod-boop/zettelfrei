import { useEffect, useRef, useState, useSyncExternalStore } from "react";
import { toast } from "../store";
import { aiError } from "./Assistant";

// Eigener kleiner Zustand: ist der Dialog offen?
let open = false; const subs = new Set<() => void>();
export function openSettings() { open = true; subs.forEach(f => f()); }
function close() { open = false; subs.forEach(f => f()); }

type Prov = "auto" | "server" | "openrouter" | "claude" | "local" | "off";
const HELP = (p: Prov, server: boolean) => ({
  auto: `Nutzt den eigenen Server, wenn Zettelfrei mit „npm start“ läuft${server ? " (gefunden)" : " (nicht gefunden)"}, sonst OpenRouter mit dem Schlüssel unten.`,
  server: `Schlüssel liegt nur auf dem Server: OPENROUTER_API_KEY=… npm start oder ANTHROPIC_API_KEY=… npm start. Am besten für Betriebe mit mehreren Rechnern.${server ? " Server gefunden." : " Kein Server gefunden."}`,
  openrouter: "Ein Schlüssel für viele Modelle: Claude, GPT, Gemini, Mistral. Schlüssel auf openrouter.ai/keys anlegen und mit Guthaben aufladen. Der Schlüssel bleibt in diesem Browser.",
  claude: "Schlüssel aus console.anthropic.com. Er wird nur in diesem Browser gespeichert und direkt an Anthropic geschickt. Nur für den eigenen Rechner.",
  local: "Daten verlassen den Rechner nicht. Für Ollama: OLLAMA_ORIGINS=* ollama serve, Adresse http://localhost:11434/v1. Für Fotos und Scans ein Modell mit Bildverständnis wählen.",
  off: "Tabelle öffnen, bearbeiten, prüfen und speichern funktioniert ohne KI." }[p]);

export function Settings() {
  const isOpen = useSyncExternalStore(cb => { subs.add(cb); return () => subs.delete(cb); }, () => open);
  const dlg = useRef<HTMLDialogElement>(null);
  const [prov, setProv] = useState<Prov>("auto"), [url, setUrl] = useState(""), [keys, setKeys] = useState<any>({}), [models, setModels] = useState<any>({});
  const [stat, setStat] = useState<{ cls: string; text: string } | null>(null);
  useEffect(() => {
    const d = dlg.current; if (!d) return;
    if (isOpen && !d.open) { const c = AI.status().cfg; setProv(c.provider || "auto"); setUrl(c.url || "http://localhost:11434/v1"); setKeys({ ...c.keys }); setModels({ ...c.models }); setStat(null); d.showModal(); }
    if (!isOpen && d.open) d.close();
  }, [isOpen]);
  const st = AI.status(), kp = prov === "auto" ? (st.server ? null : "openrouter") : ["claude", "openrouter", "local"].includes(prov) ? prov : null;
  const cfg = () => ({ provider: prov, url: url.trim(), keys, models });
  const test = () => { AI.save(cfg()); setStat({ cls: "", text: "Teste die Verbindung…" });
    AI.test().then((ok: boolean) => setStat({ cls: ok ? "ok" : "bad", text: ok ? "Verbindung steht. Die KI antwortet." : "Die KI antwortet, aber nicht im erwarteten Format. Anderes Modell versuchen." }),
      (e: any) => setStat({ cls: "bad", text: aiError(e) || "Abgebrochen." })); };
  return (
    <dialog ref={dlg} id="settings" onClose={close}>
      <form method="dialog" onSubmit={e => { const v = (e.nativeEvent as any).submitter?.value; if (v === "save") { AI.save(cfg()); toast("KI-Einstellungen gespeichert"); } }}>
        <h2>KI-Einstellungen</h2>
        <p className="note">Ohne KI funktioniert alles außer Assistent, Auslesen von PDFs ohne E-Rechnung und fachlicher Prüfung. Einstellungen bleiben nur in diesem Browser.</p>
        <label>Anbieter
          <select id="setProvider" value={prov} onChange={e => setProv(e.target.value as Prov)}>
            <option value="auto">Automatisch</option><option value="server">Eigener Server (npm start mit API-Schlüssel)</option>
            <option value="openrouter">OpenRouter, ein Schlüssel für viele Modelle</option><option value="claude">Claude API, Schlüssel in diesem Browser</option>
            <option value="local">Lokales Modell (Ollama, LM Studio; OpenAI-kompatibel)</option><option value="off">Aus</option>
          </select></label>
        {kp && <label>{kp === "local" ? "API-Schlüssel (meist leer lassen)" : kp === "openrouter" ? "OpenRouter-Schlüssel" : "Claude-API-Schlüssel"}
          <input type="password" id="setKey" autoComplete="off" value={keys[kp] || ""} placeholder={kp === "openrouter" ? "sk-or-…" : kp === "claude" ? "sk-ant-…" : "optional"}
            onChange={e => setKeys({ ...keys, [kp]: e.target.value.trim() })} /></label>}
        {prov === "local" && <label>Adresse des lokalen Modells<input type="text" value={url} onChange={e => setUrl(e.target.value)} placeholder="http://localhost:11434/v1" /></label>}
        {kp && <label>Modell (leer = Standard)<input type="text" id="setModel" value={models[kp] || ""} placeholder={st.defaults[kp]} onChange={e => setModels({ ...models, [kp]: e.target.value.trim() })} /></label>}
        <p className="note">{HELP(prov, !!st.server)}</p>
        {stat && <div className={"status " + stat.cls}>{stat.text}</div>}
        <div className="row"><button className="btn" value="save" id="setSave">Speichern</button><button className="btn quiet" type="button" onClick={test}>Verbindung testen</button><button className="btn quiet" value="cancel" formNoValidate>Abbrechen</button></div>
      </form>
    </dialog>);
}
