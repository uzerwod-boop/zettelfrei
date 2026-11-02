// Desktop-Programm: neue Version aus den GitHub-Releases finden, laden, installieren, neu starten.
// Die Update-Dateien sind signiert; das Programm installiert nur, was mit unserem Schlüssel signiert ist.
import { useEffect, useState, useSyncExternalStore } from "react";
import { IS_APP } from "../util";

type St = { s: "idle" | "checking" | "none" | "available" | "loading" | "error"; version?: string; notes?: string; pct?: number; msg?: string };
let st: St = { s: "idle" }, upd: any = null, dismissed = false; const subs = new Set<() => void>();
const set = (n: St) => { st = n; subs.forEach(f => f()); };
const T = () => (window as any).__TAURI__;

export async function checkForUpdate(manual = false) {
  if (!IS_APP || !T()?.updater) { if (manual) set({ s: "error", msg: "Updates gibt es nur im installierten Programm." }); return; }
  set({ s: "checking" });
  try { upd = await T().updater.check(); dismissed = false;
    if (upd) set({ s: "available", version: upd.version, notes: upd.body || "" }); else set({ s: manual ? "none" : "idle" }); }
  catch (e: any) { set(manual ? { s: "error", msg: "Keine Verbindung zum Update-Server. Später noch einmal versuchen." } : { s: "idle" }); }
}
async function install() {
  if (!upd) return; let total = 0, got = 0; set({ ...st, s: "loading", pct: 0 });
  try { await upd.downloadAndInstall((ev: any) => {
      if (ev.event === "Started") total = ev.data?.contentLength || 0;
      if (ev.event === "Progress") { got += ev.data?.chunkLength || 0; if (total) set({ ...st, s: "loading", pct: Math.min(99, Math.round(got / total * 100)) }); } });
    await T().process.relaunch(); }
  catch { set({ ...st, s: "error", msg: "Update fehlgeschlagen. Bitte die neue Version von GitHub herunterladen." }); }
}
export const useUpdate = () => useSyncExternalStore(cb => { subs.add(cb); return () => subs.delete(cb); }, () => st);

export function UpdateBar() {
  const u = useUpdate(), [, force] = useState(0);
  useEffect(() => { const t = setTimeout(() => checkForUpdate(false), 4000); return () => clearTimeout(t); }, []);
  if (dismissed || !["available", "loading", "error", "none"].includes(u.s) || (u.s === "error" && !u.msg)) return null;
  return (
    <div className="updatebar" role="status">
      {u.s === "available" && <><span><b>Zettelfrei {u.version} ist da.</b> {u.notes ? u.notes.split("\n").find(l => l.trim() && !l.startsWith("#"))?.replace(/[*_`]/g, "").slice(0, 110) : "Neue Version mit Verbesserungen."}</span>
        <button className="btn small" type="button" onClick={install}>Jetzt aktualisieren</button>
        <button className="btn quiet small" type="button" onClick={() => { dismissed = true; force(n => n + 1); }}>Später</button></>}
      {u.s === "loading" && <span><b>Update wird geladen …</b> {u.pct}% – Zettelfrei startet danach neu, Ihre Datei bleibt erhalten.</span>}
      {u.s === "none" && <><span>Sie haben die neueste Version.</span><button className="btn quiet small" type="button" onClick={() => set({ s: "idle" })}>OK</button></>}
      {u.s === "error" && <><span>{u.msg}</span><button className="btn quiet small" type="button" onClick={() => set({ s: "idle" })}>OK</button></>}
    </div>);
}
