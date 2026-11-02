// „Über Zettelfrei“: Version, Neuigkeiten, Updates, Anleitung, Lizenz.
import { useEffect, useRef, useSyncExternalStore } from "react";
import { GUIDE_URL, REPO } from "../config";
import { toast } from "../store";
import { IS_APP, openExternal } from "../util";
import { Mark } from "./Icons";
import { checkForUpdate } from "./UpdateBar";

let open = false; const subs = new Set<() => void>();
export function openAbout() { open = true; subs.forEach(f => f()); }
function close() { open = false; subs.forEach(f => f()); }
const os = () => { const ua = navigator.userAgent; return /Mac OS X/.test(ua) ? "macOS" : /Windows/.test(ua) ? "Windows" : /Android/.test(ua) ? "Android" : /iPhone|iPad/.test(ua) ? "iOS" : /Linux/.test(ua) ? "Linux" : "–"; };
const date = () => new Date(__BUILD_DATE__ + "T12:00:00").toLocaleDateString("de-DE", { day: "2-digit", month: "long", year: "numeric" });

export function About() {
  const isOpen = useSyncExternalStore(cb => { subs.add(cb); return () => subs.delete(cb); }, () => open);
  const dlg = useRef<HTMLDialogElement>(null);
  useEffect(() => { const d = dlg.current; if (!d) return; if (isOpen && !d.open) d.showModal(); if (!isOpen && d.open) d.close(); }, [isOpen]);
  const info = `Zettelfrei ${__APP_VERSION__} (${__BUILD_DATE__}) · ${IS_APP ? "Programm" : "Browser"} · ${os()} · KI: ${AI.status().label}`;
  return (
    <dialog ref={dlg} id="about" onClose={close}>
      <div className="about">
        <Mark size={54} />
        <h2>Zettel<em>frei</em></h2>
        <p className="ver"><b>Version {__APP_VERSION__}</b> · {date()}</p>
        <p className="note">{IS_APP ? `Programm für ${os()}` : "Browser-Version"} · KI: {AI.status().label}</p>
        <div className="row">
          <button className="btn" type="button" onClick={() => openExternal(`${REPO}/releases/tag/v${__APP_VERSION__}`)}>Was ist neu?</button>
          {IS_APP && <button className="btn quiet" type="button" onClick={() => { close(); checkForUpdate(true); }}>Nach Updates suchen</button>}
          <button className="btn quiet" type="button" onClick={() => openExternal(GUIDE_URL)}>Anleitung</button>
        </div>
        <p className="note small">Kostenlos und Open Source · MIT oder Apache-2.0 · © 2026 Bohdan Sychevskyi<br />
          <button type="button" className="linkbtn" onClick={() => openExternal(REPO)}>Quellcode auf GitHub</button> · <button type="button" className="linkbtn" onClick={() => openExternal(`${REPO}/blob/main/CHANGELOG.md`)}>Alle Änderungen</button> ·{" "}
          <button type="button" className="linkbtn" onClick={() => navigator.clipboard?.writeText(info).then(() => toast("Versionsinfo kopiert"), () => toast(info))}>Versionsinfo kopieren</button></p>
        <div className="row end"><button className="btn quiet" type="button" onClick={close}>Schließen</button></div>
      </div>
    </dialog>);
}
