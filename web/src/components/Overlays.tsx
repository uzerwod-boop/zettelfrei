// Kleine Dialoge (Bestätigen, Name eingeben) und Kontextmenü – ohne Browser-confirm/prompt, funktioniert auch im Programm.
import { useEffect, useRef, useState, useSyncExternalStore } from "react";

type Ask = { kind: "confirm" | "prompt"; title: string; text?: string; value?: string; ok: string; danger?: boolean; resolve: (v: any) => void } | null;
type Menu = { x: number; y: number; items: MenuItem[] } | null;
export type MenuItem = { label: string; onClick?: () => void; danger?: boolean; disabled?: boolean; hint?: string } | "sep";
let ask: Ask = null, menu: Menu = null; const subs = new Set<() => void>(); let ver = 0;
const ping = () => { ver++; subs.forEach(f => f()); };
const useOv = () => useSyncExternalStore(cb => { subs.add(cb); return () => subs.delete(cb); }, () => ver);

export const confirmBox = (title: string, text = "", ok = "OK", danger = false) => new Promise<boolean>(resolve => { ask = { kind: "confirm", title, text, ok, danger, resolve }; ping(); });
export const promptBox = (title: string, value = "", ok = "Übernehmen", text = "") => new Promise<string | null>(resolve => { ask = { kind: "prompt", title, value, ok, text, resolve }; ping(); });
export const openMenu = (e: { clientX: number; clientY: number; preventDefault?: () => void }, items: MenuItem[]) => { e.preventDefault?.(); menu = { x: e.clientX, y: e.clientY, items }; ping(); };
const closeMenu = () => { if (menu) { menu = null; ping(); } };

export function Overlays() {
  useOv();
  const dlg = useRef<HTMLDialogElement>(null), inp = useRef<HTMLInputElement>(null), mref = useRef<HTMLDivElement>(null);
  const [val, setVal] = useState("");
  useEffect(() => { const d = dlg.current; if (!d) return;
    if (ask && !d.open) { setVal(ask.value || ""); d.showModal(); setTimeout(() => { inp.current?.focus(); inp.current?.select(); }, 20); }
    if (!ask && d.open) d.close(); });
  useEffect(() => { if (!menu) return; const off = (e: Event) => { if (mref.current && e.target instanceof Node && mref.current.contains(e.target)) return; closeMenu(); };
    const esc = (e: KeyboardEvent) => { if (e.key === "Escape") closeMenu(); };
    window.addEventListener("mousedown", off, true); window.addEventListener("keydown", esc); window.addEventListener("blur", closeMenu); window.addEventListener("resize", closeMenu);
    return () => { window.removeEventListener("mousedown", off, true); window.removeEventListener("keydown", esc); window.removeEventListener("blur", closeMenu); window.removeEventListener("resize", closeMenu); }; });
  const finish = (v: any) => { const a = ask; ask = null; ping(); a?.resolve(v); };
  const mx = menu ? Math.min(menu.x, window.innerWidth - 250) : 0, my = menu ? Math.min(menu.y, window.innerHeight - 30 - menu.items.length * 34) : 0;
  return (<>
    <dialog ref={dlg} className="askdlg" onCancel={e => { e.preventDefault(); finish(ask?.kind === "prompt" ? null : false); }}>
      {ask && <form method="dialog" onSubmit={e => { e.preventDefault(); finish(ask!.kind === "prompt" ? val : true); }}>
        <h2>{ask.title}</h2>{ask.text && <p className="note">{ask.text}</p>}
        {ask.kind === "prompt" && <input ref={inp} type="text" value={val} onChange={e => setVal(e.target.value)} />}
        <div className="row end"><button className="btn quiet" type="button" onClick={() => finish(ask!.kind === "prompt" ? null : false)}>Abbrechen</button>
          <button className={"btn" + (ask.danger ? " danger" : "")} type="submit" autoFocus={ask.kind === "confirm"}>{ask.ok}</button></div>
      </form>}
    </dialog>
    {menu && <div ref={mref} className="ctxmenu" role="menu" style={{ left: mx, top: Math.max(8, my) }}>
      {menu.items.map((it, i) => it === "sep" ? <div key={i} className="sep" /> :
        <button key={i} type="button" role="menuitem" className={it.danger ? "danger" : ""} disabled={it.disabled} onClick={() => { closeMenu(); it.onClick?.(); }}>
          <span>{it.label}</span>{it.hint && <kbd>{it.hint}</kbd>}</button>)}
    </div>}
  </>);
}
