export const NOW = new Date();
export const TODAY = `${String(NOW.getDate()).padStart(2, "0")}.${String(NOW.getMonth() + 1).padStart(2, "0")}.${NOW.getFullYear()}`;
export const WEEKDAY = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"][NOW.getDay()];
export function de(n: number, d?: number | null) { return Number(n).toLocaleString("de-DE", { minimumFractionDigits: d || 0, maximumFractionDigits: d == null ? 2 : d }); }
export function stamp() { const d = new Date(), p = (x: number) => String(x).padStart(2, "0"); return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`; }
export function b64(s: string) { const bin = atob(s), a = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i); return a; }
export function toB64(u: Uint8Array) { let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, Array.from(u.subarray(i, i + 0x8000))); return btoa(s); }
export function firma(): any { try { return JSON.parse(localStorage.getItem("zf-firma") || "{}") || {}; } catch { return {}; } }
export function fmtVal(sh: any, k: number, v: any, cell: any): string {
  if (v == null || v === "") return "";
  if (typeof v === "number" && ((k >= 0 && sh.timeCol && sh.timeCol[k]) || (cell && cell.tm))) return ZF.fmtTime(v);
  if (typeof v === "number") {
    const z = (cell && cell.z) || (k >= 0 ? sh.colFmt[k] : null) || "", h = k >= 0 ? sh.headers[k] : "", zz = String(z).replace(/"[^"]*"|\\./g, "");
    if (/[dy]/i.test(zz) && !/h/i.test(zz) && v > 20000 && v < 80000) return ZF.serialToDate(v);
    if (/%/.test(zz)) return de(Math.round(v * 10000) / 100, null) + " %";
    if (k >= 0 && /(^|[^a-z])(nr|nummer|id|kd-?nr|kdnr|plz)([^a-z]|$)/i.test(h) && Number.isInteger(v)) return String(v);
    const money = /€|EUR/.test(z) || /€/.test(h), two = money || /0\.00/.test(z);
    return de(v, two ? 2 : null) + (money ? " €" : "");
  }
  return String(v);
}
export function rawOf(cell: any): string {
  if (!cell) return ""; if (cell.f) return "=" + cell.f;
  if (cell.tm && typeof cell.v === "number") return ZF.fmtTime(cell.v);
  if (typeof cell.v === "number") return String(cell.v).replace(".", ",");
  return cell.v == null ? "" : String(cell.v);
}

/** Läuft Zettelfrei als installiertes Programm (Tauri) statt im Browser? */
export const IS_APP = !!(window as any).__TAURI__;
export const HIER = IS_APP ? "auf diesem Rechner" : "in diesem Browser";
