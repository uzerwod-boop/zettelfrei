const P: Record<string, string> = {
  chat: "M4 5h16v11H9l-5 4zM8 9h8M8 12h5", file: "M7 3h7l5 5v13H7zM14 3v5h5", doc: "M6 3h12v18H6zM9 8h6M9 12h6M9 16h4",
  shield: "M12 3l8 4v5c0 4.5-3.4 8-8 9-4.6-1-8-4.5-8-9V7zM9 12l2 2 4-4", history: "M12 7v5l3 2M21 12a9 9 0 1 1-3-6.7M21 4v5h-5",
  send: "M5 12h13M13 6l6 6-6 6", stop: "M7 7h10v10H7z", settings: "M12 9a3 3 0 1 0 0 6 3 3 0 0 0 0-6zM19 12h2M3 12h2M12 3v2M12 19v2M17 7l1.4-1.4M5.6 18.4 7 17M17 17l1.4 1.4M5.6 5.6 7 7",
};
export function Icon({ name, size = 22 }: { name: string; size?: number }) {
  return <svg viewBox="0 0 24 24" width={size} height={size} fill="none" stroke="currentColor" strokeWidth={1.8} strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d={P[name]} /></svg>;
}
export function Mic() { return <svg viewBox="0 0 24 24" width={19} height={19} aria-hidden="true"><path fill="currentColor" d="M12 15a3 3 0 0 0 3-3V6a3 3 0 1 0-6 0v6a3 3 0 0 0 3 3Zm5-3a5 5 0 0 1-10 0H5a7 7 0 0 0 6 6.92V22h2v-3.08A7 7 0 0 0 19 12h-2Z" /></svg>; }
// Markenzeichen: 4×4 Zellen, die gefüllten bilden ein Z
export function Mark({ dark = false, size = 34 }: { dark?: boolean; size?: number }) {
  const cells = "####..y..#..####".split("");
  return (
    <svg className="mark" width={size} height={size} viewBox="0 0 64 64" aria-hidden="true">
      {cells.map((ch, i) => { const x = (i % 4) * 17, y = Math.floor(i / 4) * 17;
        return ch === "." ? <rect key={i} x={x + 1.1} y={y + 1.1} width={10.8} height={10.8} rx={2.4} fill="none" stroke={dark ? "#4a5658" : "#cdd2ca"} strokeWidth={2.2} />
          : <rect key={i} x={x} y={y} width={13} height={13} rx={3.5} fill={ch === "y" ? "#f6dd6a" : dark ? "#4fbfa6" : "#0f6b5c"} />; })}
    </svg>);
}
