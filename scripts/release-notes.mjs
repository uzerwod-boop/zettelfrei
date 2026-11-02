// Gibt den CHANGELOG-Abschnitt einer Version aus (für Release-Text und Update-Hinweis im Programm).
// Aufruf: node scripts/release-notes.mjs 0.6.2
import { readFileSync } from "node:fs";
const v = process.argv[2], md = readFileSync(new URL("../CHANGELOG.md", import.meta.url), "utf8");
const start = md.search(new RegExp(`^## \\[${v.replace(/\./g, "\\.")}\\]`, "m"));
if (start < 0) { console.log(`Zettelfrei ${v}`); process.exit(0); }
const rest = md.slice(start), next = rest.slice(3).search(/^## \[/m);
const sec = (next < 0 ? rest : rest.slice(0, next + 3)).trim().split("\n");
const title = sec.shift().replace(/^## \[[^\]]+\]\s*–\s*[\d-]+\s*·?\s*/, "").trim();
console.log([title ? `**${title}**` : "", ...sec].join("\n").trim() +
  `\n\nDownload: macOS \`.dmg\`, Windows \`.exe\`/\`.msi\` unten unter Assets · [Erste Schritte](https://github.com/uzerwod-boop/zettelfrei/blob/main/docs/Erste-Schritte.md) · [CHANGELOG](https://github.com/uzerwod-boop/zettelfrei/blob/main/CHANGELOG.md)`);
