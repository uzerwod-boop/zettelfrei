// Kopiert die bewährten Kernmodule (Engine, E-Rechnung, Dokumente, Speichern, KI) und Bibliotheken
// unverändert aus dem Hauptprojekt nach public/lib. Eine Quelle, zwei Oberflächen.
import { cpSync, mkdirSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
const here = dirname(fileURLToPath(import.meta.url)), root = join(here, "..", ".."), out = join(here, "..", "public", "lib");
rmSync(out, { recursive: true, force: true }); mkdirSync(out, { recursive: true });
for (const f of ["engine.js", "einvoice.js", "docs.js", "xlsxio.js", "examples.js", "ai.js", "workflows.js"]) cpSync(join(root, "src", f), join(out, f));
cpSync(join(root, "vendor"), join(out, "vendor"), { recursive: true });
cpSync(join(root, "fonts"), join(out, "fonts"), { recursive: true });
console.log("Kernmodule nach public/lib kopiert");
