#!/usr/bin/env node
/* Zettelfrei – kleiner Server ohne Abhängigkeiten.
 * Liefert die Oberfläche aus und leitet KI-Anfragen an die Claude API weiter,
 * damit der API-Schlüssel nur auf dem Server liegt.
 *
 *   ANTHROPIC_API_KEY=sk-ant-... npm start
 *
 * Umgebungsvariablen:
 *   ANTHROPIC_API_KEY  Schlüssel für die Claude API (ohne: Oberfläche läuft, KI im Browser einstellbar)
 *   ZF_MODEL           Modell, Standard claude-sonnet-5-5
 *   PORT               Standard 8787
 *   HOST               Standard 127.0.0.1 (nur dieser Rechner). 0.0.0.0 öffnet ihn im Netzwerk.
 */
"use strict";
const http = require("http");
const fs = require("fs");
const path = require("path");

const ROOT = __dirname;
const PORT = Number(process.env.PORT || 8787);
const HOST = process.env.HOST || "127.0.0.1";
const KEY = process.env.ANTHROPIC_API_KEY || "";
const MODEL = process.env.ZF_MODEL || "claude-sonnet-5-5";
const UPSTREAM = process.env.ZF_UPSTREAM || "https://api.anthropic.com/v1/messages";
const MAX_BODY = 25 * 1024 * 1024;

const TYPES = {
  ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8",
  ".json": "application/json", ".svg": "image/svg+xml", ".png": "image/png", ".ttf": "font/ttf", ".txt": "text/plain; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".md": "text/markdown; charset=utf-8"
};

function send(res, status, body, type) {
  res.writeHead(status, { "content-type": type || "application/json", "cache-control": "no-store", "x-content-type-options": "nosniff" });
  res.end(typeof body === "string" || Buffer.isBuffer(body) ? body : JSON.stringify(body));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0; const chunks = [];
    req.on("data", c => { size += c.length; if (size > MAX_BODY) { reject(Object.assign(new Error("zu groß"), { status: 413 })); req.destroy(); } else chunks.push(c); });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

async function proxy(req, res) {
  if (!KEY) return send(res, 503, { error: "ANTHROPIC_API_KEY ist auf dem Server nicht gesetzt" });
  let body;
  try { body = JSON.parse(await readBody(req)); } catch (e) { return send(res, e.status || 400, { error: e.status ? "Anfrage zu groß" : "Ungültige Anfrage" }); }
  // Nur das Nötige weiterreichen; Modell bestimmt der Server.
  const clean = { model: MODEL, max_tokens: Math.min(Number(body.max_tokens) || 4096, 8192), messages: body.messages };
  if (!Array.isArray(clean.messages)) return send(res, 400, { error: "messages fehlt" });
  try {
    const r = await fetch(UPSTREAM, {
      method: "POST",
      headers: { "content-type": "application/json", "x-api-key": KEY, "anthropic-version": "2023-06-01" },
      body: JSON.stringify(clean)
    });
    const text = await r.text();
    send(res, r.status, text);
  } catch (e) {
    send(res, 502, { error: "Claude API nicht erreichbar" });
  }
}

function serveFile(req, res) {
  let p = decodeURIComponent(new URL(req.url, "http://x").pathname);
  if (p === "/") p = "/index.html";
  const file = path.normalize(path.join(ROOT, p));
  if (!file.startsWith(ROOT + path.sep) || /(^|[\\/])(\.|node_modules)/.test(path.relative(ROOT, file))) return send(res, 404, "Nicht gefunden", "text/plain");
  fs.stat(file, (err, st) => {
    if (err || !st.isFile()) return send(res, 404, "Nicht gefunden", "text/plain");
    res.writeHead(200, { "content-type": TYPES[path.extname(file).toLowerCase()] || "application/octet-stream", "cache-control": "no-cache" });
    fs.createReadStream(file).pipe(res);
  });
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, "http://x").pathname;
  if (url === "/api/health" && req.method === "GET") return send(res, 200, { ok: !!KEY, provider: "claude", model: KEY ? MODEL : null });
  if (url === "/api/messages" && req.method === "POST") return proxy(req, res);
  if (req.method !== "GET" && req.method !== "HEAD") return send(res, 405, { error: "Nicht erlaubt" });
  serveFile(req, res);
});

if (require.main === module) {
  server.listen(PORT, HOST, () => {
    console.log(`Zettelfrei läuft: http://${HOST === "0.0.0.0" ? "localhost" : HOST}:${PORT}`);
    console.log(KEY ? `KI: Claude (${MODEL}) über diesen Server` : "KI: kein ANTHROPIC_API_KEY gesetzt. Oberfläche läuft, KI kann im Browser eingerichtet werden.");
  });
}
module.exports = server;
