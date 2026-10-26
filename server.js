#!/usr/bin/env node
/* Zettelfrei – kleiner Server ohne Abhängigkeiten.
 * Liefert die Oberfläche aus und leitet KI-Anfragen an die Claude API weiter,
 * damit der API-Schlüssel nur auf dem Server liegt.
 *
 *   OPENROUTER_API_KEY=sk-or-... npm start
 *   ANTHROPIC_API_KEY=sk-ant-... npm start
 *
 * Umgebungsvariablen:
 *   OPENROUTER_API_KEY Schlüssel für OpenRouter (viele Modelle über einen Schlüssel)
 *   ANTHROPIC_API_KEY  Schlüssel für die Claude API
 *                      (ohne beide: Oberfläche läuft, KI ist im Browser einstellbar)
 *   ZF_MODEL           Modell, Standard claude-sonnet-5-5 bzw. anthropic/claude-sonnet-5.5 bei OpenRouter
 *   PORT               Standard 8787
 *   HOST               Standard 127.0.0.1 (nur dieser Rechner). 0.0.0.0 öffnet ihn im Netzwerk.
 *   ZF_ALLOWED_HOSTS   weitere Hostnamen, unter denen der Server erreichbar sein darf (Komma-getrennt),
 *                      z. B. "buero-pc,192.168.1.20". Schützt vor DNS-Rebinding.
 */
"use strict";
const http = require("http");
const fs = require("fs");
const path = require("path");

const ROOT = __dirname;
const PORT = Number(process.env.PORT || 8787);
const HOST = process.env.HOST || "127.0.0.1";
const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY || "";
const OPENROUTER_KEY = process.env.OPENROUTER_API_KEY || "";
const PROVIDER = ANTHROPIC_KEY ? "claude" : OPENROUTER_KEY ? "openrouter" : null;
const KEY = ANTHROPIC_KEY || OPENROUTER_KEY;
const MODEL = process.env.ZF_MODEL || (PROVIDER === "openrouter" ? "anthropic/claude-sonnet-5.5" : "claude-sonnet-5-5");
const UPSTREAM = process.env.ZF_UPSTREAM || (PROVIDER === "openrouter" ? "https://openrouter.ai/api/v1/chat/completions" : "https://api.anthropic.com/v1/messages");
const MAX_BODY = 25 * 1024 * 1024;
const ALLOWED_HOSTS = new Set(["localhost", "127.0.0.1", "[::1]", HOST, ...(process.env.ZF_ALLOWED_HOSTS || "").split(",").map(h => h.trim().toLowerCase()).filter(Boolean)]);

// Schutz für die KI-Schnittstelle: nur Aufrufe der eigenen Oberfläche, keine fremden Webseiten (CSRF, DNS-Rebinding).
function hostName(h) { return String(h || "").toLowerCase().replace(/:\d+$/, ""); }
function allowedHost(req) { return HOST === "0.0.0.0" && !process.env.ZF_ALLOWED_HOSTS ? true : ALLOWED_HOSTS.has(hostName(req.headers.host)); }
function sameOrigin(req) {
  const site = req.headers["sec-fetch-site"];
  if (site && site !== "same-origin" && site !== "none") return false;
  const origin = req.headers.origin;
  if (origin) { try { if (new URL(origin).host.toLowerCase() !== String(req.headers.host || "").toLowerCase()) return false; } catch (e) { return false; } }
  return true;
}

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

// Anthropic-Format (vom Browser) -> OpenAI-Format (OpenRouter)
function toOpenAI(clean) {
  return {
    model: clean.model, max_tokens: clean.max_tokens, temperature: 0,
    messages: clean.messages.map(m => ({
      role: m.role,
      content: typeof m.content === "string" ? m.content : m.content.map(b =>
        b.type === "image" ? { type: "image_url", image_url: { url: `data:${b.source.media_type};base64,${b.source.data}` } }
                           : { type: "text", text: b.text })
    }))
  };
}

async function proxy(req, res) {
  if (!/^application\/json\b/i.test(req.headers["content-type"] || "")) return send(res, 415, { error: "Nur JSON" });
  if (!sameOrigin(req)) return send(res, 403, { error: "Fremde Herkunft" });
  if (!KEY) return send(res, 503, { error: "Kein OPENROUTER_API_KEY oder ANTHROPIC_API_KEY auf dem Server gesetzt" });
  let body;
  try { body = JSON.parse(await readBody(req)); } catch (e) { return send(res, e.status || 400, { error: e.status ? "Anfrage zu groß" : "Ungültige Anfrage" }); }
  // Nur das Nötige weiterreichen; Modell bestimmt der Server.
  const clean = { model: MODEL, max_tokens: Math.min(Number(body.max_tokens) || 4096, 8192), messages: body.messages };
  if (!Array.isArray(clean.messages)) return send(res, 400, { error: "messages fehlt" });
  try {
    if (PROVIDER === "openrouter") {
      const r = await fetch(UPSTREAM, {
        method: "POST",
        headers: { "content-type": "application/json", authorization: `Bearer ${KEY}`, "HTTP-Referer": "https://github.com/uzerwod-boop/zettelfrei", "X-OpenRouter-Title": "Zettelfrei" },
        body: JSON.stringify(toOpenAI(clean))
      });
      const data = await r.json().catch(() => ({}));
      if (!r.ok) return send(res, r.status, data);
      const msg = data.choices && data.choices[0] && data.choices[0].message;
      const text = msg ? (typeof msg.content === "string" ? msg.content : (msg.content || []).map(p => p.text || "").join("")) : "";
      return send(res, 200, { content: [{ type: "text", text }] });
    }
    const r = await fetch(UPSTREAM, {
      method: "POST",
      headers: { "content-type": "application/json", "x-api-key": KEY, "anthropic-version": "2023-06-01" },
      body: JSON.stringify(clean)
    });
    send(res, r.status, await r.text());
  } catch (e) {
    send(res, 502, { error: "KI-Anbieter nicht erreichbar" });
  }
}

const HAS_WEB = !process.env.ZF_KLASSISCH && fs.existsSync(path.join(ROOT, "web", "dist", "index.html"));

function serveFile(req, res) {
  let p = decodeURIComponent(new URL(req.url, "http://x").pathname);
  // Startseite: neue Oberfläche (web/dist), sobald sie gebaut ist (cd web && npm ci && npm run build),
  // sonst die klassische. Die klassische bleibt unter /klassisch/ erreichbar, /app/ von 0.5.0-Vorschau leitet um.
  if (p === "/app" || p === "/app/" || p === "/klassisch") { res.writeHead(301, { location: p === "/klassisch" ? "/klassisch/" : "/" }); return res.end(); }
  if (p.startsWith("/klassisch/")) p = p.slice(10) === "/" ? "/index.html" : p.slice(10);
  else if (p.startsWith("/app/")) p = "/web/dist/" + p.slice(5);
  else if (HAS_WEB && !/^\/(beispiele|docs)\//.test(p)) p = "/web/dist" + (p === "/" ? "/index.html" : p);
  else if (p === "/") p = "/index.html";
  const file = path.normalize(path.join(ROOT, p));
  if (!file.startsWith(ROOT + path.sep) || /(^|[\\/])(\.|node_modules)/.test(path.relative(ROOT, file))) return send(res, 404, "Nicht gefunden", "text/plain");
  fs.stat(file, (err, st) => {
    if (err || !st.isFile()) return send(res, 404, "Nicht gefunden", "text/plain");
    res.writeHead(200, { "content-type": TYPES[path.extname(file).toLowerCase()] || "application/octet-stream", "cache-control": "no-cache", "x-content-type-options": "nosniff", "referrer-policy": "no-referrer", "x-frame-options": "DENY" });
    fs.createReadStream(file).pipe(res);
  });
}

const server = http.createServer((req, res) => {
  let url = new URL(req.url, "http://x").pathname;
  if (url.startsWith("/app/api/")) url = url.slice(4);
  if (url.startsWith("/klassisch/api/")) url = url.slice(10);   // KI-Schnittstelle auch für die neue Oberfläche
  if (url.startsWith("/api/") && !allowedHost(req)) return send(res, 421, { error: "Unbekannter Hostname" });
  if (url === "/api/health" && req.method === "GET") return send(res, 200, { ok: !!KEY, provider: PROVIDER, model: KEY ? MODEL : null });
  if (url === "/api/messages" && req.method === "POST") return proxy(req, res);
  if (req.method !== "GET" && req.method !== "HEAD") return send(res, 405, { error: "Nicht erlaubt" });
  serveFile(req, res);
});

if (require.main === module) {
  server.listen(PORT, HOST, () => {
    console.log(`Zettelfrei läuft: http://${HOST === "0.0.0.0" ? "localhost" : HOST}:${PORT}`);
    console.log(KEY ? `KI: ${PROVIDER === "openrouter" ? "OpenRouter" : "Claude"} (${MODEL}) über diesen Server`
                    : "KI: kein Schlüssel gesetzt. Oberfläche läuft, KI kann im Browser eingerichtet werden (z. B. OpenRouter).");
  });
}
module.exports = server;
