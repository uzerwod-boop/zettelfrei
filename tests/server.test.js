// Prüft den Server mit einem nachgebauten Claude-Endpunkt: kein echter Schlüssel nötig.
const test = require("node:test"), assert = require("node:assert"), http = require("http");

test("Server: Oberfläche, Health, Weiterleitung an Claude", async () => {
  let seen = null;
  const fake = http.createServer((req, res) => {
    let b = ""; req.on("data", c => b += c); req.on("end", () => {
      seen = { headers: req.headers, body: JSON.parse(b) };
      res.writeHead(200, { "content-type": "application/json" });
      res.end(JSON.stringify({ content: [{ type: "text", text: "{\"ok\":true}" }] }));
    });
  });
  await new Promise(r => fake.listen(0, "127.0.0.1", r));
  process.env.ANTHROPIC_API_KEY = "test-key";
  process.env.ZF_MODEL = "claude-test";
  process.env.ZF_UPSTREAM = `http://127.0.0.1:${fake.address().port}/v1/messages`;
  delete require.cache[require.resolve("../server.js")];
  const server = require("../server.js");
  await new Promise(r => server.listen(0, "127.0.0.1", r));
  const base = `http://127.0.0.1:${server.address().port}`;
  try {
    const page = await fetch(base + "/");
    assert.equal(page.status, 200);
    assert.match(await page.text(), /<title>Zettelfrei<\/title>/);

    const health = await (await fetch(base + "/api/health")).json();
    assert.deepEqual(health, { ok: true, provider: "claude", model: "claude-test" });

    const r = await fetch(base + "/api/messages", { method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ model: "vom-browser-gewaehlt", max_tokens: 99999, messages: [{ role: "user", content: "hallo" }] }) });
    assert.equal(r.status, 200);
    assert.equal((await r.json()).content[0].text, "{\"ok\":true}");
    assert.equal(seen.headers["x-api-key"], "test-key", "Schlüssel kommt vom Server");
    assert.equal(seen.body.model, "claude-test", "Modell bestimmt der Server");
    assert.equal(seen.body.max_tokens, 8192, "max_tokens begrenzt");

    assert.equal((await fetch(base + "/%2e%2e/%2e%2e/etc/passwd")).status, 404);
    assert.equal((await fetch(base + "/.git/config")).status, 404);
  } finally {
    server.close(); fake.close();
  }
});

test("Server: OpenRouter-Schlüssel, Umwandlung ins OpenAI-Format", async () => {
  let seen = null;
  const fake = http.createServer((req, res) => {
    let b = ""; req.on("data", c => b += c); req.on("end", () => {
      seen = { headers: req.headers, body: JSON.parse(b) };
      res.writeHead(200, { "content-type": "application/json" });
      res.end(JSON.stringify({ choices: [{ message: { role: "assistant", content: "{\"ok\":true}" } }] }));
    });
  });
  await new Promise(r => fake.listen(0, "127.0.0.1", r));
  delete process.env.ANTHROPIC_API_KEY; delete process.env.ZF_MODEL;
  process.env.OPENROUTER_API_KEY = "or-key";
  process.env.ZF_UPSTREAM = `http://127.0.0.1:${fake.address().port}/api/v1/chat/completions`;
  delete require.cache[require.resolve("../server.js")];
  const server = require("../server.js");
  await new Promise(r => server.listen(0, "127.0.0.1", r));
  const base = `http://127.0.0.1:${server.address().port}`;
  try {
    assert.deepEqual(await (await fetch(base + "/api/health")).json(), { ok: true, provider: "openrouter", model: "anthropic/claude-sonnet-5.5" });
    const r = await fetch(base + "/api/messages", { method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ messages: [{ role: "user", content: [
        { type: "image", source: { type: "base64", media_type: "image/jpeg", data: "QUJD" } },
        { type: "text", text: "Rechnung lesen" }] }] }) });
    assert.equal(r.status, 200);
    assert.equal((await r.json()).content[0].text, "{\"ok\":true}", "Antwort kommt im gewohnten Format zurück");
    assert.equal(seen.headers.authorization, "Bearer or-key");
    assert.equal(seen.body.model, "anthropic/claude-sonnet-5.5");
    assert.deepEqual(seen.body.messages[0].content, [
      { type: "image_url", image_url: { url: "data:image/jpeg;base64,QUJD" } }, { type: "text", text: "Rechnung lesen" }]);
  } finally {
    server.close(); fake.close(); delete process.env.OPENROUTER_API_KEY;
  }
});
