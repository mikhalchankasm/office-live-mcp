"use strict";
// Variants of reaching the local test server; shared by taskpane.js and direct.html.
const DIAG_PORT = 47863;

async function diagFetchVariants(log) {
  const variants = [
    ["fetch 127.0.0.1", `http://127.0.0.1:${DIAG_PORT}/ping?v=ip`, {}],
    ["fetch 127.0.0.1 + targetAddressSpace=loopback", `http://127.0.0.1:${DIAG_PORT}/ping?v=ip-tas`, { targetAddressSpace: "loopback" }],
    ["fetch localhost", `http://localhost:${DIAG_PORT}/ping?v=localhost`, {}],
    ["fetch localhost + targetAddressSpace=loopback", `http://localhost:${DIAG_PORT}/ping?v=localhost-tas`, { targetAddressSpace: "loopback" }],
  ];
  const results = {};
  for (const [label, url, extra] of variants) {
    const started = performance.now();
    try {
      const resp = await fetch(url, { cache: "no-store", ...extra });
      results[label] = `OK ${resp.status}`;
      log(`${label}: OK ${resp.status} за ${Math.round(performance.now() - started)} мс`, "ok");
    } catch (e) {
      results[label] = `${e.name}: ${e.message}`;
      log(`${label}: ${e.name}: ${e.message}`, "bad");
    }
  }
  return results;
}

function diagOpenSocket(url, log) {
  return new Promise((resolve) => {
    const started = performance.now();
    let ws;
    try {
      ws = new WebSocket(url);
    } catch (e) {
      log(`WebSocket ${url}: исключение ${e.name}: ${e.message}`, "bad");
      resolve(null);
      return;
    }
    ws.onopen = () => {
      log(`WebSocket ${url}: подключено за ${Math.round(performance.now() - started)} мс`, "ok");
      resolve(ws);
    };
    ws.onclose = (event) => {
      if (ws.readyState !== WebSocket.OPEN && !ws._opened) {
        log(`WebSocket ${url}: не подключился, код ${event.code}`, "bad");
        resolve(null);
      }
    };
    ws.addEventListener("open", () => { ws._opened = true; });
  });
}

async function diagSocket(log) {
  for (const host of ["127.0.0.1", "localhost"]) {
    const ws = await diagOpenSocket(`ws://${host}:${DIAG_PORT}/ws`, log);
    if (ws) return ws;
  }
  return null;
}
