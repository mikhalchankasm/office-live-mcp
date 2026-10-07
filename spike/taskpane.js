"use strict";
// Feasibility test: can an add-in in Excel on the web reach a server on 127.0.0.1?
const PORT = 47863;
const HTTP_URL = `http://127.0.0.1:${PORT}/ping`;
const WS_URL = `ws://127.0.0.1:${PORT}/ws`;
const logBox = document.getElementById("log");
const button = document.getElementById("run");
const verdict = document.getElementById("verdict");
const report = { ua: navigator.userAgent, origin: location.origin, steps: {} };

function log(text, cls) {
  const line = document.createElement("div");
  if (cls) line.className = cls;
  line.textContent = `${new Date().toLocaleTimeString()}  ${text}`;
  logBox.appendChild(line);
}

function setVerdict(ok, text) {
  verdict.className = ok ? "ok" : "bad";
  verdict.textContent = text;
}

function policyInfo() {
  const out = {};
  const fp = document.featurePolicy || document.permissionsPolicy;
  for (const name of ["local-network-access", "local-network", "loopback-network"]) {
    try { out[name] = fp ? fp.allowsFeature(name) : "n/a"; } catch (e) { out[name] = "err"; }
  }
  return out;
}

async function permissionInfo() {
  const out = {};
  for (const name of ["local-network-access", "local-network", "loopback-network"]) {
    try { out[name] = (await navigator.permissions.query({ name })).state; } catch (e) { out[name] = "unsupported"; }
  }
  return out;
}

async function tryFetch() {
  const started = performance.now();
  try {
    const resp = await fetch(HTTP_URL, { cache: "no-store" });
    const body = await resp.json();
    const ms = Math.round(performance.now() - started);
    log(`HTTP fetch: OK ${resp.status} за ${ms} мс (сервер ${body.server})`, "ok");
    return { ok: true, ms };
  } catch (e) {
    log(`HTTP fetch: ошибка — ${e.name}: ${e.message}`, "bad");
    return { ok: false, error: `${e.name}: ${e.message}` };
  }
}

// Excel calls requested by the server.
const handlers = {
  async info() {
    return Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets.load("items/name");
      const active = ctx.workbook.worksheets.getActiveWorksheet().load("name");
      await ctx.sync();
      return { sheets: sheets.items.map((s) => s.name), active: active.name };
    });
  },
  async read({ address }) {
    return Excel.run(async (ctx) => {
      const rng = ctx.workbook.worksheets.getActiveWorksheet().getRange(address).load("values,formulas,address");
      await ctx.sync();
      return { address: rng.address, values: rng.values, formulas: rng.formulas };
    });
  },
  async write({ address, values }) {
    return Excel.run(async (ctx) => {
      const rng = ctx.workbook.worksheets.getActiveWorksheet().getRange(address);
      rng.values = values;
      rng.format.autofitColumns();
      await ctx.sync();
      return { written: address };
    });
  },
  async select({ address }) {
    return Excel.run(async (ctx) => {
      ctx.workbook.worksheets.getActiveWorksheet().getRange(address).select();
      await ctx.sync();
      return { selected: address };
    });
  },
};

function tryWebSocket(fetchResult) {
  return new Promise((resolve) => {
    const started = performance.now();
    let opened = false;
    let ws;
    try {
      ws = new WebSocket(WS_URL);
    } catch (e) {
      log(`WebSocket: исключение — ${e.name}: ${e.message}`, "bad");
      resolve({ ok: false, error: `${e.name}: ${e.message}` });
      return;
    }
    ws.onopen = () => {
      opened = true;
      log(`WebSocket: подключено за ${Math.round(performance.now() - started)} мс`, "ok");
      ws.send(JSON.stringify({
        type: "hello", ua: navigator.userAgent, origin: location.origin,
        host: Office.context.host, platform: Office.context.platform,
        excelApi17: Office.context.requirements.isSetSupported("ExcelApi", "1.7"),
        fetch: fetchResult, policy: report.policy, permissions: report.permissions,
      }));
    };
    ws.onmessage = async (event) => {
      const msg = JSON.parse(event.data);
      if (msg.type === "call") {
        try {
          const result = await handlers[msg.method](msg.params || {});
          if (msg.method !== "read" || !msg.quiet) log(`сервер → ${msg.method}: OK`, "ok");
          ws.send(JSON.stringify({ type: "result", id: msg.id, result }));
        } catch (e) {
          log(`сервер → ${msg.method}: ошибка — ${e.message}`, "bad");
          ws.send(JSON.stringify({ type: "error", id: msg.id, message: `${e.name}: ${e.message}` }));
        }
      } else if (msg.type === "done") {
        log(msg.summary, "ok");
        setVerdict(true, "✓ Работает: сервер на компьютере управляет этой книгой.");
        ws.close();
        resolve({ ok: true });
      }
    };
    ws.onerror = () => log("WebSocket: событие error (подробности браузер не сообщает)", "bad");
    ws.onclose = (event) => {
      if (!opened) {
        log(`WebSocket: не подключился, код ${event.code}`, "bad");
        resolve({ ok: false, error: `closed before open, code ${event.code}` });
      }
    };
  });
}

async function run() {
  button.disabled = true;
  logBox.textContent = "";
  verdict.className = "";
  report.policy = policyInfo();
  report.permissions = await permissionInfo();
  log(`Браузер: ${navigator.userAgent}`, "dim");
  log(`Origin: ${location.origin}; Office: ${Office.context.host} / ${Office.context.platform}`, "dim");
  log(`Permissions-Policy: ${JSON.stringify(report.policy)}`, "dim");
  log(`Разрешения: ${JSON.stringify(report.permissions)}`, "dim");
  report.steps.fetch = await tryFetch();
  report.steps.ws = await tryWebSocket(report.steps.fetch);
  if (!report.steps.ws.ok) {
    setVerdict(false, "✗ Браузер не пустил надстройку к серверу на компьютере. Если браузер спрашивал разрешение — нажмите «Разрешить» и повторите.");
  }
  button.disabled = false;
}

Office.onReady(() => {
  button.disabled = false;
  button.onclick = run;
  log("Office.js загружен. Убедитесь, что тестовый сервер запущен, и нажмите кнопку.", "dim");
});
