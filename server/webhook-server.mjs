import { createServer } from "node:http";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { randomUUID } from "node:crypto";

const PORT = Number(process.env.TELEGRAM_WEBHOOK_PORT || 8787);
const HOST = process.env.TELEGRAM_WEBHOOK_HOST || "0.0.0.0";
const STATE_FILE = resolve(process.cwd(), "server", "telegram-bridge-state.json");
const MAX_REPORTS = 2000;

function ensureStateFile() {
  const dir = dirname(STATE_FILE);
  if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
  if (!existsSync(STATE_FILE)) {
    const initial = {
      token: "",
      chatId: "",
      lastUpdateId: 0,
      webhookBaseUrl: "",
      webhookSecret: randomUUID().replace(/-/g, ""),
      webhookPath: "",
      updatedAt: 0,
      reports: []
    };
    writeFileSync(STATE_FILE, JSON.stringify(initial, null, 2), "utf8");
  }
}

function readState() {
  ensureStateFile();
  try {
    return JSON.parse(readFileSync(STATE_FILE, "utf8"));
  } catch {
    return {
      token: "",
      chatId: "",
      lastUpdateId: 0,
      webhookBaseUrl: "",
      webhookSecret: randomUUID().replace(/-/g, ""),
      webhookPath: "",
      updatedAt: 0,
      reports: []
    };
  }
}

function writeState(nextState) {
  const safe = {
    token: String(nextState.token || ""),
    chatId: String(nextState.chatId || ""),
    lastUpdateId: Number(nextState.lastUpdateId || 0),
    webhookBaseUrl: String(nextState.webhookBaseUrl || ""),
    webhookSecret: String(nextState.webhookSecret || randomUUID().replace(/-/g, "")),
    webhookPath: String(nextState.webhookPath || ""),
    updatedAt: Number(nextState.updatedAt || Date.now()),
    reports: Array.isArray(nextState.reports) ? nextState.reports.slice(-MAX_REPORTS) : []
  };
  writeFileSync(STATE_FILE, JSON.stringify(safe, null, 2), "utf8");
  return safe;
}

function sendJson(res, status, payload) {
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  });
  res.end(JSON.stringify(payload));
}

function normalizeVietnamese(input) {
  return String(input || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "")
    .trim();
}

function normalizeDate(dateText) {
  const value = String(dateText || "").trim();
  if (!value) return new Date().toISOString().slice(0, 10);
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  const ddmmyyyy = value.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (ddmmyyyy) {
    const dd = ddmmyyyy[1].padStart(2, "0");
    const mm = ddmmyyyy[2].padStart(2, "0");
    const yyyy = ddmmyyyy[3];
    return `${yyyy}-${mm}-${dd}`;
  }
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return new Date().toISOString().slice(0, 10);
  return parsed.toISOString().slice(0, 10);
}

function parseFlexibleNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const text = String(value || "").trim();
  if (!text) return 0;
  const normalized = text.replace(/,/g, ".").replace(/[^\d.\-]/g, "");
  const parsed = Number.parseFloat(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseTelegramReportMessage(text) {
  if (!text || !String(text).toLowerCase().includes("#baocao")) return null;

  const lines = String(text).split(/[\r\n]+/).map((line) => line.trim()).filter(Boolean);
  const values = {};

  for (const line of lines) {
    const separatorIndex = line.indexOf(":");
    if (separatorIndex === -1) continue;
    const key = normalizeVietnamese(line.slice(0, separatorIndex));
    const value = line.slice(separatorIndex + 1).trim();
    if (!key || !value) continue;
    values[key] = value;
  }

  const nurse = values.ten || values.dieuduong || values.nurse || "";
  const customerName = values.khach || values.khachhang || values.customer || "";

  if (!nurse || !customerName) return null;

  return {
    nurse,
    registrationDate: normalizeDate(values.ngay || values.date || ""),
    appointmentTime: values.gio || values.time || "",
    customerName,
    service: values.dichvu || values.service || "",
    shiftMinutes: parseFlexibleNumber(values.thoiluong || values.phut || values.minutes || 0),
    distanceKm: parseFlexibleNumber(values.khoangcach || values.km || 0),
    status: values.trangthai || values.status || "completed",
    note: values.ghichu || values.note || "",
    source: "Telegram Webhook"
  };
}

async function parseJsonBody(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  if (!chunks.length) return {};
  const text = Buffer.concat(chunks).toString("utf8");
  if (!text.trim()) return {};
  return JSON.parse(text);
}

async function setTelegramWebhook(token, webhookUrl) {
  const endpoint = `https://api.telegram.org/bot${token}/setWebhook`;
  const response = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      url: webhookUrl,
      drop_pending_updates: false,
      allowed_updates: ["message"]
    })
  });
  const payload = await response.json();
  if (!response.ok || payload.ok === false) {
    throw new Error(payload.description || `setWebhook failed with HTTP ${response.status}`);
  }
}

async function deleteTelegramWebhook(token) {
  const endpoint = `https://api.telegram.org/bot${token}/deleteWebhook`;
  const response = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ drop_pending_updates: false })
  });
  const payload = await response.json();
  if (!response.ok || payload.ok === false) {
    throw new Error(payload.description || `deleteWebhook failed with HTTP ${response.status}`);
  }
  return payload.result;
}

async function getWebhookInfo(token) {
  const endpoint = `https://api.telegram.org/bot${token}/getWebhookInfo`;
  const response = await fetch(endpoint);
  const payload = await response.json();
  if (!response.ok || payload.ok === false) {
    throw new Error(payload.description || `getWebhookInfo failed with HTTP ${response.status}`);
  }
  return payload.result || {};
}

function appendTelegramReport(state, update) {
  const message = update?.message;
  const chatId = String(message?.chat?.id || "");
  if (!message || !message.text || chatId !== String(state.chatId || "")) return { saved: false, ignored: true };

  const raw = parseTelegramReportMessage(message.text);
  if (!raw) return { saved: false, ignored: true };

  raw.telegramUpdateId = String(update.update_id || `${Date.now()}`);
  const report = {
    id: String(update.update_id || `${Date.now()}`),
    receivedAt: Date.now(),
    raw
  };

  const existing = new Set((state.reports || []).map((item) => String(item.id)));
  if (!existing.has(report.id)) {
    state.reports = [...(state.reports || []), report].slice(-MAX_REPORTS);
    state.updatedAt = Date.now();
    return { saved: true, ignored: false };
  }
  return { saved: false, ignored: false };
}

async function fetchTelegramUpdates(state) {
  if (!state.token) return [];
  const offset = Number(state.lastUpdateId || 0) > 0 ? Number(state.lastUpdateId) + 1 : undefined;
  const endpoint = `https://api.telegram.org/bot${state.token}/getUpdates?limit=100${offset ? `&offset=${offset}` : ""}`;
  const response = await fetch(endpoint);
  const payload = await response.json();
  if (!response.ok || payload.ok === false) {
    const error = new Error(payload.description || `getUpdates failed with HTTP ${response.status}`);
    error.code = payload.error_code || response.status;
    throw error;
  }
  return Array.isArray(payload.result) ? payload.result : [];
}

async function pollTelegramUpdatesOnce() {
  const state = readState();
  if (!state.token || !state.chatId) return;

  try {
    const updates = await fetchTelegramUpdates(state);
    if (!updates.length) return;

    updates.forEach((update) => {
      state.lastUpdateId = Math.max(Number(state.lastUpdateId || 0), Number(update.update_id || 0));
      appendTelegramReport(state, update);
    });
    writeState(state);
  } catch (error) {
    if (error.code === 409) {
      try {
        const webhookInfo = await getWebhookInfo(state.token);
        if (webhookInfo?.last_error_message || Number(webhookInfo?.pending_update_count || 0) > 0) {
          await deleteTelegramWebhook(state.token);
          state.webhookBaseUrl = "";
          state.webhookPath = "";
          writeState(state);
        }
      } catch (cleanupError) {
        console.error("[telegram-webhook-bridge] failed to recover from webhook conflict:", cleanupError.message);
      }
      return;
    }
    console.error("[telegram-webhook-bridge] polling error:", error.message);
  }
}

const server = createServer(async (req, res) => {
  const method = req.method || "GET";
  const url = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);

  if (method === "OPTIONS") {
    sendJson(res, 200, { ok: true });
    return;
  }

  if (method === "GET" && url.pathname === "/api/telegram/health") {
    sendJson(res, 200, { ok: true, service: "telegram-webhook-bridge" });
    return;
  }

  if (method === "POST" && url.pathname === "/api/telegram/config") {
    try {
      const payload = await parseJsonBody(req);
      const state = readState();
      const token = String(payload.token || "").trim();
      const chatId = String(payload.chatId || "").trim();
      const webhookBaseUrl = String(payload.webhookBaseUrl || "").trim().replace(/\/+$/, "");

      if (!token || !chatId) {
        sendJson(res, 400, { ok: false, error: "Missing token or chatId" });
        return;
      }

      state.token = token;
      state.chatId = chatId;
      state.webhookBaseUrl = webhookBaseUrl;
      state.updatedAt = Date.now();
      if (!state.webhookSecret) state.webhookSecret = randomUUID().replace(/-/g, "");

      let webhookUrl = "";
      let webhookInfo = {};
      if (webhookBaseUrl) {
        webhookUrl = `${webhookBaseUrl}/api/telegram/webhook/${state.webhookSecret}`;
        await setTelegramWebhook(token, webhookUrl);
        webhookInfo = await getWebhookInfo(token);
        state.webhookPath = `/api/telegram/webhook/${state.webhookSecret}`;
      } else {
        await deleteTelegramWebhook(token);
        state.webhookPath = "";
      }

      const saved = writeState(state);
      sendJson(res, 200, {
        ok: true,
        configured: Boolean(saved.token && saved.chatId),
        webhookUrl,
        pendingCount: saved.reports.length,
        webhookInfo
      });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Config failed" });
    }
    return;
  }

  if (method === "GET" && url.pathname === "/api/telegram/pending") {
    try {
      const state = readState();
      const since = Number(url.searchParams.get("since") || 0);
      const includeAll = url.searchParams.get("all") === "1";
      const filtered = includeAll
        ? (state.reports || [])
        : (state.reports || []).filter((item) => Number(item.receivedAt || 0) > since);
      const lastReceivedAt = filtered.length
        ? Math.max(...filtered.map((item) => Number(item.receivedAt || 0)))
        : since;
      sendJson(res, 200, {
        ok: true,
        configured: Boolean(state.token && state.chatId),
        pendingCount: filtered.length,
        rows: filtered.map((item) => ({
          ...item.raw,
          telegramUpdateId: String(item.id || "")
        })),
        lastReceivedAt
      });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Read pending failed" });
    }
    return;
  }

  if (method === "POST" && url.pathname.startsWith("/api/telegram/webhook/")) {
    try {
      const state = readState();
      const incomingSecret = url.pathname.split("/").pop() || "";
      if (!incomingSecret || incomingSecret !== state.webhookSecret) {
        sendJson(res, 403, { ok: false, error: "Invalid webhook secret" });
        return;
      }

      const update = await parseJsonBody(req);
      state.lastUpdateId = Math.max(Number(state.lastUpdateId || 0), Number(update?.update_id || 0));
      const result = appendTelegramReport(state, update);
      writeState(state);
      sendJson(res, 200, { ok: true, accepted: result.saved, ignored: result.ignored });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Webhook processing failed" });
    }
    return;
  }

  sendJson(res, 404, { ok: false, error: "Not found" });
});

server.listen(PORT, HOST, () => {
  console.log(`[telegram-webhook-bridge] listening on http://${HOST}:${PORT}`);
});

setInterval(() => {
  pollTelegramUpdatesOnce().catch((error) => {
    console.error("[telegram-webhook-bridge] unexpected polling failure:", error.message);
  });
}, 10000);

pollTelegramUpdatesOnce().catch((error) => {
  console.error("[telegram-webhook-bridge] initial polling failure:", error.message);
});
