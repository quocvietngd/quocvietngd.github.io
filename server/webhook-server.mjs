import { createServer } from "node:http";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { randomUUID } from "node:crypto";
import pg from "pg";

const { Pool } = pg;

const PORT = Number(process.env.PORT || process.env.TELEGRAM_WEBHOOK_PORT || 8787);
const HOST = process.env.TELEGRAM_WEBHOOK_HOST || "0.0.0.0";
const RENDER_DISK_ROOT = "/var/data";
const RENDER_DISK_PATH = resolve(RENDER_DISK_ROOT);
const DEFAULT_STATE_FILE = existsSync(RENDER_DISK_ROOT)
  ? resolve(RENDER_DISK_ROOT, "telegram-bridge-state.json")
  : resolve(process.cwd(), "server", "telegram-bridge-state.json");
const STATE_FILE = process.env.STATE_FILE
  ? resolve(process.env.STATE_FILE)
  : DEFAULT_STATE_FILE;
const MAX_REPORTS = 2000;
const MAX_USERS = 500;
const DATABASE_URL = String(process.env.DATABASE_URL || "").trim();
const USE_POSTGRES = Boolean(DATABASE_URL);
const PERSISTENCE_MODE = String(process.env.PERSISTENCE_MODE || "auto").trim().toLowerCase();
const PG_STATE_KEY = "global_state";
const DEFAULT_TELEGRAM_CHANNEL = "nurse";
const SUPPORTED_TELEGRAM_CHANNELS = new Set(["nurse", "marketing"]);
const DEFAULT_BACKUP_DIR = existsSync(RENDER_DISK_ROOT)
  ? resolve(RENDER_DISK_ROOT, "backups")
  : resolve(process.cwd(), "server", "backups");
const BACKUP_DIR = process.env.BACKUP_DIR
  ? resolve(process.env.BACKUP_DIR)
  : DEFAULT_BACKUP_DIR;
const BACKUP_INTERVAL_MS = Math.max(5, Number(process.env.BACKUP_INTERVAL_MINUTES || 30)) * 60 * 1000;
const IS_RENDER_DISK_ATTACHED = existsSync(RENDER_DISK_ROOT);
const IS_STATE_FILE_ON_RENDER_DISK = STATE_FILE === RENDER_DISK_PATH || STATE_FILE.startsWith(`${RENDER_DISK_PATH}/`);
const IS_DURABLE_FILE_MODE = !USE_POSTGRES && IS_RENDER_DISK_ATTACHED && IS_STATE_FILE_ON_RENDER_DISK;
const IS_DURABLE_STORAGE = USE_POSTGRES || IS_DURABLE_FILE_MODE;

let dbPool = null;
let dbReadyPromise = null;

function assertPersistenceConfig() {
  if (!["auto", "durable-only", "postgres-only", "disk-only"].includes(PERSISTENCE_MODE)) {
    throw new Error(`Unsupported PERSISTENCE_MODE: ${PERSISTENCE_MODE}`);
  }

  if (PERSISTENCE_MODE === "postgres-only" && !USE_POSTGRES) {
    throw new Error("PERSISTENCE_MODE=postgres-only nhưng DATABASE_URL chưa được cấu hình.");
  }

  if (PERSISTENCE_MODE === "disk-only" && !IS_DURABLE_FILE_MODE) {
    throw new Error(`PERSISTENCE_MODE=disk-only nhưng STATE_FILE hiện tại không nằm trên ổ bền vững: ${STATE_FILE}`);
  }

  if (PERSISTENCE_MODE === "durable-only" && !IS_DURABLE_STORAGE) {
    throw new Error(`PERSISTENCE_MODE=durable-only nhưng storage hiện tại chưa bền vững: ${USE_POSTGRES ? "postgres" : STATE_FILE}`);
  }
}

const DEFAULT_USERS = [
  {
    id: "u-admin",
    userCode: "NR001",
    username: "admin",
    password: "NORA-ADMIN-2026",
    fullName: "System Admin",
    roleKey: "admin",
    department: "Vận hành",
    phone: "",
    email: "",
    address: "",
    bankAccount: "",
    status: "active",
    createdAt: Date.now()
  },
  {
    id: "u-ceo",
    userCode: "NR002",
    username: "ceo",
    password: "NORA-CEO-2026",
    fullName: "CEO Demo",
    roleKey: "ceo",
    department: "Ban điều hành",
    phone: "",
    email: "",
    address: "",
    bankAccount: "",
    status: "active",
    createdAt: Date.now()
  },
  {
    id: "u-head-tech",
    userCode: "NR003",
    username: "head-tech",
    password: "NORA-HEAD-2026",
    fullName: "Trưởng BP Kỹ thuật",
    roleKey: "head",
    department: "Kỹ thuật",
    phone: "",
    email: "",
    address: "",
    bankAccount: "",
    status: "active",
    createdAt: Date.now()
  }
];

function normalizeAppState(input = {}) {
  return {
    schemaVersion: Number(input.schemaVersion) || 1,
    customers: Array.isArray(input.customers) ? input.customers : [],
    schedules: Array.isArray(input.schedules) ? input.schedules : [],
    inventoryItems: Array.isArray(input.inventoryItems) ? input.inventoryItems : [],
    inventoryTransactions: Array.isArray(input.inventoryTransactions) ? input.inventoryTransactions : [],
    hrFiles: input.hrFiles && typeof input.hrFiles === "object" ? input.hrFiles : {},
    customerCareProgress: input.customerCareProgress && typeof input.customerCareProgress === "object" ? input.customerCareProgress : {},
    customerCareFilters: input.customerCareFilters && typeof input.customerCareFilters === "object" ? input.customerCareFilters : {},
    activities: Array.isArray(input.activities) ? input.activities : [],
    recycleBin: Array.isArray(input.recycleBin) ? input.recycleBin : [],
    rolePermissions: input.rolePermissions && typeof input.rolePermissions === "object" ? input.rolePermissions : {},
    newsPosts: Array.isArray(input.newsPosts) ? input.newsPosts : [],
    newsPinned: Array.isArray(input.newsPinned) ? input.newsPinned : [],
    newsEvents: Array.isArray(input.newsEvents) ? input.newsEvents : [],
    accountingCashflow: Array.isArray(input.accountingCashflow) ? input.accountingCashflow : [],
    accountingCashflowFilters: input.accountingCashflowFilters && typeof input.accountingCashflowFilters === "object" ? input.accountingCashflowFilters : {},
    accountingAttendance: Array.isArray(input.accountingAttendance) ? input.accountingAttendance : [],
    accountingAttendanceSource: input.accountingAttendanceSource && typeof input.accountingAttendanceSource === "object" ? input.accountingAttendanceSource : {},
    accountingAttendanceFilters: input.accountingAttendanceFilters && typeof input.accountingAttendanceFilters === "object" ? input.accountingAttendanceFilters : {},
    accountingServicePayrollFilters: input.accountingServicePayrollFilters && typeof input.accountingServicePayrollFilters === "object" ? input.accountingServicePayrollFilters : {},
    nurseReportOverrides: input.nurseReportOverrides && typeof input.nurseReportOverrides === "object" ? input.nurseReportOverrides : {},
    telegramSource: input.telegramSource && typeof input.telegramSource === "object" ? input.telegramSource : {},
    dataSourceConfig: input.dataSourceConfig && typeof input.dataSourceConfig === "object" ? input.dataSourceConfig : { type: "local", url: "" },
    reports: Array.isArray(input.reports) ? input.reports : [],
    updatedAt: Number(input.updatedAt) || 0
  };
}

function normalizeTelegramChannelKey(input) {
  const value = String(input || DEFAULT_TELEGRAM_CHANNEL).trim().toLowerCase();
  return value || DEFAULT_TELEGRAM_CHANNEL;
}

function getDefaultTelegramChannelState(raw = {}) {
  return {
    token: String(raw.token || ""),
    chatId: String(raw.chatId || ""),
    lastUpdateId: Number(raw.lastUpdateId || 0),
    webhookBaseUrl: String(raw.webhookBaseUrl || ""),
    webhookSecret: String(raw.webhookSecret || randomUUID().replace(/-/g, "")),
    webhookPath: String(raw.webhookPath || ""),
    updatedAt: Number(raw.updatedAt || 0),
    reports: Array.isArray(raw.reports) ? raw.reports.slice(-MAX_REPORTS) : []
  };
}

function normalizeTelegramChannels(raw = {}) {
  const normalized = {};
  const source = raw && typeof raw === "object" ? raw : {};
  const legacyNurse = {
    token: source.token,
    chatId: source.chatId,
    lastUpdateId: source.lastUpdateId,
    webhookBaseUrl: source.webhookBaseUrl,
    webhookSecret: source.webhookSecret,
    webhookPath: source.webhookPath,
    updatedAt: source.updatedAt,
    reports: source.reports
  };
  const channelSource = source.telegramChannels && typeof source.telegramChannels === "object"
    ? source.telegramChannels
    : {};

  SUPPORTED_TELEGRAM_CHANNELS.forEach((channelKey) => {
    const rawChannel = channelSource[channelKey] && typeof channelSource[channelKey] === "object"
      ? channelSource[channelKey]
      : (channelKey === DEFAULT_TELEGRAM_CHANNEL ? legacyNurse : {});
    normalized[channelKey] = getDefaultTelegramChannelState(rawChannel);
  });

  return normalized;
}

function getTelegramChannelState(state, channelKey) {
  const normalizedKey = normalizeTelegramChannelKey(channelKey);
  if (!state.telegramChannels || typeof state.telegramChannels !== "object") {
    state.telegramChannels = normalizeTelegramChannels(state);
  }
  if (!state.telegramChannels[normalizedKey]) {
    state.telegramChannels[normalizedKey] = getDefaultTelegramChannelState();
  }
  return state.telegramChannels[normalizedKey];
}

function normalizeState(raw = {}) {
  const telegramChannels = normalizeTelegramChannels(raw);
  const nurseChannel = telegramChannels[DEFAULT_TELEGRAM_CHANNEL] || getDefaultTelegramChannelState();
  return {
    token: String(nurseChannel.token || ""),
    chatId: String(nurseChannel.chatId || ""),
    lastUpdateId: Number(nurseChannel.lastUpdateId || 0),
    webhookBaseUrl: String(nurseChannel.webhookBaseUrl || ""),
    webhookSecret: String(nurseChannel.webhookSecret || randomUUID().replace(/-/g, "")),
    webhookPath: String(nurseChannel.webhookPath || ""),
    updatedAt: Number(raw.updatedAt || nurseChannel.updatedAt || 0),
    reports: Array.isArray(nurseChannel.reports) ? nurseChannel.reports.slice(-MAX_REPORTS) : [],
    telegramChannels,
    users: normalizeUsersList(raw.users && Array.isArray(raw.users) && raw.users.length ? raw.users : DEFAULT_USERS),
    kpiReports: normalizeKpiReportsList(raw.kpiReports),
    appState: normalizeAppState(raw.appState)
  };
}

async function ensureDbReady() {
  if (!USE_POSTGRES) return;
  if (dbReadyPromise) {
    await dbReadyPromise;
    return;
  }

  dbPool = new Pool({
    connectionString: DATABASE_URL,
    ssl: DATABASE_URL.includes("localhost") ? false : { rejectUnauthorized: false }
  });

  dbReadyPromise = (async () => {
    await dbPool.query(`
      CREATE TABLE IF NOT EXISTS app_kv (
        key TEXT PRIMARY KEY,
        value JSONB NOT NULL,
        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
      )
    `);
  })();

  await dbReadyPromise;
}

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
      reports: [],
      users: DEFAULT_USERS,
      kpiReports: []
    };
    writeFileSync(STATE_FILE, JSON.stringify(initial, null, 2), "utf8");
  }
}

function readStateFromFileSync() {
  ensureStateFile();
  try {
    const raw = JSON.parse(readFileSync(STATE_FILE, "utf8"));
    return normalizeState(raw);
  } catch {
    return normalizeState({
      token: "",
      chatId: "",
      lastUpdateId: 0,
      webhookBaseUrl: "",
      webhookSecret: randomUUID().replace(/-/g, ""),
      webhookPath: "",
      updatedAt: 0,
      reports: [],
      users: DEFAULT_USERS,
      kpiReports: [],
      appState: {}
    });
  }
}

function writeStateToFileSync(nextState) {
  const safe = normalizeState({
    ...nextState,
    updatedAt: Number(nextState.updatedAt || Date.now())
  });
  writeFileSync(STATE_FILE, JSON.stringify(safe, null, 2), "utf8");
  return safe;
}

async function readStateFromDb() {
  await ensureDbReady();
  const result = await dbPool.query("SELECT value FROM app_kv WHERE key = $1", [PG_STATE_KEY]);
  if (!result.rows.length) {
    const initial = normalizeState({
      token: "",
      chatId: "",
      lastUpdateId: 0,
      webhookBaseUrl: "",
      webhookSecret: randomUUID().replace(/-/g, ""),
      webhookPath: "",
      updatedAt: 0,
      reports: [],
      users: DEFAULT_USERS,
      kpiReports: [],
      appState: {}
    });
    await writeStateToDb(initial);
    return initial;
  }
  return normalizeState(result.rows[0].value || {});
}

async function writeStateToDb(nextState) {
  await ensureDbReady();
  const safe = normalizeState({
    ...nextState,
    updatedAt: Number(nextState.updatedAt || Date.now())
  });
  await dbPool.query(
    `INSERT INTO app_kv(key, value, updated_at)
     VALUES ($1, $2::jsonb, NOW())
     ON CONFLICT (key)
     DO UPDATE SET value = EXCLUDED.value, updated_at = NOW()`,
    [PG_STATE_KEY, JSON.stringify(safe)]
  );
  return safe;
}

async function readState() {
  if (USE_POSTGRES) return readStateFromDb();
  return readStateFromFileSync();
}

async function writeState(nextState) {
  if (USE_POSTGRES) return writeStateToDb(nextState);
  return writeStateToFileSync(nextState);
}

function normalizeUser(input = {}) {
  return {
    id: String(input.id || `u-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`),
    userCode: String(input.userCode || "").trim().toUpperCase(),
    username: String(input.username || "").trim().toLowerCase(),
    password: String(input.password || ""),
    fullName: String(input.fullName || "").trim(),
    roleKey: String(input.roleKey || "staff"),
    department: String(input.department || "Ban điều hành"),
    phone: String(input.phone || ""),
    email: String(input.email || ""),
    address: String(input.address || ""),
    bankAccount: String(input.bankAccount || ""),
    status: input.status === "suspended" ? "suspended" : "active",
    createdAt: Number(input.createdAt) || Date.now()
  };
}

function normalizeUsersList(list) {
  if (!Array.isArray(list)) return [];
  return list
    .map((item) => normalizeUser(item))
    .filter((user) => user.username)
    .slice(-MAX_USERS);
}

function normalizeKpiReport(input = {}) {
  return {
    date: normalizeDate(input.date),
    department: String(input.department || "Chưa xác định").trim() || "Chưa xác định",
    completion: parseFlexibleNumber(input.completion),
    quality: parseFlexibleNumber(input.quality),
    issues: parseFlexibleNumber(input.issues),
    submitter: String(input.submitter || "Unknown").trim() || "Unknown",
    updatedAt: Number(input.updatedAt) || Date.now()
  };
}

function normalizeKpiReportsList(list) {
  if (!Array.isArray(list)) return [];
  return list.map((item) => normalizeKpiReport(item)).slice(-MAX_REPORTS);
}

function sendJson(res, status, payload) {
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,PUT,DELETE,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  });
  res.end(JSON.stringify(payload));
}

function pickUsersPayload(payload) {
  if (Array.isArray(payload)) return payload;
  if (payload && Array.isArray(payload.users)) return payload.users;
  return [];
}

function pickReportsPayload(payload) {
  if (Array.isArray(payload)) return payload;
  if (payload && Array.isArray(payload.reports)) return payload.reports;
  return [];
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

function parseTelegramMarketingMessage(text) {
  const value = String(text || "");
  if (!value) return null;
  const lowered = value.toLowerCase();
  const hasKeyword = lowered.includes("#mkt")
    || lowered.includes("#marketing")
    || lowered.includes("#baocao_marketing")
    || lowered.includes("#baocaomarketing");
  if (!hasKeyword) return null;

  const lines = value.split(/[\r\n]+/).map((line) => line.trim()).filter(Boolean);
  const fields = {};
  for (const line of lines) {
    const separatorIndex = line.indexOf(":");
    if (separatorIndex === -1) continue;
    const key = normalizeVietnamese(line.slice(0, separatorIndex));
    const fieldValue = line.slice(separatorIndex + 1).trim();
    if (!key || !fieldValue) continue;
    fields[key] = fieldValue;
  }

  const reportDate = normalizeDate(fields.ngay || fields.date || "");
  const marketingName = fields.ten || fields.marketing || fields.marketer || fields.nhanvien || "";
  const budget = parseFlexibleNumber(fields.ngansach || fields.budget || fields.ads || fields.adspend || 0);
  const messCount = Math.max(0, Math.round(parseFlexibleNumber(fields.mess || fields.luongmess || fields.tuongtac || 0)));
  const phoneCount = Math.max(0, Math.round(parseFlexibleNumber(fields.sdt || fields.sodienthoai || fields.phone || fields.phones || 0)));
  const bookedCount = Math.max(0, Math.round(parseFlexibleNumber(fields.lich || fields.datlich || fields.booked || 0)));
  const contractCount = Math.max(0, Math.round(parseFlexibleNumber(fields.hopdong || fields.contracts || fields.hd || 0)));
  const revenue = parseFlexibleNumber(fields.doanhso || fields.revenue || fields.giatridh || fields.giatrihopdong || 0);

  if (!marketingName && messCount <= 0 && phoneCount <= 0 && bookedCount <= 0 && contractCount <= 0 && revenue <= 0 && budget <= 0) {
    return null;
  }

  return {
    registrationDate: reportDate,
    marketingName: marketingName || "Chưa gán",
    marketingBudget: budget,
    marketingMessCount: messCount,
    marketingPhoneCount: phoneCount,
    marketingBookedCount: bookedCount,
    marketingContractCount: contractCount,
    marketingRevenue: revenue,
    contractAmount: revenue,
    source: "Telegram Marketing",
    status: bookedCount > 0 || contractCount > 0 ? "completed" : "pending"
  };
}

function parseTelegramMessageByChannel(channelKey, text) {
  if (channelKey === "marketing") return parseTelegramMarketingMessage(text);
  return parseTelegramReportMessage(text);
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

function appendTelegramReport(state, update, channelKey = DEFAULT_TELEGRAM_CHANNEL) {
  const channelState = getTelegramChannelState(state, channelKey);
  const message = update?.message;
  const chatId = String(message?.chat?.id || "");
  if (!message || !message.text || chatId !== String(channelState.chatId || "")) return { saved: false, ignored: true };

  const raw = parseTelegramMessageByChannel(channelKey, message.text);
  if (!raw) return { saved: false, ignored: true };

  raw.telegramUpdateId = String(update.update_id || `${Date.now()}`);
  raw.telegramChannel = channelKey;
  const report = {
    id: `${channelKey}-${String(update.update_id || `${Date.now()}`)}`,
    receivedAt: Date.now(),
    raw
  };

  const existing = new Set((channelState.reports || []).map((item) => String(item.id)));
  if (!existing.has(report.id)) {
    channelState.reports = [...(channelState.reports || []), report].slice(-MAX_REPORTS);
    channelState.updatedAt = Date.now();
    state.updatedAt = Date.now();
    return { saved: true, ignored: false };
  }
  return { saved: false, ignored: false };
}

async function fetchTelegramUpdates(channelState) {
  if (!channelState.token) return [];
  const offset = Number(channelState.lastUpdateId || 0) > 0 ? Number(channelState.lastUpdateId) + 1 : undefined;
  const endpoint = `https://api.telegram.org/bot${channelState.token}/getUpdates?limit=100${offset ? `&offset=${offset}` : ""}`;
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
  const state = await readState();
  const channels = state.telegramChannels && typeof state.telegramChannels === "object"
    ? state.telegramChannels
    : normalizeTelegramChannels(state);
  state.telegramChannels = channels;

  let hasChanges = false;
  for (const [channelKey, channelState] of Object.entries(channels)) {
    if (!channelState?.token || !channelState?.chatId) continue;
    try {
      const updates = await fetchTelegramUpdates(channelState);
      if (!updates.length) continue;
      updates.forEach((update) => {
        channelState.lastUpdateId = Math.max(Number(channelState.lastUpdateId || 0), Number(update.update_id || 0));
        const result = appendTelegramReport(state, update, channelKey);
        if (result.saved) hasChanges = true;
      });
      hasChanges = true;
    } catch (error) {
      if (error.code === 409) {
        try {
          const webhookInfo = await getWebhookInfo(channelState.token);
          if (webhookInfo?.last_error_message || Number(webhookInfo?.pending_update_count || 0) > 0) {
            await deleteTelegramWebhook(channelState.token);
            channelState.webhookBaseUrl = "";
            channelState.webhookPath = "";
            hasChanges = true;
          }
        } catch (cleanupError) {
          console.error(`[telegram-webhook-bridge] failed to recover channel ${channelKey}:`, cleanupError.message);
        }
        continue;
      }
      console.error(`[telegram-webhook-bridge] polling error (${channelKey}):`, error.message);
    }
  }

  if (hasChanges) {
    state.updatedAt = Date.now();
    await writeState(state);
  }
}

function resolveRequestedChannel(url, payload = {}) {
  const fromPayload = normalizeTelegramChannelKey(payload.channel || payload.channelKey);
  const fromQuery = normalizeTelegramChannelKey(url.searchParams.get("channel"));
  const channelKey = (payload.channel || payload.channelKey) ? fromPayload : fromQuery;
  if (!SUPPORTED_TELEGRAM_CHANNELS.has(channelKey)) {
    throw new Error(`Unsupported telegram channel: ${channelKey}`);
  }
  return channelKey;
}

async function writeBackupSnapshot() {
  try {
    const state = await readState();
    await mkdir(BACKUP_DIR, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const filePath = resolve(BACKUP_DIR, `nora-state-${stamp}.json`);
    await writeFile(filePath, JSON.stringify(state, null, 2), "utf8");
  } catch (error) {
    console.error("[telegram-webhook-bridge] backup snapshot failed:", error.message);
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

  if (method === "GET" && url.pathname === "/api/users") {
    const state = await readState();
    sendJson(res, 200, state.users || []);
    return;
  }

  if ((method === "PUT" || method === "POST") && url.pathname === "/api/users") {
    try {
      const payload = await parseJsonBody(req);
      const state = await readState();
      const incomingUsers = pickUsersPayload(payload);
      const nextUsers = normalizeUsersList(incomingUsers);
      if (!nextUsers.length) {
        sendJson(res, 400, { ok: false, error: "Users payload phải là mảng users hợp lệ" });
        return;
      }
      state.users = nextUsers;
      state.updatedAt = Date.now();
      const saved = await writeState(state);
      sendJson(res, 200, { ok: true, count: saved.users.length });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Save users failed" });
    }
    return;
  }

  if (method === "GET" && url.pathname === "/api/reports") {
    const state = await readState();
    sendJson(res, 200, state.kpiReports || []);
    return;
  }

  if (method === "PUT" && url.pathname === "/api/reports") {
    try {
      const payload = await parseJsonBody(req);
      const state = await readState();
      const incomingReports = pickReportsPayload(payload);
      const nextReports = normalizeKpiReportsList(incomingReports);
      state.kpiReports = nextReports;
      state.updatedAt = Date.now();
      const saved = await writeState(state);
      sendJson(res, 200, { ok: true, count: saved.kpiReports.length });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Save reports failed" });
    }
    return;
  }

  if (method === "POST" && url.pathname === "/api/reports") {
    try {
      const payload = await parseJsonBody(req);
      const state = await readState();
      const incoming = Array.isArray(payload)
        ? pickReportsPayload(payload)
        : (payload && Array.isArray(payload.reports) ? payload.reports : [payload]);
      const normalized = normalizeKpiReportsList(incoming);
      if (!normalized.length) {
        sendJson(res, 400, { ok: false, error: "Reports payload không hợp lệ" });
        return;
      }
      state.kpiReports = [...(state.kpiReports || []), ...normalized].slice(-MAX_REPORTS);
      state.updatedAt = Date.now();
      const saved = await writeState(state);
      sendJson(res, 200, { ok: true, count: saved.kpiReports.length });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Append report failed" });
    }
    return;
  }

  if (method === "POST" && url.pathname === "/api/telegram/config") {
    try {
      const payload = await parseJsonBody(req);
      const state = await readState();
      const channelKey = resolveRequestedChannel(url, payload);
      const channelState = getTelegramChannelState(state, channelKey);
      const token = String(payload.token || "").trim();
      const chatId = String(payload.chatId || "").trim();
      const webhookBaseUrl = String(payload.webhookBaseUrl || "").trim().replace(/\/+$/, "");

      if (!token || !chatId) {
        sendJson(res, 400, { ok: false, error: "Missing token or chatId" });
        return;
      }

      channelState.token = token;
      channelState.chatId = chatId;
      channelState.webhookBaseUrl = webhookBaseUrl;
      channelState.updatedAt = Date.now();
      state.updatedAt = Date.now();
      if (!channelState.webhookSecret) channelState.webhookSecret = randomUUID().replace(/-/g, "");

      let webhookUrl = "";
      let webhookInfo = {};
      if (webhookBaseUrl) {
        webhookUrl = `${webhookBaseUrl}/api/telegram/webhook/${channelKey}/${channelState.webhookSecret}`;
        await setTelegramWebhook(token, webhookUrl);
        webhookInfo = await getWebhookInfo(token);
        channelState.webhookPath = `/api/telegram/webhook/${channelKey}/${channelState.webhookSecret}`;
      } else {
        await deleteTelegramWebhook(token);
        channelState.webhookPath = "";
      }

      const saved = await writeState(state);
      const savedChannel = getTelegramChannelState(saved, channelKey);
      sendJson(res, 200, {
        ok: true,
        channel: channelKey,
        configured: Boolean(savedChannel.token && savedChannel.chatId),
        webhookUrl,
        pendingCount: (savedChannel.reports || []).length,
        webhookInfo
      });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Config failed" });
    }
    return;
  }

  if (method === "GET" && url.pathname === "/api/telegram/pending") {
    try {
      const state = await readState();
      const channelKey = resolveRequestedChannel(url, {});
      const channelState = getTelegramChannelState(state, channelKey);
      const since = Number(url.searchParams.get("since") || 0);
      const includeAll = url.searchParams.get("all") === "1";
      const filtered = includeAll
        ? (channelState.reports || [])
        : (channelState.reports || []).filter((item) => Number(item.receivedAt || 0) > since);
      const lastReceivedAt = filtered.length
        ? Math.max(...filtered.map((item) => Number(item.receivedAt || 0)))
        : since;
      sendJson(res, 200, {
        ok: true,
        channel: channelKey,
        configured: Boolean(channelState.token && channelState.chatId),
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
      const state = await readState();
      const segments = url.pathname.split("/").filter(Boolean);
      const hasExplicitChannel = segments.length >= 5;
      const channelKey = hasExplicitChannel
        ? normalizeTelegramChannelKey(segments[3])
        : DEFAULT_TELEGRAM_CHANNEL;
      if (!SUPPORTED_TELEGRAM_CHANNELS.has(channelKey)) {
        sendJson(res, 404, { ok: false, error: "Unknown telegram channel" });
        return;
      }
      const incomingSecret = hasExplicitChannel ? segments[4] : segments[3];
      const channelState = getTelegramChannelState(state, channelKey);
      if (!incomingSecret || incomingSecret !== channelState.webhookSecret) {
        sendJson(res, 403, { ok: false, error: "Invalid webhook secret" });
        return;
      }

      const update = await parseJsonBody(req);
      channelState.lastUpdateId = Math.max(Number(channelState.lastUpdateId || 0), Number(update?.update_id || 0));
      const result = appendTelegramReport(state, update, channelKey);
      await writeState(state);
      sendJson(res, 200, { ok: true, channel: channelKey, accepted: result.saved, ignored: result.ignored });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Webhook processing failed" });
    }
    return;
  }

  if (method === "GET" && url.pathname === "/api/app-state") {
    try {
      const state = await readState();
      sendJson(res, 200, state.appState || normalizeAppState());
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Read app state failed" });
    }
    return;
  }

  if ((method === "PUT" || method === "POST") && url.pathname === "/api/app-state") {
    try {
      const payload = await parseJsonBody(req);
      const state = await readState();
      state.appState = normalizeAppState(payload);
      state.appState.updatedAt = Date.now();
      state.updatedAt = Date.now();
      const saved = await writeState(state);
      sendJson(res, 200, {
        ok: true,
        updatedAt: saved.appState.updatedAt,
        counts: {
          customers: saved.appState.customers.length,
          schedules: saved.appState.schedules.length,
          inventoryItems: saved.appState.inventoryItems.length,
          inventoryTransactions: saved.appState.inventoryTransactions.length,
          activities: saved.appState.activities.length,
          recycleBin: saved.appState.recycleBin.length,
          newsPosts: saved.appState.newsPosts.length,
          newsEvents: saved.appState.newsEvents.length,
          accountingCashflow: saved.appState.accountingCashflow.length,
          accountingAttendance: saved.appState.accountingAttendance.length,
          reports: saved.appState.reports.length
        }
      });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Save app state failed" });
    }
    return;
  }

  if (method === "GET" && url.pathname === "/api/storage/status") {
    try {
      const state = await readState();
      const payloadSizeBytes = Buffer.byteLength(JSON.stringify(state), "utf8");
      sendJson(res, 200, {
        ok: true,
        persistenceMode: PERSISTENCE_MODE,
        mode: USE_POSTGRES ? "postgres" : "json-file",
        durable: IS_DURABLE_STORAGE,
        renderDiskMounted: IS_RENDER_DISK_ATTACHED,
        stateFile: USE_POSTGRES ? null : STATE_FILE,
        payloadSizeBytes,
        backupDir: BACKUP_DIR,
        backupIntervalMinutes: Math.round(BACKUP_INTERVAL_MS / 60000)
      });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Storage status failed" });
    }
    return;
  }

  sendJson(res, 404, { ok: false, error: "Not found" });
});

async function bootstrap() {
  assertPersistenceConfig();
  await readState();

  server.listen(PORT, HOST, () => {
    console.log(`[telegram-webhook-bridge] listening on http://${HOST}:${PORT}`);
    console.log(`[telegram-webhook-bridge] persistence mode: ${PERSISTENCE_MODE}`);
    console.log(`[telegram-webhook-bridge] storage mode: ${USE_POSTGRES ? "postgres" : "json-file"}`);
    console.log(`[telegram-webhook-bridge] durable storage: ${IS_DURABLE_STORAGE ? "yes" : "no"}`);
  });

  setInterval(() => {
    pollTelegramUpdatesOnce().catch((error) => {
      console.error("[telegram-webhook-bridge] unexpected polling failure:", error.message);
    });
  }, 10000);

  pollTelegramUpdatesOnce().catch((error) => {
    console.error("[telegram-webhook-bridge] initial polling failure:", error.message);
  });

  setInterval(() => {
    writeBackupSnapshot().catch((error) => {
      console.error("[telegram-webhook-bridge] backup interval failure:", error.message);
    });
  }, BACKUP_INTERVAL_MS);
}

bootstrap().catch((error) => {
  console.error("[telegram-webhook-bridge] startup failed:", error.message);
  process.exit(1);
});
