import { createServer } from "node:http";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { randomUUID } from "node:crypto";
import pg from "pg";

const { Pool } = pg;

const PORT = Number(process.env.PORT || process.env.TELEGRAM_WEBHOOK_PORT || 8787);
const BUILD_TS = new Date().toISOString();
const HOST = process.env.TELEGRAM_WEBHOOK_HOST || "0.0.0.0";
const ENV_TELEGRAM_TOKEN = String(process.env.TELEGRAM_TOKEN || "").trim();
const ENV_TELEGRAM_CHAT_ID = String(process.env.TELEGRAM_CHAT_ID || "").trim();
const ENV_WEBHOOK_SECRET = String(process.env.TELEGRAM_WEBHOOK_SECRET || "").trim();
const ENV_WEBHOOK_BASE_URL = String(process.env.TELEGRAM_WEBHOOK_BASE_URL || "").trim().replace(/\/+$/, "");
const RENDER_DISK_ROOT = "/var/data";
const RENDER_DISK_PATH = resolve(RENDER_DISK_ROOT);
const DEFAULT_STATE_FILE = existsSync(RENDER_DISK_ROOT)
  ? resolve(RENDER_DISK_ROOT, "telegram-bridge-state.json")
  : resolve(process.cwd(), "server", "telegram-bridge-state.json");
const STATE_FILE = process.env.STATE_FILE
  ? resolve(process.env.STATE_FILE)
  : DEFAULT_STATE_FILE;
const MAX_REPORTS = 2000;
const MAX_TELEGRAM_DEBUG_EVENTS = 100;
const MAX_USERS = 500;
const DATABASE_URL = String(process.env.DATABASE_URL || "").trim();
const USE_POSTGRES = Boolean(DATABASE_URL);
const PERSISTENCE_MODE = String(process.env.PERSISTENCE_MODE || "auto").trim().toLowerCase();
const PG_STATE_KEY = "global_state";
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
const TELEGRAM_SYNC_MIN_INTERVAL_MS = Math.max(3000, Number(process.env.TELEGRAM_SYNC_MIN_INTERVAL_MS || 8000));
const TELEGRAM_SUSPICIOUS_UPDATE_ID = 100000000;

let dbPool = null;
let dbReadyPromise = null;
let telegramSyncPromise = null;
let telegramSyncRuntime = {
  lastAttemptAt: 0,
  lastSuccessAt: 0,
  lastDurationMs: 0,
  lastError: "",
  lastResult: null
};

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

function normalizeState(raw = {}) {
  const token = String(raw.token || ENV_TELEGRAM_TOKEN || "");
  const chatId = String(raw.chatId || ENV_TELEGRAM_CHAT_ID || "");
  const webhookSecret = String(raw.webhookSecret || ENV_WEBHOOK_SECRET || randomUUID().replace(/-/g, ""));
  const webhookBaseUrl = String(raw.webhookBaseUrl || ENV_WEBHOOK_BASE_URL || "");
  const reports = Array.isArray(raw.reports)
    ? raw.reports.filter((item) => {
        const tags = Array.isArray(item?.raw?.telegramTags)
          ? item.raw.telegramTags.map((tag) => String(tag || "").trim()).filter(Boolean)
          : [];
        const route = String(item?.raw?.telegramRoute || "").trim();
        return tags.length > 0 || Boolean(route);
      }).slice(-MAX_REPORTS)
    : [];
  return {
    token,
    chatId,
    lastUpdateId: Number(raw.lastUpdateId || 0),
    webhookBaseUrl,
    webhookSecret,
    webhookPath: String(raw.webhookPath || (webhookBaseUrl && webhookSecret ? `/api/telegram/webhook/${webhookSecret}` : "")),
    updatedAt: Number(raw.updatedAt || 0),
    reports,
    telegramDebug: normalizeTelegramDebug(raw.telegramDebug),
    users: normalizeUsersList(raw.users && Array.isArray(raw.users) && raw.users.length ? raw.users : DEFAULT_USERS),
    kpiReports: normalizeKpiReportsList(raw.kpiReports),
    appState: normalizeAppState(raw.appState)
  };
}

function normalizeTelegramDebug(input = {}) {
  return {
    acceptedCount: Number(input.acceptedCount || 0),
    ignoredCount: Number(input.ignoredCount || 0),
    duplicateCount: Number(input.duplicateCount || 0),
    parseFailedCount: Number(input.parseFailedCount || 0),
    disallowedChatCount: Number(input.disallowedChatCount || 0),
    emptyMessageCount: Number(input.emptyMessageCount || 0),
    lastAcceptedAt: Number(input.lastAcceptedAt || 0),
    lastIgnoredAt: Number(input.lastIgnoredAt || 0),
    lastReason: String(input.lastReason || ""),
    droppedMessages: Array.isArray(input.droppedMessages) ? input.droppedMessages.slice(0, MAX_TELEGRAM_DEBUG_EVENTS) : []
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
  const rawStatus = String(input.status || "").toLowerCase();
  const normalizedStatus = rawStatus === "resigned" || rawStatus === "suspended" ? "resigned" : "active";
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
    status: normalizedStatus,
    createdAt: Number(input.createdAt) || Date.now(),
    branch: String(input.branch || "HN").toUpperCase() === "HCM" ? "HCM" : "HN",
    resignationDate: String(input.resignationDate || ""),
    resignationReason: String(input.resignationReason || ""),
    employeeGroup: String(input.employeeGroup || ""),
    position: String(input.position || ""),
    dateOfBirth: String(input.dateOfBirth || ""),
    gender: String(input.gender || ""),
    maritalStatus: String(input.maritalStatus || ""),
    ethnicity: String(input.ethnicity || ""),
    religion: String(input.religion || ""),
    identityNumber: String(input.identityNumber || ""),
    identityIssueDate: String(input.identityIssueDate || ""),
    taxCode: String(input.taxCode || ""),
    insuranceNumber: String(input.insuranceNumber || ""),
    permanentAddress: String(input.permanentAddress || ""),
    currentAddress: String(input.currentAddress || ""),
    startDate: String(input.startDate || ""),
    contractType: String(input.contractType || ""),
    compensation: String(input.compensation || ""),
    emergencyContact: String(input.emergencyContact || ""),
    profileSubmissionStatus: String(input.profileSubmissionStatus || ""),
    contractSigningStatus: String(input.contractSigningStatus || ""),
    notes: String(input.notes || "")
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

function getTelegramSyncRuntimeSnapshot() {
  return {
    inFlight: Boolean(telegramSyncPromise),
    cooldownMs: TELEGRAM_SYNC_MIN_INTERVAL_MS,
    lastAttemptAt: Number(telegramSyncRuntime.lastAttemptAt || 0),
    lastSuccessAt: Number(telegramSyncRuntime.lastSuccessAt || 0),
    lastDurationMs: Number(telegramSyncRuntime.lastDurationMs || 0),
    lastError: String(telegramSyncRuntime.lastError || ""),
    lastResult: telegramSyncRuntime.lastResult || null
  };
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
    .replace(/[đ]/g, "d")
    .replace(/\s+/g, "")
    .trim();
}

function normalizeTelegramFieldKey(input) {
  return normalizeVietnamese(
    String(input || "")
      .replace(/^[^\p{L}\p{N}]+/u, "")
      .replace(/[：]/g, ":")
      .replace(/[^\p{L}\p{N}\s]/gu, " ")
      .trim()
  );
}

function normalizeDate(dateText) {
  const value = String(dateText || "").trim();
  if (!value) return new Date().toISOString().slice(0, 10);
  const compact = value.replace(/\s+/g, " ").trim();
  const normalizedSeparator = compact.replace(/\s*([\/-])\s*/g, "$1");
  if (/^\d{4}-\d{2}-\d{2}$/.test(normalizedSeparator)) return normalizedSeparator;
  const ddmmyyyy = normalizedSeparator.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2}|\d{4})$/);
  if (ddmmyyyy) {
    const dd = ddmmyyyy[1].padStart(2, "0");
    const mm = ddmmyyyy[2].padStart(2, "0");
    const yearToken = ddmmyyyy[3];
    const yyyy = yearToken.length === 2 ? `20${yearToken}` : yearToken;
    return `${yyyy}-${mm}-${dd}`;
  }
  const ddmm = normalizedSeparator.match(/^(\d{1,2})[\/-](\d{1,2})$/);
  if (ddmm) {
    const dd = ddmm[1].padStart(2, "0");
    const mm = ddmm[2].padStart(2, "0");
    const yyyy = String(new Date().getFullYear());
    return `${yyyy}-${mm}-${dd}`;
  }
  const parsed = new Date(normalizedSeparator);
  if (Number.isNaN(parsed.getTime())) return new Date().toISOString().slice(0, 10);
  return parsed.toISOString().slice(0, 10);
}

function parseFlexibleNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const text = String(value || "").trim();
  if (!text) return 0;

  const parseSingleAmount = (input) => {
    const sample = String(input || "").trim();
    if (!sample) return 0;
    const tokenMatch = sample.match(/-?\d[\d.,]*\s*k?/i);
    const token = tokenMatch ? tokenMatch[0].trim() : sample;
    const tokenEnd = tokenMatch ? (tokenMatch.index || 0) + tokenMatch[0].length : sample.length;
    const charAfterToken = sample[tokenEnd] || "";
    const hasKSuffix = /k\s*$/i.test(token) && !/[a-z]/i.test(charAfterToken);
    const raw = token.replace(/k\s*$/i, "").trim();
    if (!raw) return 0;

    let normalized;
    const rawDigits = raw.replace(/[^\d.,]/g, "");
    if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(rawDigits)) {
      normalized = raw.replace(/[^\d.,-]/g, "").replace(/\./g, "").replace(/,/g, ".");
    } else if (/^\d{1,3}(,\d{3})+$/.test(rawDigits)) {
      normalized = raw.replace(/[^\d.,-]/g, "").replace(/,/g, "");
    } else {
      normalized = raw.replace(/,/g, ".").replace(/[^\d.\-]/g, "");
    }

    const parsed = Number.parseFloat(normalized);
    if (!Number.isFinite(parsed)) return 0;
    return hasKSuffix ? parsed * 1000 : parsed;
  };

  // Handle expressions like "249k + 150k pp" by summing each amount term.
  const expressionParts = text.split("+").map((part) => part.trim()).filter(Boolean);
  if (expressionParts.length > 1) {
    let sum = 0;
    let parsedAny = false;
    expressionParts.forEach((part) => {
      const valuePart = parseSingleAmount(part);
      if (valuePart > 0) {
        sum += valuePart;
        parsedAny = true;
      }
    });
    if (parsedAny) return sum;
  }

  return parseSingleAmount(text);
}

function normalizeContractCode(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const compact = raw.toUpperCase().replace(/\s+/g, "");
  if (/^NR[0-9A-Z\/-]+$/.test(compact)) return compact;
  if (/^[0-9][0-9A-Z\/-]*$/.test(compact)) return `NR${compact}`;
  return compact;
}

function parseTelegramAllowedChatIds(chatIdValue) {
  return String(chatIdValue || "")
    .split(/[\s,;|]+/)
    .map((id) => id.trim())
    .filter(Boolean);
}

function isTelegramChatAllowed(state, chatId) {
  const allowed = parseTelegramAllowedChatIds(state?.chatId || "");
  if (!allowed.length) return false;
  return allowed.includes(String(chatId || ""));
}

function extractTelegramHashtags(text) {
  const matches = String(text || "").match(/#[^\s#]+/g) || [];
  const unique = new Set();
  matches.forEach((tag) => {
    const normalized = normalizeVietnamese(tag.replace(/^#/, ""));
    if (normalized) unique.add(normalized);
  });
  return Array.from(unique);
}

function detectTelegramRoute(values, hashtags) {
  const all = new Set([...(hashtags || []), ...Object.keys(values || {}).map((key) => normalizeVietnamese(key))]);

  const hasAny = (keys) => keys.some((key) => all.has(normalizeVietnamese(key)));
  if (hasAny(["congno", "cong_no", "conno"])) return "congno";
  if (hasAny(["mkt", "marketing", "mk"])) return "marketing";
  if (hasAny(["tuvan", "tuvan", "tv", "consultant"])) return "consultant";
  if (hasAny(["telesale", "sale", "ts"])) return "telesale";
  if (hasAny(["dieuduong", "dieu_duong", "dd", "nurse", "baocao"])) return "nurse";
  return "";
}

function inferTelegramRoute(values) {
  const keys = new Set(Object.keys(values || {}).map((key) => normalizeVietnamese(key)));
  const hasAny = (list) => list.some((item) => keys.has(normalizeVietnamese(item)));

  if (hasAny(["marketing", "mkt", "marketer", "ngansach", "budget", "adspend", "mess", "luongmess"])) return "marketing";
  if (hasAny(["consultant", "tuvan", "tv"])) return "consultant";
  if (hasAny(["telesale", "sale", "ts"])) return "telesale";
  if (hasAny(["dieuduong", "nurse", "dd"])) return "nurse";

  // Backward compatibility: format cu chi co ten/khach/ngay/gio thi mac dinh dieu duong
  if (hasAny(["ten", "nhanvien"]) && hasAny(["khach", "khachhang", "customer", "tenkhach"])) return "nurse";
  return "";
}

function firstTelegramValue(values, keys) {
  for (const key of keys) {
    const normalizedKey = normalizeVietnamese(key);
    if (values[normalizedKey] !== undefined && values[normalizedKey] !== null && String(values[normalizedKey]).trim() !== "") {
      return values[normalizedKey];
    }
  }
  return "";
}

function firstTelegramValueByKeyRegex(values, regexList = []) {
  for (const [key, value] of Object.entries(values || {})) {
    if (value === undefined || value === null || String(value).trim() === "") continue;
    if (regexList.some((rx) => rx.test(key))) return value;
  }
  return "";
}

const TELEGRAM_FALLBACK_LINE_RULES = [
  { label: "ten dieu duong", key: "tendieuduong" },
  { label: "ten dd", key: "tendd" },
  { label: "ten khach", key: "tenkhach" },
  { label: "ten kh", key: "tenkh" },
  { label: "ma hd", key: "mahd" },
  { label: "so buoi", key: "sobuoi" },
  { label: "khoang cach", key: "khoangcach" },
  { label: "dich vu", key: "dichvu" },
  { label: "trang thai", key: "trangthai" },
  { label: "ghi chu", key: "ghichu" },
  { label: "nhay", key: "ngay" },
  { label: "ngay", key: "ngay" },
  { label: "gio", key: "gio" },
  { label: "bau", key: "dichvu", servicePrefix: "Cham soc bau" },
  { label: "ss", key: "dichvu", servicePrefix: "Cham soc me sau sinh" },
  { label: "tam be", key: "dichvu", servicePrefix: "Tam be" }
];

function canonicalizeTelegramFieldKey(key) {
  const normalized = normalizeTelegramFieldKey(key);
  const aliases = {
    nhay: "ngay",
    ngayy: "ngay",
    bau: "dichvu",
    ss: "dichvu",
    tambe: "dichvu",
    dvu: "dichvu",
    dv: "dichvu",
    tenkh: "tenkhach",
    tenkhh: "tenkhach",
    tennv: "ten",
    kc: "khoangcach"
  };
  const exact = aliases[normalized] || normalized;
  if (exact !== normalized) return exact;

  // Flexible matching for misspelling/no-accent variants.
  if (/^ngay+/.test(normalized) || normalized.includes("ngay")) return "ngay";
  if (/^gio+/.test(normalized) || normalized.includes("gio")) return "gio";
  if (normalized.includes("tenkh") || normalized.includes("khach")) return "tenkhach";
  if (normalized.includes("dichvu") || normalized.includes("dv")) return "dichvu";
  if (normalized.includes("ghichu") || normalized.includes("note") || normalized.includes("noidung")) return "ghichu";
  if (normalized.includes("mahd") || normalized.includes("hopdong") || normalized === "hd") return "mahd";
  if (normalized.includes("khoangcach") || normalized === "km") return "khoangcach";
  return normalized;
}

function parseTelegramLineToField(line) {
  const raw = String(line || "").trim();
  if (!raw) return null;

  const kvMatch = raw.match(/^(.{1,60}?)\s*(?:[:=\-–—|;,]+|\.)\s*(.+)$/u);
  if (kvMatch) {
    const key = canonicalizeTelegramFieldKey(kvMatch[1]);
    const value = kvMatch[2].trim();
    if (!key || !value) return null;
    return { key, value };
  }

  const isLooseTokenMatch = (actual, expected) => {
    const a = String(actual || "");
    const e = String(expected || "");
    if (!a || !e) return false;
    if (a === e) return true;
    if (a.startsWith(e) || e.startsWith(a)) return true;
    if (Math.abs(a.length - e.length) > 1) return false;

    let i = 0;
    let j = 0;
    let edits = 0;
    while (i < a.length && j < e.length) {
      if (a[i] === e[j]) {
        i += 1;
        j += 1;
        continue;
      }
      edits += 1;
      if (edits > 1) return false;
      if (a.length > e.length) i += 1;
      else if (e.length > a.length) j += 1;
      else {
        i += 1;
        j += 1;
      }
    }
    if (i < a.length || j < e.length) edits += 1;
    return edits <= 1;
  };

  const plain = raw
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[đ]/g, "d")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!plain) return null;

  const plainTokens = plain.split(" ");
  const rawTokens = raw.split(/\s+/);
  for (const rule of TELEGRAM_FALLBACK_LINE_RULES) {
    const labelTokens = rule.label.split(" ");
    if (plainTokens.length <= labelTokens.length) continue;
    const matched = labelTokens.every((token, idx) => isLooseTokenMatch(plainTokens[idx], token));
    if (!matched) continue;
    let value = rawTokens.slice(labelTokens.length).join(" ").trim();
    if (!value) return null;
    if (rule.servicePrefix) value = `${rule.servicePrefix} ${value}`.trim();
    return { key: rule.key, value };
  }

  return null;
}

function parseTelegramReportMessage(text) {
  if (!text) return null;

  const lines = String(text).split(/[\r\n]+/).map((line) => line.trim()).filter(Boolean);
  const values = {};
  const hashtags = extractTelegramHashtags(text);

  // Require at least one hashtag so only explicit report messages are recorded.
  if (!hashtags.length) return null;

  for (const line of lines) {
    const parsedField = parseTelegramLineToField(line);
    if (!parsedField) continue;
    values[parsedField.key] = parsedField.value;
  }

  const route = detectTelegramRoute(values, hashtags) || inferTelegramRoute(values);
  if (!route) return null;

  const registrationDateRaw = String(firstTelegramValue(values, ["ngay", "date"]) || "").trim();
  const registrationDate = registrationDateRaw ? normalizeDate(registrationDateRaw) : "";
  const appointmentTime = String(firstTelegramValue(values, ["gio", "time"]) || "").trim();
  const customerName = String(firstTelegramValue(values, ["khach", "khachhang", "customer", "tenkhach", "kh", "hoten", "tenkhachhang", "hotenkhach"]) || "").trim();
  const phone = String(firstTelegramValue(values, ["sdt", "sodienthoai", "phone"]) || "").trim();
  const service = String(firstTelegramValue(values, ["dichvu", "service"]) || "").trim();
  const note = String(firstTelegramValue(values, ["ghichu", "note"]) || "").trim();
  const status = String(firstTelegramValue(values, ["trangthai", "status"]) || "completed").trim() || "completed";
  const explicitShiftMinutes = parseFlexibleNumber(firstTelegramValue(values, ["thoiluong", "phut", "minutes"]));
  const inferredMinutesFromService = Number((service.match(/(\d{2,3})\s*(?:p|phut|phút)?/i) || [])[1] || 0);
  const inferredMinutesFromNote = Number((note.match(/(\d{2,3})\s*(?:p|phut|phút)?/i) || [])[1] || 0);
  const shiftMinutes = explicitShiftMinutes || inferredMinutesFromService || inferredMinutesFromNote || 0;
  const distanceKm = parseFlexibleNumber(firstTelegramValue(values, ["khoangcach", "km", "distance"]));
  const contractAmount = parseFlexibleNumber(firstTelegramValue(values, ["hopdong", "contract", "contractamount", "doanhso"]));

  const base = {
    registrationDate,
    appointmentTime,
    customerName,
    phone,
    service,
    status,
    note,
    shiftMinutes,
    distanceKm,
    contractAmount,
    telegramTags: hashtags,
    telegramRoute: route,
    source: "Telegram Webhook"
  };

  if (route === "nurse") {
    const nurse = String(
      firstTelegramValue(values, ["tenndd", "tendd", "tendieuduong", "ten", "dieuduong", "nurse", "nhanvien", "dd"]) ||
      firstTelegramValueByKeyRegex(values, [/^ten.*dd$/, /^ten.*dieu/, /^dd$/, /^dieuduong$/, /^ten(?!.*kh)/]) ||
      ""
    ).trim();
    const nhomDichVu = String(firstTelegramValue(values, ["dichvu", "service"]) || "").trim();
    const tenkh = String(firstTelegramValue(values, ["tenkhach", "tenkh", "khach", "khachhang", "customer"]) || "").trim();
    const mahd = normalizeContractCode(String(
      firstTelegramValue(values, ["mahd", "mahopdong", "contract"]) ||
      firstTelegramValueByKeyRegex(values, [/^ma.*hd$/, /^ma.*hopdong$/, /^hd$/]) ||
      ""
    ).trim());
    const sobuoi = String(firstTelegramValue(values, ["sobuoi", "buoi"]) || "").trim();
    const khoangcach = String(firstTelegramValue(values, ["khoangcach", "km", "distance"]) || "").trim();
    
    return {
      ...base,
      nurse: nurse || "",
      customerName: tenkh || customerName,
      service: nhomDichVu || service,
      formVersion: nurse && mahd && sobuoi && khoangcach ? "nurse_v2" : null,
      mahd: mahd || "",
      sobuoi: sobuoi || "",
      khoangcach: khoangcach || "",
      source: route ? `Telegram Webhook #${route}` : "Telegram Webhook"
    };
  }

  if (route === "marketing") {
    const marketingName = String(firstTelegramValue(values, [
      "tennvmkt",
      "tennvmarketing",
      "tenmkt",
      "tenmarketing",
      "tennv",
      "marketingname",
      "marketingstaff",
      "marketing",
      "mkt",
      "marketer",
      "nhanvienmkt",
      "nhanvienmarketing",
      "nhanvien",
      "ten"
    ]) || "").trim();
    const customerLooksLikeMarketer = /(mkt|marketing)/i.test(String(customerName || ""));
    const effectiveMarketingName = marketingName || (customerLooksLikeMarketer ? customerName : "");
    const chiphí = parseFlexibleNumber(firstTelegramValue(values, ["chiphi", "chiphí", "chi", "ngansach", "budget"]));
    const mess = parseFlexibleNumber(firstTelegramValue(values, ["mess", "luongmess", "interactions"])) || 0;
    const sdt = parseFlexibleNumber(firstTelegramValue(values, ["sdt", "sodienthoai", "phone", "phones"])) || 0;
    const lich = parseFlexibleNumber(firstTelegramValue(values, ["lich", "datlich", "booked"])) || 0;
    const hopdong = parseFlexibleNumber(firstTelegramValue(values, ["hopdong", "hd", "contract"])) || 0;
    const doanso = parseFlexibleNumber(firstTelegramValue(values, ["doanso", "doanhso", "revenue"])) || 0;
    
    return {
      ...base,
      customerName: customerName || effectiveMarketingName,
      marketingName: effectiveMarketingName,
      marketingStaff: effectiveMarketingName,
      marketingBudget: chiphí || parseFlexibleNumber(firstTelegramValue(values, ["budget", "ads"])),
      marketingMessCount: Math.max(0, mess),
      marketingPhoneCount: Math.max(0, sdt),
      marketingBookedCount: Math.max(0, lich),
      marketingContractCount: Math.max(0, hopdong),
      marketingRevenue: doanso,
      contractAmount: doanso,
      formVersion: marketingName && chiphí > 0 && mess > 0 && sdt > 0 ? "marketing_v2" : null,
      source: route ? `Telegram Marketing #${route}` : "Telegram Webhook"
    };
  }

  if (route === "congno") {
    const consultant = String(firstTelegramValue(values, ["tentv", "ten", "tuvan", "consultant", "tv", "nhanvien"]) || "").trim();
    const tenkh = String(firstTelegramValue(values, ["tenkhach", "tenkh", "khach", "khachhang", "customer"]) || "").trim();
    const mahd = normalizeContractCode(String(firstTelegramValue(values, ["mahd", "mahopdong", "contract"]) || "").trim());
    const thucthu = parseFlexibleNumber(firstTelegramValue(values, ["thucthu", "thuc thu"]));
    const conno = parseFlexibleNumber(firstTelegramValue(values, ["conno", "con no", "conn", "noco", "congno", "cong no"]));
    const pttt = String(firstTelegramValue(values, ["pttt", "phuongthuc", "method"]) || "").trim();
    const ghichu = String(firstTelegramValue(values, ["ghichu", "note", "ghi"]) || "").trim();
    return {
      ...base,
      consultant: consultant || customerName,
      customerName: tenkh || customerName,
      mahd,
      thucthu,
      receivableAmount: conno,
      pttt,
      note: ghichu || note,
      source: `Telegram Cong No #congno`
    };
  }

  if (route === "consultant") {
    const consultant = String(firstTelegramValue(values, ["tentv", "ten", "tuvan", "consultant", "tv", "nhanvien"]) || "").trim();
    const tenkh = String(firstTelegramValue(values, ["tenkhach", "tenkh", "khach", "khachhang", "customer"]) || "").trim();
    const kq = String(firstTelegramValue(values, ["ketqua", "kq", "result"]) || "").trim();
    const mahd = normalizeContractCode(String(firstTelegramValue(values, ["mahd", "mahopdong", "contract"]) || "").trim());
    const giaTriHD = parseFlexibleNumber(firstTelegramValue(values, ["giatrihd", "giatri", "sotien", "so", "amount", "tien"]));
    const thucthu = parseFlexibleNumber(firstTelegramValue(values, ["thucthu", "thuc thu", "thực thu"]));
    const receivableAmount = parseFlexibleNumber(firstTelegramValue(values, ["congno", "cong no", "receivable", "debt"]));
    const sotien = giaTriHD || thucthu;
    const pttt = String(firstTelegramValue(values, ["pttt", "phuongthuc", "method"]) || "").trim();
    const ghichu = String(firstTelegramValue(values, ["ghichu", "note", "ghi"]) || "").trim();
    
    return {
      ...base,
      consultant: consultant || customerName,
      customerName: tenkh || customerName,
      formVersion: consultant && sotien > 0 ? "consultant_v2" : null,
      kq,
      mahd,
      sotien,
      pttt,
      note: ghichu || note,
      contractAmount: giaTriHD || sotien,
      thucthu: thucthu || (giaTriHD && receivableAmount ? giaTriHD - receivableAmount : 0),
      receivableAmount,
      source: route ? `Telegram Tu Van #${route}` : "Telegram Webhook"
    };
  }

  if (route === "telesale") {
    const saleStaff = String(firstTelegramValue(values, ["tennv", "ten", "telesale", "sale", "ts", "nhanvien"]) || "").trim();
    const mess = parseFlexibleNumber(firstTelegramValue(values, ["mess", "luongmess", "interactions"])) || 0;
    const sdt = parseFlexibleNumber(firstTelegramValue(values, ["sdt", "sodienthoai", "phone", "phones"])) || 0;
    const lich = parseFlexibleNumber(firstTelegramValue(values, ["lich", "datlich", "booked"])) || 0;
    const cahoanhuy = parseFlexibleNumber(firstTelegramValue(values, ["cahoanhuy", "hoanhuy", "cancel", "huy", "caho"])) || 0;
    const hopdong = parseFlexibleNumber(firstTelegramValue(values, ["hopdong", "hd", "contract"])) || 0;
    const doanso = parseFlexibleNumber(firstTelegramValue(values, ["doanso", "doanhso", "revenue"])) || 0;
    
    return {
      ...base,
      saleStaff: saleStaff || customerName,
      formVersion: saleStaff && mess > 0 ? "telesale_v2" : null,
      marketingMessCount: Math.max(0, mess),
      marketingPhoneCount: Math.max(0, sdt),
      marketingBookedCount: Math.max(0, lich),
      caCancelled: Math.max(0, cahoanhuy),
      marketingContractCount: Math.max(0, hopdong),
      marketingRevenue: doanso,
      contractAmount: doanso,
      source: route ? `Telegram Telesale #${route}` : "Telegram Webhook"
    };
  }

  return null;
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
      allowed_updates: ["message", "edited_message", "channel_post", "edited_channel_post", "my_chat_member", "chat_member"]
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
  const message = update?.message || update?.edited_message || update?.channel_post || update?.edited_channel_post;
  const chatMeta = extractTelegramChatMeta(update);
  const chatId = String(chatMeta.chatId || "");
  const debug = normalizeTelegramDebug(state.telegramDebug);

  const markIgnored = (reason, counterKey) => {
    const messageContent = String(message?.text || message?.caption || "");
    const event = {
      at: Date.now(),
      reason,
      updateId: String(update?.update_id || ""),
      chatId,
      chatTitle: String(chatMeta.chatTitle || "").trim(),
      updateType: String(chatMeta.updateType || "unknown"),
      textPreview: messageContent.slice(0, 280)
    };
    debug.ignoredCount += 1;
    if (counterKey && typeof debug[counterKey] === "number") debug[counterKey] += 1;
    debug.lastIgnoredAt = event.at;
    debug.lastReason = reason;
    debug.droppedMessages = [event, ...(debug.droppedMessages || [])].slice(0, MAX_TELEGRAM_DEBUG_EVENTS);
    state.telegramDebug = debug;
  };

  const messageContent = String(message?.text || message?.caption || "").trim();
  if (!message || !messageContent) {
    markIgnored("empty_message", "emptyMessageCount");
    return { saved: false, ignored: true, reason: "empty_message" };
  }

  if (!isTelegramChatAllowed(state, chatId)) {
    markIgnored("chat_not_allowed", "disallowedChatCount");
    return { saved: false, ignored: true, reason: "chat_not_allowed" };
  }

  const raw = parseTelegramReportMessage(messageContent);
  if (!raw) {
    markIgnored("parse_failed", "parseFailedCount");
    return { saved: false, ignored: true, reason: "parse_failed" };
  }

  const messageId = String(message?.message_id || "");
  const messageKey = chatId && messageId ? `${chatId}:${messageId}` : "";
  raw.telegramUpdateId = String(update.update_id || `${Date.now()}`);
  raw.telegramMessageId = messageId;
  raw.telegramChatId = chatId;
  raw.telegramChatTitle = String(message?.chat?.title || message?.chat?.username || "").trim();
  raw.telegramUpdateType = String(chatMeta.updateType || "message");
  const reports = Array.isArray(state.reports) ? state.reports : [];

  const processedUpdateIds = new Set(
    reports.map((item) => String(item?.raw?.telegramUpdateId || "")).filter(Boolean)
  );
  if (processedUpdateIds.has(raw.telegramUpdateId)) {
    debug.duplicateCount += 1;
    state.telegramDebug = debug;
    return { saved: false, ignored: false, reason: "duplicate_update" };
  }

  const existingByMessageIndex = messageKey
    ? reports.findIndex((item) => {
        if (String(item?.messageKey || "") === messageKey) return true;
        const itemChatId = String(item?.raw?.telegramChatId || "");
        const itemMessageId = String(item?.raw?.telegramMessageId || "");
        return itemChatId && itemMessageId && `${itemChatId}:${itemMessageId}` === messageKey;
      })
    : -1;

  const fallbackId = String(update.update_id || `${Date.now()}`);
  const report = {
    id: messageKey || fallbackId,
    messageKey,
    receivedAt: Date.now(),
    raw
  };

  if (existingByMessageIndex >= 0) {
    const previous = reports[existingByMessageIndex] || {};
    const merged = {
      ...previous,
      ...report,
      id: String(previous.id || report.id),
      messageKey: messageKey || String(previous.messageKey || ""),
      receivedAt: Date.now(),
      raw: {
        ...(previous.raw || {}),
        ...raw
      }
    };
    const nextReports = reports.slice();
    nextReports[existingByMessageIndex] = merged;
    state.reports = nextReports.slice(-MAX_REPORTS);
    state.updatedAt = Date.now();
    debug.acceptedCount += 1;
    debug.lastAcceptedAt = Date.now();
    state.telegramDebug = debug;
    return { saved: true, ignored: false, updated: true };
  }

  state.reports = [...reports, report].slice(-MAX_REPORTS);
    state.updatedAt = Date.now();
    debug.acceptedCount += 1;
    debug.lastAcceptedAt = Date.now();
    state.telegramDebug = debug;

    // If this is a congno (debt payment) report, patch matching consultant records by mahd
    if (raw.telegramRoute === "congno" && raw.mahd) {
      const targetMahd = normalizeContractCode(raw.mahd).toLowerCase();
      const newReceivable = Number(raw.receivableAmount) || 0;
      let patched = 0;
      state.reports = state.reports.map((item) => {
        if (String(item?.raw?.telegramRoute || "") !== "consultant") return item;
        if (normalizeContractCode(item?.raw?.mahd).toLowerCase() !== targetMahd) return item;
        patched++;
        return { ...item, raw: { ...item.raw, receivableAmount: newReceivable } };
      });
      console.log(`[congno] Patched ${patched} consultant records for mahd=${raw.mahd}, receivable=${newReceivable}`);
    }

    return { saved: true, ignored: false };
}

function extractTelegramChatMeta(update = {}) {
  const message = update?.message || update?.edited_message || update?.channel_post || update?.edited_channel_post;
  const memberUpdate = update?.my_chat_member || update?.chat_member;
  const chat = message?.chat || memberUpdate?.chat || null;
  return {
    chatId: String(chat?.id || ""),
    chatTitle: String(chat?.title || chat?.username || [chat?.first_name, chat?.last_name].filter(Boolean).join(" ") || "").trim(),
    updateType: message
      ? (update?.message ? "message" : update?.edited_message ? "edited_message" : update?.channel_post ? "channel_post" : "edited_channel_post")
      : (update?.my_chat_member ? "my_chat_member" : update?.chat_member ? "chat_member" : "unknown")
  };
}

function getDiscoveredTelegramChats(state) {
  const discovered = new Map();
  const pushChat = (chatId, chatTitle, source, lastSeenAt = 0, updateType = "") => {
    const normalizedChatId = String(chatId || "").trim();
    if (!normalizedChatId) return;
    const current = discovered.get(normalizedChatId) || { chatId: normalizedChatId, chatTitle: "", sources: [], lastSeenAt: 0, updateType: "" };
    current.chatTitle = String(chatTitle || current.chatTitle || "").trim();
    current.lastSeenAt = Math.max(Number(current.lastSeenAt || 0), Number(lastSeenAt || 0));
    if (updateType) current.updateType = String(updateType || current.updateType || "");
    if (source && !current.sources.includes(source)) current.sources.push(source);
    discovered.set(normalizedChatId, current);
  };

  (state.reports || []).forEach((item) => {
    pushChat(
      item?.raw?.telegramChatId,
      item?.raw?.telegramChatTitle,
      String(item?.raw?.source || "report"),
      Number(item?.receivedAt || 0),
      String(item?.raw?.telegramUpdateType || "message")
    );
  });

  const debug = normalizeTelegramDebug(state.telegramDebug);
  (debug.droppedMessages || []).forEach((item) => {
    pushChat(
      item?.chatId,
      item?.chatTitle,
      String(item?.reason || "debug"),
      Number(item?.at || 0),
      String(item?.updateType || "unknown")
    );
  });

  return Array.from(discovered.values()).sort((a, b) => String(a.chatTitle || a.chatId).localeCompare(String(b.chatTitle || b.chatId), "vi"));
}

async function fetchTelegramUpdatesBatch(token, offset) {
  if (!token) return { updates: [], offsetUsed: offset || 0, recoveredOffsetSkew: false };
  const endpoint = `https://api.telegram.org/bot${token}/getUpdates?limit=100${offset ? `&offset=${offset}` : ""}`;
  const response = await fetch(endpoint);
  const payload = await response.json();
  if (!response.ok || payload.ok === false) {
    const error = new Error(payload.description || `getUpdates failed with HTTP ${response.status}`);
    error.code = payload.error_code || response.status;
    throw error;
  }
  return {
    updates: Array.isArray(payload.result) ? payload.result : [],
    offsetUsed: Number(offset || 0),
    recoveredOffsetSkew: false
  };
}

async function fetchTelegramUpdates(state) {
  if (!state.token) return { updates: [], offsetUsed: 0, recoveredOffsetSkew: false };

  try {
    const webhookInfo = await getWebhookInfo(state.token);
    const pendingUpdateCount = Number(webhookInfo?.pending_update_count || 0);
    const currentLastUpdateId = Number(state.lastUpdateId || 0);

    if (pendingUpdateCount > 0 && currentLastUpdateId > TELEGRAM_SUSPICIOUS_UPDATE_ID) {
      const preview = await fetchTelegramUpdatesBatch(state.token);
      const firstUpdateId = Number(preview.updates[0]?.update_id || 0);
      if (firstUpdateId > 0 && firstUpdateId < currentLastUpdateId) {
        return {
          updates: preview.updates,
          offsetUsed: 0,
          recoveredOffsetSkew: true,
          resetLastUpdateIdTo: Math.max(0, firstUpdateId - 1),
          pendingUpdateCount
        };
      }
      return { ...preview, pendingUpdateCount };
    }

    const currentOffset = currentLastUpdateId > 0 ? currentLastUpdateId + 1 : undefined;
    const primary = await fetchTelegramUpdatesBatch(state.token, currentOffset);
    if (primary.updates.length > 0 || !currentOffset || pendingUpdateCount <= 0) {
      return { ...primary, pendingUpdateCount };
    }

    const fallback = await fetchTelegramUpdatesBatch(state.token);
    if (!fallback.updates.length) return { ...primary, pendingUpdateCount };

    const firstUpdateId = Number(fallback.updates[0]?.update_id || 0);
    if (!firstUpdateId || firstUpdateId > currentLastUpdateId) return { ...primary, pendingUpdateCount };

    return {
      updates: fallback.updates,
      offsetUsed: 0,
      recoveredOffsetSkew: true,
      resetLastUpdateIdTo: Math.max(0, firstUpdateId - 1),
      pendingUpdateCount
    };
  } catch (error) {
    console.error("[telegram-webhook-bridge] offset recovery probe failed:", error.message);
    const currentOffset = Number(state.lastUpdateId || 0) > 0 ? Number(state.lastUpdateId) + 1 : undefined;
    const primary = await fetchTelegramUpdatesBatch(state.token, currentOffset);
    return primary;
  }
}

async function pollTelegramUpdatesOnce() {
  const state = await readState();
  if (!state.token || !parseTelegramAllowedChatIds(state.chatId).length) {
    return {
      configured: false,
      fetchedUpdates: 0,
      savedReports: 0,
      updatedReports: 0,
      ignoredUpdates: 0,
      lastUpdateId: Number(state.lastUpdateId || 0)
    };
  }

  try {
    const fetchResult = await fetchTelegramUpdates(state);
    const updates = Array.isArray(fetchResult?.updates) ? fetchResult.updates : [];
    if (!updates.length) {
      return {
        configured: true,
        fetchedUpdates: 0,
        savedReports: 0,
        updatedReports: 0,
        ignoredUpdates: 0,
        lastUpdateId: Number(state.lastUpdateId || 0),
        recoveredOffsetSkew: Boolean(fetchResult?.recoveredOffsetSkew)
      };
    }

    if (Number(fetchResult?.resetLastUpdateIdTo || 0) >= 0 && fetchResult?.recoveredOffsetSkew) {
      state.lastUpdateId = Number(fetchResult.resetLastUpdateIdTo || 0);
    }

    let savedReports = 0;
    let updatedReports = 0;
    let ignoredUpdates = 0;

    updates.forEach((update) => {
      state.lastUpdateId = Math.max(Number(state.lastUpdateId || 0), Number(update.update_id || 0));
      const result = appendTelegramReport(state, update);
      if (result?.saved) savedReports += 1;
      if (result?.updated) updatedReports += 1;
      if (result?.ignored) ignoredUpdates += 1;
    });
    await writeState(state);
    return {
      configured: true,
      fetchedUpdates: updates.length,
      savedReports,
      updatedReports,
      ignoredUpdates,
      lastUpdateId: Number(state.lastUpdateId || 0),
      recoveredOffsetSkew: Boolean(fetchResult?.recoveredOffsetSkew),
      offsetUsed: Number(fetchResult?.offsetUsed || 0),
      pendingUpdateCount: Number(fetchResult?.pendingUpdateCount || 0)
    };
  } catch (error) {
    if (error.code === 409) {
      try {
        const webhookInfo = await getWebhookInfo(state.token);
        if (webhookInfo?.last_error_message || Number(webhookInfo?.pending_update_count || 0) > 0) {
          await deleteTelegramWebhook(state.token);
          state.webhookBaseUrl = "";
          state.webhookPath = "";
          await writeState(state);
        }
      } catch (cleanupError) {
        console.error("[telegram-webhook-bridge] failed to recover from webhook conflict:", cleanupError.message);
      }
      return {
        configured: true,
        fetchedUpdates: 0,
        savedReports: 0,
        updatedReports: 0,
        ignoredUpdates: 0,
        lastUpdateId: Number(state.lastUpdateId || 0),
        recoveredWebhookConflict: true,
        error: error.message
      };
    }
    console.error("[telegram-webhook-bridge] polling error:", error.message);
    return {
      configured: true,
      fetchedUpdates: 0,
      savedReports: 0,
      updatedReports: 0,
      ignoredUpdates: 0,
      lastUpdateId: Number(state.lastUpdateId || 0),
      error: error.message
    };
  }
}

async function syncTelegramUpdatesOnDemand(options = {}) {
  const force = Boolean(options.force);
  const now = Date.now();

  if (telegramSyncPromise) return telegramSyncPromise;

  if (!force && telegramSyncRuntime.lastAttemptAt && now - telegramSyncRuntime.lastAttemptAt < TELEGRAM_SYNC_MIN_INTERVAL_MS) {
    return {
      ok: true,
      skipped: true,
      reason: "cooldown",
      ...getTelegramSyncRuntimeSnapshot()
    };
  }

  telegramSyncRuntime.lastAttemptAt = now;
  telegramSyncRuntime.lastError = "";
  const startedAt = Date.now();

  telegramSyncPromise = (async () => {
    try {
      const result = await pollTelegramUpdatesOnce();
      const durationMs = Date.now() - startedAt;
      telegramSyncRuntime.lastDurationMs = durationMs;
      telegramSyncRuntime.lastResult = result;
      telegramSyncRuntime.lastError = String(result?.error || "");
      if (!result?.error) telegramSyncRuntime.lastSuccessAt = Date.now();
      return {
        ok: !Boolean(result?.error),
        durationMs,
        ...result
      };
    } finally {
      telegramSyncPromise = null;
    }
  })();

  return telegramSyncPromise;
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
    const state = await readState();
    const debug = normalizeTelegramDebug(state.telegramDebug);
    const reports = Array.isArray(state.reports) ? state.reports : [];
    const lastReportAt = reports.length ? Number(reports[reports.length - 1]?.receivedAt || 0) : 0;
    sendJson(res, 200, {
      ok: true,
      service: "telegram-webhook-bridge",
      buildTs: BUILD_TS,
      configured: Boolean(state.token && state.chatId),
      reportsCount: reports.length,
      lastReportAt,
      debug: {
        acceptedCount: debug.acceptedCount,
        ignoredCount: debug.ignoredCount,
        duplicateCount: debug.duplicateCount,
        parseFailedCount: debug.parseFailedCount,
        disallowedChatCount: debug.disallowedChatCount,
        emptyMessageCount: debug.emptyMessageCount,
        lastAcceptedAt: debug.lastAcceptedAt,
        lastIgnoredAt: debug.lastIgnoredAt,
        lastReason: debug.lastReason
      },
      sync: getTelegramSyncRuntimeSnapshot()
    });
    return;
  }

  if (method === "GET" && url.pathname === "/api/telegram/debug") {
    try {
      const state = await readState();
      const debug = normalizeTelegramDebug(state.telegramDebug);
      sendJson(res, 200, {
        ok: true,
        configured: Boolean(state.token && state.chatId),
        chatId: state.chatId,
        reportsCount: Array.isArray(state.reports) ? state.reports.length : 0,
        debug,
        sync: getTelegramSyncRuntimeSnapshot(),
        latestReports: (state.reports || []).slice(-5).map((item) => ({
          id: String(item.id || ""),
          receivedAt: Number(item.receivedAt || 0),
          source: String(item?.raw?.source || ""),
          customerName: String(item?.raw?.customerName || ""),
          telegramRoute: String(item?.raw?.telegramRoute || "")
        }))
      });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Read telegram debug failed" });
    }
    return;
  }

  if (method === "GET" && url.pathname === "/api/telegram/discover-chats") {
    try {
      const state = await readState();
      sendJson(res, 200, {
        ok: true,
        configured: Boolean(state.token),
        currentChatId: String(state.chatId || ""),
        chats: getDiscoveredTelegramChats(state)
      });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Discover chats failed" });
    }
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

      const saved = await writeState(state);
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
      const shouldSync = url.searchParams.get("sync") !== "0";
      const forceSync = url.searchParams.get("forceSync") === "1";
      let sync = null;
      if (shouldSync) {
        sync = await syncTelegramUpdatesOnDemand({ force: forceSync });
      }
      const state = await readState();
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
        sync: sync || getTelegramSyncRuntimeSnapshot(),
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

  if (method === "POST" && url.pathname === "/api/telegram/purge") {
    try {
      const payload = await parseJsonBody(req);
      const targetRoute = normalizeVietnamese(String(payload.route || "").trim());
      const allowedRoutes = new Set(["marketing", "telesale", "consultant", "nurse", "all"]);
      if (!allowedRoutes.has(targetRoute)) {
        sendJson(res, 400, { ok: false, error: "route không hợp lệ. Dùng: marketing|telesale|consultant|nurse|all" });
        return;
      }

      const state = await readState();
      const reports = Array.isArray(state.reports) ? state.reports : [];
      const before = reports.length;

      const filtered = reports.filter((item) => {
        if (targetRoute === "all") return false;
        const route = normalizeVietnamese(String(item?.raw?.telegramRoute || ""));
        const source = normalizeVietnamese(String(item?.raw?.source || ""));
        if (route === targetRoute) return false;
        // Legacy fallback: some old rows only tagged in source text
        if (targetRoute === "marketing" && source.includes("marketing")) return false;
        if (targetRoute === "telesale" && source.includes("telesale")) return false;
        if (targetRoute === "consultant" && source.includes("tu van")) return false;
        if (targetRoute === "nurse" && source.includes("nurse")) return false;
        return true;
      });

      const removed = before - filtered.length;
      state.reports = filtered;
      state.updatedAt = Date.now();
      await writeState(state);

      sendJson(res, 200, {
        ok: true,
        route: targetRoute,
        removed,
        remaining: filtered.length
      });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Purge telegram failed" });
    }
    return;
  }

  if (method === "POST" && url.pathname === "/api/telegram/admin/fix-small-amounts") {
    // One-time migration: multiply amount fields < threshold by 1000 for a given route
    try {
      const payload = await parseJsonBody(req);
      const route = String(payload.route || "consultant");
      const threshold = Number(payload.threshold ?? 10000);
      const amountFields = Array.isArray(payload.fields)
        ? payload.fields
        : ["sotien", "contractAmount", "thanhtoan", "congno"];
      const state = await readState();
      const reports = Array.isArray(state.reports) ? state.reports : [];
      let patchedCount = 0;
      const log = [];
      const nextReports = reports.map((item) => {
        const raw = item?.raw || {};
        if (raw.telegramRoute !== route) return item;
        const needsPatch = amountFields.some((f) => {
          const v = Number(raw[f] ?? item[f]);
          return Number.isFinite(v) && v > 0 && v < threshold;
        });
        if (!needsPatch) return item;
        const newRaw = { ...raw };
        const newItem = { ...item };
        const entry = { id: item.id, date: raw.registrationDate, name: raw.consultant, before: {}, after: {} };
        for (const f of amountFields) {
          const v = Number(raw[f] ?? item[f]);
          if (Number.isFinite(v) && v > 0 && v < threshold) {
            newRaw[f] = v * 1000;
            newItem[f] = v * 1000;
            entry.before[f] = v;
            entry.after[f] = v * 1000;
          }
        }
        newItem.raw = newRaw;
        log.push(entry);
        patchedCount++;
        return newItem;
      });
      state.reports = nextReports;
      state.updatedAt = Date.now();
      await writeState(state);
      sendJson(res, 200, { ok: true, route, threshold, patchedCount, log });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message });
    }
    return;
  }

  if (method === "POST" && url.pathname === "/api/telegram/admin/patch-amounts") {
    // Targeted correction for specific telegram updates.
    try {
      const payload = await parseJsonBody(req);
      const patches = Array.isArray(payload?.patches) ? payload.patches : [];
      if (!patches.length) {
        sendJson(res, 400, { ok: false, error: "Thiếu patches" });
        return;
      }

      const patchMap = new Map();
      patches.forEach((item) => {
        const updateId = String(item?.updateId || "").trim();
        const values = item?.values && typeof item.values === "object" ? item.values : null;
        if (!updateId || !values) return;
        patchMap.set(updateId, values);
      });

      if (!patchMap.size) {
        sendJson(res, 400, { ok: false, error: "patches không hợp lệ" });
        return;
      }

      const state = await readState();
      const reports = Array.isArray(state.reports) ? state.reports : [];
      let patchedCount = 0;
      const log = [];

      const nextReports = reports.map((item) => {
        const raw = item?.raw || {};
        const messageKey = String(item?.id || "").trim();
        const rawUpdateId = String(raw.telegramUpdateId || "").trim();
        const legacyMessageKey = [String(raw.telegramChatId || "").trim(), String(raw.telegramMessageId || "").trim()].filter(Boolean).join(":");
        const matchedKey = [messageKey, rawUpdateId, legacyMessageKey].find((key) => key && patchMap.has(key)) || "";
        const values = matchedKey ? patchMap.get(matchedKey) : null;
        if (!values) return item;

        const newRaw = { ...raw };
        const newItem = { ...item };
        const before = {};
        const after = {};

        Object.entries(values).forEach(([field, value]) => {
          const numericValue = Number(value);
          if (!Number.isFinite(numericValue)) return;
          before[field] = Number(raw[field] ?? item[field] ?? 0);
          newRaw[field] = numericValue;
          newItem[field] = numericValue;
          after[field] = numericValue;
        });

        newItem.raw = newRaw;
        patchedCount += 1;
        log.push({
          updateId: matchedKey,
          date: raw.registrationDate,
          name: raw.consultant || raw.marketingName || raw.saleStaff || raw.nurse || "",
          before,
          after
        });
        return newItem;
      });

      state.reports = nextReports;
      state.updatedAt = Date.now();
      await writeState(state);
      sendJson(res, 200, { ok: true, patchedCount, log });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Patch amounts failed" });
    }
    return;
  }

  if (method === "POST" && url.pathname === "/api/telegram/admin/reparse-dropped-previews") {
    // Recovery helper: try parsing stored dropped message previews again with latest parser rules.
    try {
      const payload = await parseJsonBody(req);
      const onlyParseFailed = payload?.onlyParseFailed !== false;
      const limit = Math.max(1, Math.min(500, Number(payload?.limit || 200)));

      const state = await readState();
      const debug = normalizeTelegramDebug(state.telegramDebug);
      const dropped = Array.isArray(debug.droppedMessages) ? debug.droppedMessages : [];
      const existingUpdateIds = new Set(
        (Array.isArray(state.reports) ? state.reports : [])
          .map((item) => String(item?.raw?.telegramUpdateId || "").trim())
          .filter(Boolean)
      );

      let scanned = 0;
      let recovered = 0;
      let skippedDuplicate = 0;
      let skippedUnparsed = 0;

      const nextReports = Array.isArray(state.reports) ? state.reports.slice() : [];

      for (const event of dropped) {
        if (scanned >= limit) break;
        scanned += 1;
        const reason = String(event?.reason || "");
        if (onlyParseFailed && reason !== "parse_failed") continue;

        const updateId = String(event?.updateId || "").trim();
        if (updateId && existingUpdateIds.has(updateId)) {
          skippedDuplicate += 1;
          continue;
        }

        const text = String(event?.textPreview || "").trim();
        if (!text) {
          skippedUnparsed += 1;
          continue;
        }

        const raw = parseTelegramReportMessage(text);
        if (!raw) {
          skippedUnparsed += 1;
          continue;
        }

        raw.telegramUpdateId = updateId || `reparse-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
        raw.telegramMessageId = "";
        raw.telegramChatId = String(event?.chatId || "").trim();
        raw.telegramChatTitle = String(event?.chatTitle || "").trim();
        raw.telegramUpdateType = String(event?.updateType || "message");
        raw.source = `${String(raw.source || "Telegram Webhook")} (reparsed)`;

        const id = raw.telegramChatId && raw.telegramMessageId
          ? `${raw.telegramChatId}:${raw.telegramMessageId}`
          : `reparse:${raw.telegramUpdateId}`;

        nextReports.push({
          id,
          messageKey: "",
          receivedAt: Date.now(),
          raw
        });

        existingUpdateIds.add(raw.telegramUpdateId);
        recovered += 1;
      }

      state.reports = nextReports.slice(-MAX_REPORTS);
      state.updatedAt = Date.now();
      await writeState(state);

      sendJson(res, 200, {
        ok: true,
        scanned,
        recovered,
        skippedDuplicate,
        skippedUnparsed,
        totalReports: state.reports.length
      });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "Reparse dropped previews failed" });
    }
    return;
  }

  if (method === "POST" && url.pathname === "/api/telegram/admin/fix-date-swap") {
    // Migration: fix records where DD/MM/YYYY with extra whitespace was parsed by JS Date as MM/DD/YYYY,
    // resulting in month and day being swapped.
    try {
      const payload = await parseJsonBody(req).catch(() => ({}));
      const inferredMonth = Number(new Date().toLocaleString("en-US", { timeZone: "Asia/Ho_Chi_Minh", month: "numeric" }));
      const targetMonthInput = Number(payload?.targetMonth || inferredMonth || 5);
      const targetMonth = Math.max(1, Math.min(12, Number.isFinite(targetMonthInput) ? targetMonthInput : 5));

      const state = await readState();
      const reports = Array.isArray(state.reports) ? state.reports : [];
      const fixed = [];
      const skipped = [];

      for (const item of reports) {
        const raw = item?.raw;
        if (!raw || raw.telegramRoute !== "nurse") continue;
        const d = raw.registrationDate;
        if (!d) continue;
        // Expect YYYY-MM-DD
        const m = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (!m) continue;
        const [, yyyy, mm, dd] = m;
        const month = Number(mm);
        const day = Number(dd);
        // Swapped pattern for DD/MM/YYYY interpreted as MM/DD/YYYY:
        // month = original day, day = original month. For current-month reports, day tends to be targetMonth.
        const likelySwapped = month !== targetMonth && day === targetMonth && month >= 1 && month <= 12;
        if (likelySwapped) {
          const corrected = `${yyyy}-${dd}-${mm}`;
          fixed.push({ old: d, new: corrected, nurse: raw.nurse, customer: raw.customerName });
          raw.registrationDate = corrected;
        } else {
          skipped.push(d);
        }
      }

      if (fixed.length > 0) {
        state.reports = reports;
        state.updatedAt = Date.now();
        await writeState(state);
      }

      sendJson(res, 200, {
        ok: true,
        targetMonth,
        fixed: fixed.length,
        skippedCount: skipped.length,
        details: fixed
      });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message || "fix-date-swap failed" });
    }
    return;
  }

  if (method === "POST" && url.pathname === "/api/telegram/test-parse") {
    try {
      const payload = await parseJsonBody(req);
      const text = String(payload.text || "");
      const raw = parseTelegramReportMessage(text);
      sendJson(res, 200, {
        ok: true,
        inputText: text.slice(0, 200),
        parsed: raw ? Object.keys(raw).length : 0,
        result: raw || { error: "parse returned null" }
      });
    } catch (error) {
      sendJson(res, 500, { ok: false, error: error.message });
    }
    return;
  }

  if (method === "POST" && url.pathname.startsWith("/api/telegram/webhook/")) {
    try {
      const state = await readState();
      const incomingSecret = url.pathname.split("/").pop() || "";
      if (!incomingSecret || incomingSecret !== state.webhookSecret) {
        sendJson(res, 403, { ok: false, error: "Invalid webhook secret" });
        return;
      }

      const update = await parseJsonBody(req);
      state.lastUpdateId = Math.max(Number(state.lastUpdateId || 0), Number(update?.update_id || 0));
      const result = appendTelegramReport(state, update);
      await writeState(state);
      sendJson(res, 200, { ok: true, accepted: result.saved, ignored: result.ignored });
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
      const incomingAppState = normalizeAppState(payload);
      const existingAppState = normalizeAppState(state.appState || {});

      // Defensive merge: avoid losing schedules when a stale browser overwrites app-state.
      const mergeSchedules = (existingList = [], incomingList = []) => {
        const byId = new Map();
        const extras = [];

        existingList.forEach((item) => {
          const id = String(item?.id || "").trim();
          if (!id) {
            extras.push(item);
            return;
          }
          byId.set(id, item);
        });

        incomingList.forEach((item) => {
          const id = String(item?.id || "").trim();
          if (!id) {
            extras.push(item);
            return;
          }
          byId.set(id, { ...(byId.get(id) || {}), ...item, id });
        });

        return [...byId.values(), ...extras];
      };

      incomingAppState.schedules = mergeSchedules(existingAppState.schedules, incomingAppState.schedules);
      state.appState = incomingAppState;
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
  const state = await readState();

  // Auto-register webhook on startup if ENV vars are set and webhook not yet registered
  if (state.token && state.webhookBaseUrl && !state.webhookPath) {
    try {
      const webhookUrl = `${state.webhookBaseUrl}/api/telegram/webhook/${state.webhookSecret}`;
      await setTelegramWebhook(state.token, webhookUrl);
      state.webhookPath = `/api/telegram/webhook/${state.webhookSecret}`;
      await writeState(state);
      console.log(`[telegram-webhook-bridge] webhook registered: ${webhookUrl}`);
    } catch (e) {
      console.error("[telegram-webhook-bridge] webhook registration failed:", e.message);
    }
  } else if (state.token && state.webhookPath) {
    console.log(`[telegram-webhook-bridge] webhook already set: ${state.webhookBaseUrl}${state.webhookPath}`);
  }

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
// Updated at Thu Apr 23 15:27:42 +07 2026
