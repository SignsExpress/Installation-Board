const express = require("express");
const cors = require("cors");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const {
  bootstrapPasswordsFromEnv,
  ensureUsersFile,
  readUsersStore,
  sanitizeUser,
  verifyPassword
} = require("./auth-store");

const PORT = Number(process.env.PORT || 3030);
const HOST = process.env.HOST || "0.0.0.0";
const DEV_FRONTEND_PORT = Number(process.env.DEV_FRONTEND_PORT || 5173);
const DIST_DIR = path.join(__dirname, "..", "dist");
const DEFAULT_DATA_FILE = path.join(__dirname, "..", "data", "jobs.json");
const DEFAULT_INSTALLERS_FILE = path.join(__dirname, "..", "data", "installers-live.json");
const DEFAULT_REQUESTS_FILE = path.join(__dirname, "..", "data", "requests.json");
const LEGACY_INSTALLER_DIRECTORY = "/var/data/sx-installer-directory";
const LEGACY_INSTALLERS_FILE = path.join(LEGACY_INSTALLER_DIRECTORY, "installers.json");
const LEGACY_REQUESTS_FILE = path.join(LEGACY_INSTALLER_DIRECTORY, "requests.json");
const PUBLIC_DIR = path.join(__dirname, "..", "public");
const TIME_ZONE = "Europe/London";
const streamClients = new Set();
const DEFAULT_COREBRIDGE_BASE_URL = "https://corebridgev3.azure-api.net";
const DEFAULT_COREBRIDGE_ORDER_PATH = "/core/api/order";
const SESSION_COOKIE_NAME = "installation_board_session";
const SESSION_TTL_MS = 1000 * 60 * 60 * 24 * 30;
const sessions = new Map();

const weekdayFormatter = new Intl.DateTimeFormat("en-GB", { weekday: "short", timeZone: TIME_ZONE });
const longDateFormatter = new Intl.DateTimeFormat("en-GB", {
  weekday: "long",
  day: "numeric",
  month: "long",
  year: "numeric",
  timeZone: TIME_ZONE
});
const monthFormatter = new Intl.DateTimeFormat("en-GB", { month: "long", year: "numeric", timeZone: TIME_ZONE });
const dayFormatter = new Intl.DateTimeFormat("en-CA", {
  year: "numeric",
  month: "2-digit",
  day: "2-digit",
  timeZone: TIME_ZONE
});

function getDataFile() {
  return process.env.DATA_FILE || DEFAULT_DATA_FILE;
}

function getInstallersFile() {
  if (process.env.INSTALLERS_FILE) {
    return process.env.INSTALLERS_FILE;
  }

  if (fs.existsSync(LEGACY_INSTALLERS_FILE)) {
    return LEGACY_INSTALLERS_FILE;
  }

  if (process.env.DATA_FILE) {
    return path.join(path.dirname(process.env.DATA_FILE), "installers.json");
  }

  return DEFAULT_INSTALLERS_FILE;
}

function getRequestsFile() {
  if (process.env.REQUESTS_FILE) {
    return process.env.REQUESTS_FILE;
  }

  if (fs.existsSync(LEGACY_REQUESTS_FILE)) {
    return LEGACY_REQUESTS_FILE;
  }

  if (process.env.DATA_FILE) {
    return path.join(path.dirname(process.env.DATA_FILE), "requests.json");
  }

  return DEFAULT_REQUESTS_FILE;
}

function parseCookies(headerValue) {
  return String(headerValue || "")
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean)
    .reduce((accumulator, pair) => {
      const separator = pair.indexOf("=");
      if (separator === -1) return accumulator;
      const key = decodeURIComponent(pair.slice(0, separator).trim());
      const value = decodeURIComponent(pair.slice(separator + 1).trim());
      accumulator[key] = value;
      return accumulator;
    }, {});
}

function serializeSessionCookie(sessionId, { expiresAt, clear = false } = {}) {
  const parts = [
    `${SESSION_COOKIE_NAME}=${clear ? "" : encodeURIComponent(sessionId)}`,
    "HttpOnly",
    "Path=/",
    "SameSite=Lax"
  ];

  if (process.env.NODE_ENV === "production") {
    parts.push("Secure");
  }

  if (clear) {
    parts.push("Max-Age=0");
    parts.push("Expires=Thu, 01 Jan 1970 00:00:00 GMT");
  } else if (expiresAt) {
    parts.push(`Expires=${new Date(expiresAt).toUTCString()}`);
  }

  return parts.join("; ");
}

function createSession(user) {
  const sessionId = crypto.randomUUID();
  const expiresAt = Date.now() + SESSION_TTL_MS;
  sessions.set(sessionId, {
    user,
    expiresAt
  });
  return { sessionId, expiresAt };
}

function getSessionFromRequest(request) {
  const cookies = parseCookies(request.headers.cookie);
  const sessionId = cookies[SESSION_COOKIE_NAME];
  if (!sessionId) return null;

  const session = sessions.get(sessionId);
  if (!session) return null;
  if (session.expiresAt <= Date.now()) {
    sessions.delete(sessionId);
    return null;
  }

  return { sessionId, ...session };
}

function clearExpiredSessions() {
  const now = Date.now();
  for (const [sessionId, session] of sessions.entries()) {
    if (session.expiresAt <= now) {
      sessions.delete(sessionId);
    }
  }
}

function requireHost(request, response) {
  if (request.user?.role === "host") return true;
  response.status(403).json({ error: "Host access required." });
  return false;
}

function ensureStoreFile() {
  const file = getDataFile();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  if (!fs.existsSync(file)) {
    fs.writeFileSync(file, `${JSON.stringify({ jobs: [], holidays: [] }, null, 2)}\n`, "utf8");
  }
}

function ensureInstallersFile() {
  const file = getInstallersFile();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  if (!fs.existsSync(file)) {
    const seedInstallers = fs.existsSync(DEFAULT_INSTALLERS_FILE)
      ? fs.readFileSync(DEFAULT_INSTALLERS_FILE, "utf8")
      : "[]\n";
    fs.writeFileSync(file, seedInstallers, "utf8");
  }
}

function ensureRequestsFile() {
  const file = getRequestsFile();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  if (!fs.existsSync(file)) {
    const seedRequests = fs.existsSync(DEFAULT_REQUESTS_FILE)
      ? fs.readFileSync(DEFAULT_REQUESTS_FILE, "utf8")
      : "[]\n";
    fs.writeFileSync(file, seedRequests, "utf8");
  }
}

async function readStore() {
  ensureStoreFile();
  const raw = await fsp.readFile(getDataFile(), "utf8");

  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return { jobs: parsed, holidays: [] };
    }
    return {
      jobs: Array.isArray(parsed.jobs) ? parsed.jobs : [],
      holidays: Array.isArray(parsed.holidays) ? parsed.holidays : []
    };
  } catch (error) {
    console.error("Invalid board store JSON, returning empty store.", error);
    return { jobs: [], holidays: [] };
  }
}

async function writeStore(store) {
  ensureStoreFile();
  const nextStore = {
    jobs: [...store.jobs].sort((left, right) => {
      if (left.date !== right.date) return left.date.localeCompare(right.date);
      return String(left.customerName || "").localeCompare(String(right.customerName || ""));
    }),
    holidays: [...store.holidays].sort((left, right) => {
      if (left.date !== right.date) return left.date.localeCompare(right.date);
      return String(left.person || "").localeCompare(String(right.person || ""));
    })
  };
  await fsp.writeFile(getDataFile(), `${JSON.stringify(nextStore, null, 2)}\n`, "utf8");
  return nextStore;
}

async function readInstallersStore() {
  ensureInstallersFile();
  const installersFile = getInstallersFile();
  const raw = await fsp.readFile(installersFile, "utf8");

  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed) && installersFile !== DEFAULT_INSTALLERS_FILE && fs.existsSync(DEFAULT_INSTALLERS_FILE)) {
      const seedRaw = await fsp.readFile(DEFAULT_INSTALLERS_FILE, "utf8");
      const seedParsed = JSON.parse(seedRaw);

      if (Array.isArray(seedParsed) && seedParsed.length > 0) {
        const isSubsetOfSeed =
          parsed.length === 0 ||
          parsed.every((item) => seedParsed.some((seedItem) => String(seedItem?.id || "") === String(item?.id || "")));

        if (isSubsetOfSeed && parsed.length < seedParsed.length) {
          await fsp.writeFile(installersFile, `${JSON.stringify(seedParsed, null, 2)}\n`, "utf8");
          return seedParsed;
        }

        if (parsed.length === 0) {
          await fsp.writeFile(installersFile, `${JSON.stringify(seedParsed, null, 2)}\n`, "utf8");
          return seedParsed;
        }
      }
    }

    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    console.error("Invalid installers store JSON, returning empty list.", error);
    return [];
  }
}

async function writeInstallersStore(installers) {
  ensureInstallersFile();
  const nextInstallers = [...installers].sort((left, right) =>
    String(left.name || "").localeCompare(String(right.name || ""))
  );
  await fsp.writeFile(getInstallersFile(), `${JSON.stringify(nextInstallers, null, 2)}\n`, "utf8");
  return nextInstallers;
}

async function readRequestsStore() {
  ensureRequestsFile();
  const raw = await fsp.readFile(getRequestsFile(), "utf8");

  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    console.error("Invalid requests store JSON, returning empty list.", error);
    return [];
  }
}

async function writeRequestsStore(requests) {
  ensureRequestsFile();
  const nextRequests = [...requests].sort((left, right) =>
    String(right.createdAt || "").localeCompare(String(left.createdAt || ""))
  );
  await fsp.writeFile(getRequestsFile(), `${JSON.stringify(nextRequests, null, 2)}\n`, "utf8");
  return nextRequests;
}

async function inspectJsonArrayFile(file) {
  const exists = fs.existsSync(file);
  if (!exists) {
    return { file, exists: false, size: 0, count: 0 };
  }

  const raw = await fsp.readFile(file, "utf8");
  try {
    const parsed = JSON.parse(raw);
    return {
      file,
      exists: true,
      size: Buffer.byteLength(raw, "utf8"),
      count: Array.isArray(parsed) ? parsed.length : 0
    };
  } catch (error) {
    return {
      file,
      exists: true,
      size: Buffer.byteLength(raw, "utf8"),
      count: 0,
      invalidJson: true
    };
  }
}

function makeId() {
  return crypto.randomUUID();
}

function parseIsoDate(isoDate) {
  const [year, month, day] = String(isoDate || "").split("-").map(Number);
  if (!year || !month || !day) return null;
  return new Date(Date.UTC(year, month - 1, day));
}

function toIsoDate(date) {
  return dayFormatter.format(date);
}

function addDays(date, days) {
  const next = new Date(date);
  next.setUTCDate(next.getUTCDate() + days);
  return next;
}

function getTodayInLondon() {
  return parseIsoDate(dayFormatter.format(new Date()));
}

function getStartOfWeek(date) {
  const weekday = date.getUTCDay();
  const diff = weekday === 0 ? -6 : 1 - weekday;
  return addDays(date, diff);
}

function getRollingWindow(today = getTodayInLondon()) {
  const currentWeekStart = getStartOfWeek(today);
  return {
    start: addDays(currentWeekStart, -7),
    end: addDays(currentWeekStart, 19)
  };
}

function getWeekdaysInRange(start, end) {
  const dates = [];
  let cursor = start;
  while (cursor <= end) {
    const day = cursor.getUTCDay();
    if (day >= 1 && day <= 5) dates.push(cursor);
    cursor = addDays(cursor, 1);
  }
  return dates;
}

function formatWeekLabel(date) {
  const weekEnd = addDays(date, 4);
  const startLabel = `${date.getUTCDate()} ${monthFormatter.format(date).split(" ")[0]}`;
  const endLabel = `${weekEnd.getUTCDate()} ${monthFormatter.format(weekEnd).split(" ")[0]}`;
  return `${startLabel} to ${endLabel}`;
}

function getEasterSunday(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(Date.UTC(year, month - 1, day));
}

function nthWeekdayOfMonth(year, month, weekday, occurrence) {
  const first = new Date(Date.UTC(year, month, 1));
  const firstWeekday = first.getUTCDay();
  const delta = (weekday - firstWeekday + 7) % 7;
  return new Date(Date.UTC(year, month, 1 + delta + (occurrence - 1) * 7));
}

function lastWeekdayOfMonth(year, month, weekday) {
  const last = new Date(Date.UTC(year, month + 1, 0));
  const lastWeekday = last.getUTCDay();
  const delta = (lastWeekday - weekday + 7) % 7;
  return new Date(Date.UTC(year, month + 1, last.getUTCDate() - delta));
}

function applySubstituteHoliday(date) {
  const weekday = date.getUTCDay();
  if (weekday === 6) return addDays(date, 2);
  if (weekday === 0) return addDays(date, 1);
  return date;
}

function getUkBankHolidays(year) {
  const easterSunday = getEasterSunday(year);
  const goodFriday = addDays(easterSunday, -2);
  const easterMonday = addDays(easterSunday, 1);
  const newYearsDay = applySubstituteHoliday(new Date(Date.UTC(year, 0, 1)));
  const earlyMay = nthWeekdayOfMonth(year, 4, 1, 1);
  const springBank = lastWeekdayOfMonth(year, 4, 1);
  const summerBank = lastWeekdayOfMonth(year, 7, 1);

  let christmasDayObserved = new Date(Date.UTC(year, 11, 25));
  let boxingDayObserved = new Date(Date.UTC(year, 11, 26));
  const christmasWeekday = christmasDayObserved.getUTCDay();
  const boxingWeekday = boxingDayObserved.getUTCDay();

  if (christmasWeekday === 6) {
    christmasDayObserved = new Date(Date.UTC(year, 11, 27));
    boxingDayObserved = new Date(Date.UTC(year, 11, 28));
  } else if (christmasWeekday === 0) {
    christmasDayObserved = new Date(Date.UTC(year, 11, 27));
    boxingDayObserved = new Date(Date.UTC(year, 11, 26));
  } else if (boxingWeekday === 6 || boxingWeekday === 0) {
    boxingDayObserved = new Date(Date.UTC(year, 11, 28));
  }

  return [
    { date: toIsoDate(newYearsDay), label: "New Year's Day" },
    { date: toIsoDate(goodFriday), label: "Good Friday" },
    { date: toIsoDate(easterMonday), label: "Easter Monday" },
    { date: toIsoDate(earlyMay), label: "Early May Bank Holiday" },
    { date: toIsoDate(springBank), label: "Spring Bank Holiday" },
    { date: toIsoDate(summerBank), label: "Summer Bank Holiday" },
    { date: toIsoDate(christmasDayObserved), label: "Christmas Day" },
    { date: toIsoDate(boxingDayObserved), label: "Boxing Day" }
  ];
}

function getHolidayMap(start, end) {
  const years = new Set([start.getUTCFullYear(), end.getUTCFullYear()]);
  const map = new Map();
  for (const year of years) {
    getUkBankHolidays(year).forEach((holiday) => {
      map.set(holiday.date, holiday.label);
    });
  }
  return map;
}

function buildBoardRows(jobs, staffHolidays, today = getTodayInLondon()) {
  const { start, end } = getRollingWindow(today);
  const weekdayDates = getWeekdaysInRange(start, end);
  const holidayMap = getHolidayMap(start, end);
  const todayIso = toIsoDate(today);

  const jobsByDate = jobs.reduce((map, job) => {
    const existing = map.get(job.date) || [];
    existing.push(job);
    map.set(job.date, existing);
    return map;
  }, new Map());

  const staffHolidaysByDate = staffHolidays.reduce((map, entry) => {
    const existing = map.get(entry.date) || [];
    existing.push(entry);
    map.set(entry.date, existing);
    return map;
  }, new Map());

  const rows = weekdayDates.map((date) => {
    const isoDate = toIsoDate(date);
    const sameDayJobs = (jobsByDate.get(isoDate) || []).sort((left, right) =>
      String(left.customerName || "").localeCompare(String(right.customerName || ""))
    );
    const sameDayStaffHolidays = (staffHolidaysByDate.get(isoDate) || []).sort((left, right) =>
      String(left.person || "").localeCompare(String(right.person || ""))
    );

    return {
      isoDate,
      dayLabel: weekdayFormatter.format(date).toUpperCase(),
      dayNumber: String(date.getUTCDate()).padStart(2, "0"),
      fullDateLabel: longDateFormatter.format(date),
      bankHoliday: holidayMap.get(isoDate) || "",
      staffHolidays: sameDayStaffHolidays,
      isToday: isoDate === todayIso,
      isPast: isoDate < todayIso,
      jobs: sameDayJobs
    };
  });

  const weeks = [];
  for (let index = 0; index < rows.length; index += 5) {
    const slice = rows.slice(index, index + 5);
    if (slice.length) {
      weeks.push({
        id: slice[0].isoDate,
        label: formatWeekLabel(parseIsoDate(slice[0].isoDate)),
        rows: slice
      });
    }
  }

  return {
    generatedAt: new Date().toISOString(),
    today: todayIso,
    start: toIsoDate(start),
    end: toIsoDate(end),
    weeks
  };
}

function sanitizeJob(payload) {
  const rawInstallers = Array.isArray(payload.installers)
    ? payload.installers.map(String)
    : typeof payload.installers === "string" && payload.installers.trim()
      ? payload.installers.split(/[,/]+/).map((item) => item.trim()).filter(Boolean)
      : [];

  return {
    id: String(payload.id || makeId()),
    date: String(payload.date || "").trim(),
    orderReference: String(payload.orderReference || "").trim(),
    customerName: String(payload.customerName || "").trim(),
    description: String(payload.description || "").trim(),
    contact: String(payload.contact || "").trim(),
    number: String(payload.number || "").trim(),
    address: String(payload.address || "").trim(),
    installers: rawInstallers,
    customInstaller: String(payload.customInstaller || "").trim(),
    jobType: String(payload.jobType || "Install").trim(),
    customJobType: String(payload.customJobType || "").trim(),
    notes: String(payload.notes || "").trim(),
    createdAt: String(payload.createdAt || new Date().toISOString()),
    updatedAt: new Date().toISOString()
  };
}

function sanitizeStaffHoliday(payload) {
  return {
    id: String(payload.id || makeId()),
    date: String(payload.date || "").trim(),
    person: String(payload.person || "").trim(),
    duration: String(payload.duration || "Full Day").trim(),
    createdAt: String(payload.createdAt || new Date().toISOString()),
    updatedAt: new Date().toISOString()
  };
}

function sanitizeInstaller(payload) {
  return {
    id: String(payload.id || makeId()),
    name: String(payload.name || "").trim(),
    company: String(payload.company || "").trim(),
    phone: String(payload.phone || "").trim(),
    email: String(payload.email || "").trim(),
    address: String(payload.address || "").trim(),
    notes: String(payload.notes || "").trim(),
    rating: Number.isFinite(Number(payload.rating)) ? Math.max(0, Math.min(5, Number(payload.rating))) : 0,
    regions: Array.isArray(payload.regions)
      ? payload.regions.map((region) => String(region || "").trim()).filter(Boolean)
      : [],
    tags: Array.isArray(payload.tags)
      ? payload.tags.map((tag) => String(tag || "").trim()).filter(Boolean)
      : [],
    createdAt: String(payload.createdAt || new Date().toISOString()),
    updatedAt: new Date().toISOString()
  };
}

function sanitizeRequest(payload) {
  return {
    ...sanitizeInstaller(payload),
    createdAt: String(payload.createdAt || new Date().toISOString()),
    updatedAt: new Date().toISOString()
  };
}

function isValidIsoDate(isoDate) {
  return /^\d{4}-\d{2}-\d{2}$/.test(isoDate) && Boolean(parseIsoDate(isoDate));
}

function broadcast(event, payload) {
  const body = `event: ${event}\ndata: ${JSON.stringify(payload)}\n\n`;
  for (const client of streamClients) {
    client.write(body);
  }
}

async function getBoardPayload() {
  const store = await readStore();
  return {
    jobs: store.jobs,
    holidays: store.holidays,
    board: buildBoardRows(store.jobs, store.holidays)
  };
}

function getCoreBridgeConfig() {
  return {
    baseUrl: String(process.env.COREBRIDGE_BASE_URL || DEFAULT_COREBRIDGE_BASE_URL).trim(),
    token: String(process.env.COREBRIDGE_TOKEN || process.env.COREBRIDGE_AUTH_TOKEN || "").trim(),
    subscriptionKey: String(
      process.env.COREBRIDGE_SUBSCRIPTION_KEY || process.env.COREBRIDGE_OCP_KEY || ""
    ).trim(),
    orderPath: String(process.env.COREBRIDGE_ORDER_PATH || process.env.COREBRIDGE_ENDPOINT_PATH || DEFAULT_COREBRIDGE_ORDER_PATH).trim(),
    apiVersion: String(process.env.COREBRIDGE_API_VERSION || "v3.0").trim()
  };
}

function buildCoreBridgeOrderUrl(config, params = {}) {
  const normalizedBase = config.baseUrl.endsWith("/") ? config.baseUrl : `${config.baseUrl}/`;
  const normalizedPath = config.orderPath.startsWith("/") ? config.orderPath.slice(1) : config.orderPath;
  const url = new URL(normalizedPath, normalizedBase);

  url.searchParams.set("apiversion", config.apiVersion);

  for (const [key, value] of Object.entries(params)) {
    if (value !== undefined && value !== null && value !== "") {
      url.searchParams.set(key, String(value));
    }
  }

  return url.toString();
}

function buildCoreBridgeOrderDetailUrl(config, orderId) {
  const normalizedBase = config.baseUrl.endsWith("/") ? config.baseUrl : `${config.baseUrl}/`;
  const normalizedPath = config.orderPath.startsWith("/") ? config.orderPath.slice(1) : config.orderPath;
  const detailPath = `${normalizedPath}/${orderId}`;
  const url = new URL(detailPath, normalizedBase);

  url.searchParams.set("apiversion", config.apiVersion);
  url.searchParams.set("companylevel", "full");
  url.searchParams.set("contactlevel", "full");
  url.searchParams.set("notelevel", "full");
  url.searchParams.set("destinationlevel", "full");
  url.searchParams.set("itemlevel", "full");

  return url.toString();
}

function flattenRecord(record, prefix = "", bucket = {}) {
  if (!record || typeof record !== "object") return bucket;

  for (const [key, value] of Object.entries(record)) {
    const nextPrefix = prefix ? `${prefix}.${key}` : key;
    const normalizedPrefix = nextPrefix.toLowerCase();

    if (Array.isArray(value)) {
      const joined = value
        .map((item) => (typeof item === "string" || typeof item === "number" ? String(item).trim() : ""))
        .filter(Boolean)
        .join(", ");
      if (joined) {
        bucket[normalizedPrefix] = joined;
        if (!(key.toLowerCase() in bucket)) {
          bucket[key.toLowerCase()] = joined;
        }
      }
      value.forEach((item, index) => {
        if (item && typeof item === "object") {
          flattenRecord(item, `${nextPrefix}.${index}`, bucket);
        }
      });
      continue;
    }

    if (value && typeof value === "object") {
      flattenRecord(value, nextPrefix, bucket);
      continue;
    }

    if (value !== undefined && value !== null) {
      const stringValue = String(value).trim();
      if (stringValue) {
        bucket[normalizedPrefix] = stringValue;
        if (!(key.toLowerCase() in bucket)) {
          bucket[key.toLowerCase()] = stringValue;
        }
      }
    }
  }

  return bucket;
}

function pickFirst(flatRecord, aliases) {
  for (const alias of aliases) {
    const value = flatRecord[alias.toLowerCase()];
    if (value) return value;
  }
  return "";
}

function looksLikePhone(value) {
  const normalized = String(value || "").trim();
  if (!normalized) return false;
  if (/^\d{4}-\d{2}-\d{2}t\d{2}:\d{2}:\d{2}/i.test(normalized)) return false;
  if (normalized.includes(":")) return false;
  if (/[A-Za-z]{3,}/.test(normalized)) return false;
  if (!/^[\d\s+().\-\/&]+$/.test(normalized)) return false;

  const digits = normalized.replace(/\D/g, "");
  return digits.length >= 7;
}

function pickFirstPhone(flatRecord, aliases) {
  for (const alias of aliases) {
    const value = flatRecord[alias.toLowerCase()];
    if (value && looksLikePhone(value)) return value;
  }
  return "";
}

function pickMatchingValue(flatRecord, matcher) {
  for (const [key, value] of Object.entries(flatRecord)) {
    if (matcher(key, value)) {
      return value;
    }
  }
  return "";
}

function isGenericCoreBridgeDescription(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (!normalized) return true;

  return [
    "courier",
    "delivery",
    "install",
    "installation",
    "survey",
    "labor",
    "labour",
    "work sheet labor",
    "worksheet labor",
    "work sheet labour",
    "worksheet labour"
  ].includes(normalized);
}

function pickBestCoreBridgeDescription(flatRecord) {
  const aliases = [
    "se_estimatedescription",
    "estimate.descriptiontext",
    "orderdescription",
    "order.description",
    "transactionheaderdata.description",
    "header.description",
    "estimatedescription",
    "estimate.description",
    "estimateheader.description",
    "estimateheader.displayname",
    "projectdescription",
    "projectname",
    "jobdescription",
    "title",
    "summary",
    "subject",
    "displayname",
    "headerdescription",
    "transactiondescription",
    "destinations.0.displayname",
    "simpledestinations.0.displayname",
    "items.0.displayname",
    "items.0.description",
    "items.0.invoicetext",
    "description",
    "items.0.name",
    "name"
  ];

  let fallback = "";
  for (const alias of aliases) {
    const value = flatRecord[alias.toLowerCase()];
    if (!value) continue;
    if (!isGenericCoreBridgeDescription(value)) return value;
    if (!fallback) fallback = value;
  }

  const scanned = pickMatchingValue(flatRecord, (key, value) => {
    const normalizedKey = String(key || "").toLowerCase();
    if (
      normalizedKey.includes("company") ||
      normalizedKey.includes("contact") ||
      normalizedKey.includes("locator") ||
      normalizedKey.includes("status") ||
      normalizedKey.includes("type")
    ) {
      return false;
    }

    if (
      normalizedKey.includes("description") ||
      normalizedKey.includes("displayname") ||
      normalizedKey.includes("title") ||
      normalizedKey.includes("summary") ||
      normalizedKey.includes("subject") ||
      normalizedKey.includes("project") ||
      normalizedKey.includes("name")
    ) {
      return !isGenericCoreBridgeDescription(value) && String(value).trim().length > 6;
    }

    return false;
  });

  return scanned || fallback;
}

function pickBestCoreBridgePhone(flatRecord) {
  const contactRoleLocatorPhone = pickMatchingValue(flatRecord, (key, value) => {
    const normalizedKey = String(key || "").toLowerCase();
    if (!normalizedKey.includes("contactroles.0") || !normalizedKey.endsWith(".locator")) return false;
    if (!looksLikePhone(value)) return false;

    const baseKey = normalizedKey.slice(0, -".locator".length);
    const subType =
      flatRecord[`${baseKey}.locatorsubtypenavigation.name`] ||
      flatRecord[`${baseKey}.locatortypenavigation.name`] ||
      flatRecord[`${baseKey}.locatorsubtype`] ||
      flatRecord[`${baseKey}.locatortype`] ||
      "";

    return /phone|mobile|tel|cell/i.test(String(subType));
  });

  if (contactRoleLocatorPhone) return contactRoleLocatorPhone;

  const directPhone = pickFirstPhone(flatRecord, [
    "phone",
    "telephone",
    "mobilenumber",
    "contactphone",
    "contactnumber",
    "company.phone",
    "company.telephone",
    "contactroles.0.phone",
    "contactroles.0.contactphone",
    "contactroles.0.ordercontactrolelocators.0.locator"
  ]);

  if (directPhone) return directPhone;

  return pickMatchingValue(flatRecord, (key, value) => {
    const normalizedKey = String(key || "").toLowerCase();
    if (!looksLikePhone(value)) return false;
    return (
      normalizedKey.includes("phone") ||
      normalizedKey.includes("telephone") ||
      normalizedKey.includes("mobile") ||
      normalizedKey.includes("locator")
    );
  });
}

function buildAddressFromAliases(flatRecord, aliasesByLine) {
  return aliasesByLine
    .map((aliases) => pickFirst(flatRecord, aliases))
    .filter(Boolean)
    .join(", ");
}

function pickPreferredCoreBridgeContactRole(record) {
  const roles = Array.isArray(record?.ContactRoles) ? record.ContactRoles : [];
  if (!roles.length) return null;

  return (
    roles.find((role) => String(role?.RoleType || "").toLowerCase() === "shipto") ||
    roles.find((role) => role?.DestinationID) ||
    roles[0]
  );
}

function pickRoleLocator(role, locatorType) {
  const locators = Array.isArray(role?.OrderContactRoleLocators) ? role.OrderContactRoleLocators : [];
  return locators.find((locator) => Number(locator?.LocatorType) === locatorType) || null;
}

function buildAddressFromRole(role) {
  const addressLocator = pickRoleLocator(role, 1);
  const metadata = addressLocator?.MetaData || addressLocator?.MetaDataObject || addressLocator?.metadata || {};

  const parts = [
    metadata.Street1 || metadata.street1 || "",
    metadata.Street2 || metadata.street2 || "",
    metadata.City || metadata.city || "",
    metadata.State || metadata.state || "",
    metadata.PostalCode || metadata.postalcode || metadata.Postcode || metadata.postcode || ""
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean);

  if (parts.length) return parts.join(", ");
  return String(addressLocator?.Locator || "").trim();
}

function buildPhoneFromRole(role) {
  const phoneLocator = pickRoleLocator(role, 2);
  const locatorValue = String(phoneLocator?.Locator || "").trim();
  return looksLikePhone(locatorValue) ? locatorValue : "";
}

function pickBestCoreBridgeAddress(flatRecord) {
  const destinationAddress = buildAddressFromAliases(flatRecord, [
    [
      "contactroles.0.ordercontactrolelocators.0.metadata.street1",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.street1"
    ],
    [
      "contactroles.0.ordercontactrolelocators.0.metadata.street2",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.street2"
    ],
    [
      "contactroles.0.ordercontactrolelocators.0.metadata.city",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.city"
    ],
    [
      "contactroles.0.ordercontactrolelocators.0.metadata.state",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.state"
    ],
    [
      "contactroles.0.ordercontactrolelocators.0.metadata.postalcode",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.postcode",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.postalcode",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.postcode"
    ],
    [
      "destinations.0.toaddress",
      "destinations.0.shiptoaddress",
      "destinations.0.address1",
      "destinations.0.toaddress1",
      "destinations.0.shiptoaddress1",
      "shiptoaddress1",
      "toaddress1"
    ],
    [
      "destinations.0.address2",
      "destinations.0.toaddress2",
      "destinations.0.shiptoaddress2",
      "shiptoaddress2",
      "toaddress2"
    ],
    [
      "destinations.0.address3",
      "destinations.0.toaddress3",
      "destinations.0.shiptoaddress3",
      "shiptoaddress3",
      "toaddress3"
    ],
    [
      "destinations.0.city",
      "destinations.0.tocity",
      "destinations.0.shiptocity",
      "shiptocity",
      "tocity"
    ],
    [
      "destinations.0.county",
      "destinations.0.state",
      "destinations.0.tocounty",
      "destinations.0.shiptocounty",
      "shiptocounty",
      "tocounty"
    ],
    [
      "destinations.0.postcode",
      "destinations.0.postalcode",
      "destinations.0.shiptopostcode",
      "destinations.0.topostcode",
      "shiptopostcode",
      "topostcode"
    ]
  ]);

  if (destinationAddress) return destinationAddress;

  const orderLevelAddress = buildAddressFromAliases(flatRecord, [
    ["shiptoaddress1", "toaddress1", "address1", "address.line1", "siteaddress1"],
    ["shiptoaddress2", "toaddress2", "address2", "address.line2", "siteaddress2"],
    ["shiptoaddress3", "toaddress3", "address3", "address.line3", "siteaddress3"],
    ["shiptocity", "tocity", "city", "address.city", "town", "sitecity"],
    ["shiptocounty", "tocounty", "county", "address.county", "state", "sitestate"],
    ["shiptopostcode", "topostcode", "postcode", "postalcode", "zip", "address.postcode", "sitepostcode"]
  ]);

  if (orderLevelAddress) return orderLevelAddress;

  return pickFirst(flatRecord, [
    "destinations.0.toaddress",
    "destinations.0.shiptoaddress",
    "shiptoaddress",
    "toaddress",
    "address",
    "siteaddress"
  ]);
}

function normalizeCoreBridgeOrder(record, index) {
  const flat = flattenRecord(record);
  const preferredRole = pickPreferredCoreBridgeContactRole(record);
  const directDescription = String(
    record?.SE_EstimateDescription ||
    record?.EstimateDescription ||
    record?.OrderDescription ||
    record?.Description ||
    ""
  ).trim();
  const preferredRoleAddress = buildAddressFromRole(preferredRole);
  const directRoleAddress = buildAddressFromAliases(flat, [
    [
      "contactroles.0.ordercontactrolelocators.0.metadata.street1",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.street1"
    ],
    [
      "contactroles.0.ordercontactrolelocators.0.metadata.street2",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.street2"
    ],
    [
      "contactroles.0.ordercontactrolelocators.0.metadata.city",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.city"
    ],
    [
      "contactroles.0.ordercontactrolelocators.0.metadata.state",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.state"
    ],
    [
      "contactroles.0.ordercontactrolelocators.0.metadata.postalcode",
      "contactroles.0.ordercontactrolelocators.0.metadata.postcode",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.postalcode",
      "ordercontactroles.0.ordercontactrolelocators.0.metadata.postcode"
    ]
  ]);
  const preferredRolePhone = buildPhoneFromRole(preferredRole);
  const directRolePhone = pickFirstPhone(flat, [
    "contactroles.0.ordercontactrolelocators.1.locator",
    "ordercontactroles.0.ordercontactrolelocators.1.locator",
    "contactroles.0.ordercontactrolelocators.0.metadata.businessphone",
    "ordercontactroles.0.ordercontactrolelocators.0.metadata.businessphone"
  ]);

  const normalized = {
    id: pickFirst(flat, ["id", "orderid", "jobid", "salesorderid"]) || `corebridge-${index}`,
    orderReference: pickFirst(flat, [
      "formattednumber",
      "ordernumber",
      "orderreference",
      "reference",
      "jobnumber",
      "documentnumber",
      "salesordernumber"
    ]),
    customerName: pickFirst(flat, [
      "company.displayname",
      "company.name",
      "company.longname",
      "company.legalname",
      "contactcompany",
      "customername",
      "customer",
      "companyname",
      "company",
      "accountname",
      "clientname",
      "name"
    ]),
    description: directDescription || pickBestCoreBridgeDescription(flat),
    contact: String(preferredRole?.ContactName || "").trim() || pickFirst(flat, [
      "contactroles.0.contactname",
      "ordercontactroles.0.contactname",
      "estimatecontactname",
      "estimatecontact",
      "ordercontactname",
      "ordercontact",
      "contact",
      "contactname",
      "primarycontact",
      "contactperson",
      "customercontact"
    ]),
    number: preferredRolePhone || directRolePhone,
    address: preferredRoleAddress || directRoleAddress || "",
    notes: pickFirst(flat, [
      "notes.0.note",
      "note",
      "notes",
      "internalnotes",
      "comments",
      "specialinstructions",
      "instructions"
    ]),
    status: pickFirst(flat, ["enumorderorderstatus.name", "status", "orderstatus", "jobstatus"]),
    raw: record
  };

  return normalized;
}

function buildCoreBridgeDebugFields(record) {
  const flat = flattenRecord(record);
  return Object.entries(flat)
    .sort(([leftKey], [rightKey]) => leftKey.localeCompare(rightKey))
    .map(([key, value]) => ({
      key,
      value
    }));
}

function looksLikeCoreBridgeOrderRecord(record, orderId) {
  if (!record || typeof record !== "object" || Array.isArray(record)) return false;
  const numericOrderId = Number(orderId);
  const recordId = Number(record.ID ?? record.id ?? record.OrderID ?? record.orderId);
  if (!Number.isFinite(recordId) || recordId !== numericOrderId) return false;

  return Boolean(
    record.FormattedNumber ||
    record.formattedNumber ||
    Array.isArray(record.ContactRoles) ||
    Array.isArray(record.contactRoles) ||
    record.CompanyID !== undefined ||
    record.companyId !== undefined ||
    record.TransactionType !== undefined ||
    record.transactionType !== undefined
  );
}

function extractCoreBridgeDetailRecord(payload, orderId) {
  if (looksLikeCoreBridgeOrderRecord(payload, orderId)) return payload;

  const directCandidates = [
    payload?.data,
    payload?.order,
    payload?.result,
    payload?.item,
    payload?.value
  ];

  for (const candidate of directCandidates) {
    if (looksLikeCoreBridgeOrderRecord(candidate, orderId)) return candidate;
  }

  const queue = [payload];
  const visited = new Set();

  while (queue.length) {
    const current = queue.shift();
    if (!current || typeof current !== "object") continue;
    if (visited.has(current)) continue;
    visited.add(current);

    if (looksLikeCoreBridgeOrderRecord(current, orderId)) {
      return current;
    }

    if (Array.isArray(current)) {
      for (const item of current) {
        if (item && typeof item === "object") {
          queue.push(item);
        }
      }
      continue;
    }

    for (const value of Object.values(current)) {
      if (value && typeof value === "object") {
        queue.push(value);
      }
    }
  }

  return payload;
}

async function fetchCoreBridgeOrderDetail(config, orderId, includeDebug = false) {
  const response = await fetch(buildCoreBridgeOrderDetailUrl(config, orderId), {
    headers: {
      Authorization: `Bearer ${config.token}`,
      "Ocp-Apim-Subscription-Key": config.subscriptionKey,
      Accept: "application/json"
    }
  });

  if (!response.ok) {
    throw new Error(`Detail lookup failed for ${orderId} (${response.status})`);
  }

  const contentType = String(response.headers.get("content-type") || "");
  const rawBody = await response.text();
  if (contentType.includes("text/html") || /^\s*</.test(rawBody)) {
    throw new Error(`Detail lookup returned HTML for ${orderId}`);
  }

  const body = JSON.parse(rawBody);
  const record = extractCoreBridgeDetailRecord(body, orderId);
  const normalized = normalizeCoreBridgeOrder(record, 0);
  normalized._detailFetched = true;
  normalized._detailOrderId = orderId;

  if (includeDebug) {
    normalized.debugFields = buildCoreBridgeDebugFields(record);
    normalized.debugRaw = JSON.stringify(record, null, 2);
  }

  return normalized;
}

function extractCoreBridgeRecords(payload) {
  if (Array.isArray(payload)) return payload;
  if (!payload || typeof payload !== "object") return [];

  const directCandidates = [
    payload.data,
    payload.items,
    payload.results,
    payload.records,
    payload.orders,
    payload.value,
    payload.rows
  ];

  for (const candidate of directCandidates) {
    if (Array.isArray(candidate)) return candidate;
  }

  if (payload.data && typeof payload.data === "object") {
    for (const nested of [payload.data.items, payload.data.results, payload.data.records, payload.data.orders, payload.data.rows]) {
      if (Array.isArray(nested)) return nested;
    }
  }

  return [];
}

function filterCoreBridgeOrders(orders, searchTerm = "") {
  const normalizedSearch = String(searchTerm || "").trim().toLowerCase();
  const filteredByStatus = orders.filter((order) => {
    const status = String(order.status || "").toLowerCase();
    return !status || !["closed", "cancelled", "canceled", "complete", "completed", "invoiced"].includes(status);
  });

  if (!normalizedSearch) {
    return filteredByStatus.slice(0, 100);
  }

  return filteredByStatus.filter((order) =>
    [
      order.orderReference,
      order.customerName,
      order.description,
      order.contact,
      order.number,
      order.address,
      order.notes,
      order.status
    ]
      .join(" ")
      .toLowerCase()
      .includes(normalizedSearch)
  );
}

async function fetchCoreBridgeOrders(searchTerm = "", includeDebug = false) {
  const config = getCoreBridgeConfig();
  if (!config.token || !config.subscriptionKey) {
    const error = new Error("CoreBridge is not configured yet.");
    error.statusCode = 503;
    throw error;
  }

  const attempts = [];
  const normalizedSearch = String(searchTerm || "").trim();
  const looksLikeFormattedNumber = /[a-z]{2,5}-?\d+/i.test(normalizedSearch);
  const requestPlans = [
    {
      label: "detailed",
      url: buildCoreBridgeOrderUrl(config, {
        take: 200,
        sortBy: "-modifiedDT",
        companylevel: "full",
        contactlevel: "full",
        notelevel: "full",
        destinationlevel: "full",
        itemlevel: "full",
        formattednumber: looksLikeFormattedNumber ? normalizedSearch : ""
      })
    },
    {
      label: "basic",
      url: buildCoreBridgeOrderUrl(config, {
        take: 200,
        sortBy: "-modifiedDT",
        formattednumber: looksLikeFormattedNumber ? normalizedSearch : ""
      })
    }
  ];

  for (const plan of requestPlans) {
    try {
      const response = await fetch(plan.url, {
        headers: {
          Authorization: `Bearer ${config.token}`,
          "Ocp-Apim-Subscription-Key": config.subscriptionKey,
          Accept: "application/json"
        }
      });

      if (!response.ok) {
        attempts.push(`${response.status} ${plan.url}`);
        continue;
      }

      const contentType = String(response.headers.get("content-type") || "");
      const rawBody = await response.text();
      if (contentType.includes("text/html") || /^\s*</.test(rawBody)) {
        attempts.push(`HTML ${plan.url}`);
        continue;
      }

      let body;
      try {
        body = contentType.includes("application/json") ? JSON.parse(rawBody) : JSON.parse(rawBody);
      } catch (error) {
        attempts.push(`NONJSON ${plan.url}`);
        continue;
      }

      const records = extractCoreBridgeRecords(body);
      let orders = filterCoreBridgeOrders(
        records.map((record, index) => {
          const normalized = normalizeCoreBridgeOrder(record, index);
          if (includeDebug) {
            normalized.debugFields = buildCoreBridgeDebugFields(record);
            normalized.debugRaw = JSON.stringify(record, null, 2);
          }
          return normalized;
        }),
        normalizedSearch
      ).filter((order) => order.orderReference || order.customerName);

      if (looksLikeFormattedNumber && orders.length) {
          orders = await Promise.all(
          orders.slice(0, 10).map(async (order) => {
            if (!order.id) return order;

            try {
              const detailedOrder = await fetchCoreBridgeOrderDetail(config, order.id, includeDebug);
              return {
                ...order,
                ...detailedOrder,
                _detailFetched: true,
                _detailOrderId: order.id
              };
            } catch (error) {
              attempts.push(`DETAIL ${order.id} ${error.message}`);
              return {
                ...order,
                _detailFetched: false,
                _detailError: error.message,
                _detailOrderId: order.id
              };
            }
          })
        );
      }

      if (orders.length || !normalizedSearch) {
        return {
          orders,
          sourceUrl: plan.url
        };
      }
    } catch (error) {
      attempts.push(`ERR ${plan.url} ${error.message}`);
    }
  }

  const error = new Error(`CoreBridge lookup failed. Tried: ${attempts.join(" | ") || "no endpoints"}`);
  error.statusCode = 502;
  throw error;
}

function createServer() {
  ensureStoreFile();
  ensureInstallersFile();
  ensureRequestsFile();
  ensureUsersFile();
  bootstrapPasswordsFromEnv()
    .then((result) => {
      if (result.updated) {
        console.log(`Bootstrapped passwords for ${result.updated} users.`);
      }
    })
    .catch((error) => {
      console.error("Could not bootstrap auth passwords.", error.message);
    });
  const app = express();

  app.use(cors());
  app.use(express.json({ limit: "1mb" }));
  app.use(express.static(PUBLIC_DIR));
  app.use((request, response, next) => {
    clearExpiredSessions();
    next();
  });

  app.get("/api/auth/users", async (request, response) => {
    const store = await readUsersStore();
    response.json(store.users.map(sanitizeUser));
  });

  app.get("/api/auth/me", (request, response) => {
    const session = getSessionFromRequest(request);
    if (!session) {
      response.status(401).json({ error: "Not signed in." });
      return;
    }

    response.json({ user: session.user });
  });

  app.post("/api/auth/login", async (request, response) => {
    const displayName = String(request.body?.displayName || "").trim();
    const password = String(request.body?.password || "");
    const store = await readUsersStore();
    const user = store.users.find(
      (entry) => String(entry.displayName || "").toLowerCase() === displayName.toLowerCase()
    );

    if (!user || !user.passwordHash) {
      response.status(401).json({ error: "That account is not ready yet." });
      return;
    }

    if (!verifyPassword(password, user)) {
      response.status(401).json({ error: "Incorrect password." });
      return;
    }

    const sessionUser = sanitizeUser(user);
    const { sessionId, expiresAt } = createSession(sessionUser);
    response.setHeader("Set-Cookie", serializeSessionCookie(sessionId, { expiresAt }));
    response.json({ user: sessionUser });
  });

  app.post("/api/auth/logout", (request, response) => {
    const session = getSessionFromRequest(request);
    if (session?.sessionId) {
      sessions.delete(session.sessionId);
    }
    response.setHeader("Set-Cookie", serializeSessionCookie("", { clear: true }));
    response.json({ ok: true });
  });

  app.use("/api", (request, response, next) => {
    if (request.path.startsWith("/auth/")) {
      next();
      return;
    }

    const session = getSessionFromRequest(request);
    if (!session) {
      response.status(401).json({ error: "Login required." });
      return;
    }

    request.user = session.user;
    next();
  });

  app.get("/healthz", (request, response) => {
    response.json({ ok: true });
  });

  app.get("/api/board", async (request, response) => {
    const payload = await getBoardPayload();
    response.json(payload.board);
  });

  app.get("/api/jobs", async (request, response) => {
    const store = await readStore();
    response.json(store.jobs);
  });

  app.get("/api/installers", async (request, response) => {
    if (!requireHost(request, response)) return;
    const installers = await readInstallersStore();
    console.log(`Serving ${installers.length} installers from ${getInstallersFile()}`);
    response.json(installers);
  });

  app.get("/api/installers/status", async (request, response) => {
    if (!requireHost(request, response)) return;
    const installers = await readInstallersStore();
    const requests = await readRequestsStore();
    response.json({
      installers: installers.length,
      requests: requests.length
    });
  });

  app.get("/api/installers/debug", async (request, response) => {
    if (!requireHost(request, response)) return;

    const candidates = [
      LEGACY_INSTALLERS_FILE,
      process.env.DATA_FILE ? path.join(path.dirname(process.env.DATA_FILE), "installers.json") : "",
      DEFAULT_INSTALLERS_FILE
    ].filter(Boolean);

    const uniqueCandidates = [...new Set(candidates)];
    const inspections = await Promise.all(uniqueCandidates.map((file) => inspectJsonArrayFile(file)));
    const currentFile = getInstallersFile();
    const installers = await readInstallersStore();

    response.json({
      currentFile,
      currentCount: installers.length,
      sample: installers.slice(0, 3),
      candidates: inspections
    });
  });

  app.get("/api/requests", async (request, response) => {
    if (!requireHost(request, response)) return;
    response.json(await readRequestsStore());
  });

app.get("/api/corebridge/orders", async (request, response) => {
  if (!requireHost(request, response)) return;
  try {
    const searchTerm = String(request.query.q || "").trim();
    const includeDebug = String(request.query.debug || "").trim() === "1";
      const payload = await fetchCoreBridgeOrders(searchTerm, includeDebug);
      response.json(payload);
    } catch (error) {
      console.error("CoreBridge lookup failed.", error.message);
      response.status(error.statusCode || 500).json({
        error:
          error.statusCode === 503
            ? "CoreBridge is not configured yet."
            : "Could not reach CoreBridge order lookup yet.",
        detail: error.message
      });
    }
  });

  app.get("/api/corebridge/orders/:id", async (request, response) => {
    if (!requireHost(request, response)) return;
    try {
      const orderId = String(request.params.id || "").trim();
      const includeDebug = String(request.query.debug || "").trim() === "1";
      if (!orderId) {
        response.status(400).json({ error: "An order id is required." });
        return;
      }

      const config = getCoreBridgeConfig();
      if (!config.token || !config.subscriptionKey) {
        response.status(503).json({ error: "CoreBridge is not configured yet." });
        return;
      }

      const order = await fetchCoreBridgeOrderDetail(config, orderId, includeDebug);
      response.json(order);
    } catch (error) {
      console.error("CoreBridge detail lookup failed.", error.message);
      response.status(500).json({
        error: "Could not load the CoreBridge order detail.",
        detail: error.message
      });
    }
  });

  app.post("/api/jobs", async (request, response) => {
    if (!requireHost(request, response)) return;
    const nextJob = sanitizeJob(request.body || {});
    if (!nextJob.customerName || !isValidIsoDate(nextJob.date)) {
      response.status(400).json({ error: "A valid date and customer name are required." });
      return;
    }

    const store = await readStore();
    const index = store.jobs.findIndex((job) => job.id === nextJob.id);
    if (index >= 0) {
      nextJob.createdAt = store.jobs[index].createdAt || nextJob.createdAt;
      store.jobs[index] = nextJob;
    } else {
      store.jobs.unshift(nextJob);
    }

    const savedStore = await writeStore(store);
    const payload = {
      jobs: savedStore.jobs,
      holidays: savedStore.holidays,
      board: buildBoardRows(savedStore.jobs, savedStore.holidays)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.delete("/api/jobs/:id", async (request, response) => {
    if (!requireHost(request, response)) return;
    const store = await readStore();
    store.jobs = store.jobs.filter((job) => job.id !== request.params.id);
    const savedStore = await writeStore(store);
    const payload = {
      jobs: savedStore.jobs,
      holidays: savedStore.holidays,
      board: buildBoardRows(savedStore.jobs, savedStore.holidays)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.post("/api/installers", async (request, response) => {
    if (!requireHost(request, response)) return;
    const nextInstaller = sanitizeInstaller(request.body || {});

    if (!nextInstaller.name) {
      response.status(400).json({ error: "Installer name is required." });
      return;
    }

    const installers = await readInstallersStore();
    const index = installers.findIndex((installer) => installer.id === nextInstaller.id);

    if (index >= 0) {
      nextInstaller.createdAt = installers[index].createdAt || nextInstaller.createdAt;
      installers[index] = nextInstaller;
    } else {
      installers.unshift(nextInstaller);
    }

    response.json(await writeInstallersStore(installers));
  });

  app.post("/api/requests", async (request, response) => {
    if (!requireHost(request, response)) return;
    const nextRequest = sanitizeRequest(request.body || {});

    if (!nextRequest.name && !nextRequest.company) {
      response.status(400).json({ error: "A company or contact name is required." });
      return;
    }

    const requests = await readRequestsStore();
    const index = requests.findIndex((requestItem) => requestItem.id === nextRequest.id);

    if (index >= 0) {
      nextRequest.createdAt = requests[index].createdAt || nextRequest.createdAt;
      requests[index] = nextRequest;
    } else {
      requests.unshift(nextRequest);
    }

    response.json(await writeRequestsStore(requests));
  });

  app.delete("/api/installers/:id", async (request, response) => {
    if (!requireHost(request, response)) return;
    const installers = await readInstallersStore();
    const filtered = installers.filter((installer) => installer.id !== request.params.id);
    response.json(await writeInstallersStore(filtered));
  });

  app.delete("/api/requests/:id", async (request, response) => {
    if (!requireHost(request, response)) return;
    const requests = await readRequestsStore();
    const filtered = requests.filter((requestItem) => requestItem.id !== request.params.id);
    response.json(await writeRequestsStore(filtered));
  });

  app.get("/api/holidays", async (request, response) => {
    const store = await readStore();
    response.json(store.holidays);
  });

  app.post("/api/holidays", async (request, response) => {
    if (!requireHost(request, response)) return;
    const nextHoliday = sanitizeStaffHoliday(request.body || {});
    if (!nextHoliday.person || !isValidIsoDate(nextHoliday.date)) {
      response.status(400).json({ error: "A valid date and person are required." });
      return;
    }

    const store = await readStore();
    const index = store.holidays.findIndex((item) => item.id === nextHoliday.id);
    if (index >= 0) {
      nextHoliday.createdAt = store.holidays[index].createdAt || nextHoliday.createdAt;
      store.holidays[index] = nextHoliday;
    } else {
      store.holidays.unshift(nextHoliday);
    }

    const savedStore = await writeStore(store);
    const payload = {
      jobs: savedStore.jobs,
      holidays: savedStore.holidays,
      board: buildBoardRows(savedStore.jobs, savedStore.holidays)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.delete("/api/holidays/:id", async (request, response) => {
    if (!requireHost(request, response)) return;
    const store = await readStore();
    store.holidays = store.holidays.filter((item) => item.id !== request.params.id);
    const savedStore = await writeStore(store);
    const payload = {
      jobs: savedStore.jobs,
      holidays: savedStore.holidays,
      board: buildBoardRows(savedStore.jobs, savedStore.holidays)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.get("/api/events", (request, response) => {
    response.writeHead(200, {
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-cache, no-transform",
      Connection: "keep-alive"
    });

    response.write(`event: ready\ndata: ${JSON.stringify({ ok: true })}\n\n`);
    streamClients.add(response);

    const keepAlive = setInterval(() => {
      response.write(`event: ping\ndata: ${Date.now()}\n\n`);
    }, 15000);

    request.on("close", () => {
      clearInterval(keepAlive);
      streamClients.delete(response);
    });
  });

  if (fs.existsSync(DIST_DIR)) {
    app.use(express.static(DIST_DIR));
    app.get("*", (request, response) => {
      response.sendFile(path.join(DIST_DIR, "index.html"));
    });
  } else {
    app.get("*", (request, response) => {
      const requestHost = String(request.headers.host || `localhost:${PORT}`).split(":")[0];
      response.redirect(`http://${requestHost}:${DEV_FRONTEND_PORT}${request.originalUrl}`);
    });
  }

  return app;
}

async function startServer() {
  const app = createServer();
  return new Promise((resolve) => {
    const server = app.listen(PORT, HOST, () => {
      console.log(`Installation board server running on http://${HOST}:${PORT}`);
      console.log(`Installation board data file: ${getDataFile()}`);
      resolve(server);
    });
  });
}

if (require.main === module) {
  startServer().catch((error) => {
    console.error(error);
    process.exit(1);
  });
}

module.exports = {
  buildBoardRows,
  createServer,
  getUkBankHolidays,
  normalizeCoreBridgeOrder,
  startServer
};
