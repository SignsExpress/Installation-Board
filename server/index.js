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
  updateUserPermissions,
  verifyPassword
} = require("./auth-store");

const PORT = Number(process.env.PORT || 3030);
const HOST = process.env.HOST || "0.0.0.0";
const DEV_FRONTEND_PORT = Number(process.env.DEV_FRONTEND_PORT || 5173);
const DIST_DIR = path.join(__dirname, "..", "dist");
const DEFAULT_DATA_FILE = path.join(__dirname, "..", "data", "jobs.json");
const DEFAULT_INSTALLERS_FILE = path.join(__dirname, "..", "data", "installers-live.json");
const DEFAULT_REQUESTS_FILE = path.join(__dirname, "..", "data", "requests.json");
const DEFAULT_HOLIDAY_SEED_FILE = path.join(__dirname, "..", "data", "holiday-seed.json");
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
const HOLIDAY_STAFF = [
  { code: "MR", name: "Matt Rutlidge", person: "Matt R", birthDate: "1989-05-04" },
  { code: "DD", name: "Dawn Dewhurst", person: "Dawn D", birthDate: "1971-10-09" },
  { code: "TVB", name: "Tom Van-Boyd", person: "Tom V-B", birthDate: "1993-07-27" },
  { code: "AH", name: "Amber Hardman", person: "Amber H", birthDate: "2002-08-08" },
  { code: "ED", name: "Eddy D'Antonio", person: "Eddy D'A", birthDate: "1997-02-06" },
  { code: "PM", name: "Paul Morris", person: "Paul M", birthDate: "1983-03-11" },
  { code: "KW", name: "Kyle Wright", person: "Kyle W", birthDate: "2004-12-12" },
  { code: "MC", name: "Matt Carroll", person: "Matt C", birthDate: "1992-11-22" },
  { code: "KC", name: "Keilan Curtis", person: "Keilan C", birthDate: "1998-10-24" },
  { code: "TS", name: "Tamas", person: "Tamas" }
];

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

async function getFreshSessionUser(session) {
  if (!session?.user?.id) return session?.user || null;
  const store = await readUsersStore();
  const storedUser = store.users.find((user) => String(user.id || "") === String(session.user.id || ""));
  return storedUser ? sanitizeUser(storedUser) : session.user;
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

function getUserPermission(user, key, fallback = "none") {
  const value = String(user?.permissions?.[key] || "").trim().toLowerCase();
  return ["admin", "user", "none"].includes(value) ? value : fallback;
}

function canAccessBoard(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "board", user?.role === "host" ? "admin" : "user") !== "none";
}

function canEditBoard(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "board", user?.role === "host" ? "admin" : "user") === "admin";
}

function canAccessInstaller(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "installer", user?.role === "host" ? "admin" : "none") !== "none";
}

function canEditInstaller(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "installer", user?.role === "host" ? "admin" : "none") === "admin";
}

function canAccessHolidays(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "holidays", user?.role === "host" ? "admin" : "user") !== "none";
}

function canEditHolidays(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "holidays", user?.role === "host" ? "admin" : "user") === "admin";
}

function canManagePermissions(user) {
  return String(user?.displayName || "").trim().toLowerCase() === "matt rutlidge";
}

function getHolidayStaffCode(displayName) {
  const normalized = String(displayName || "").trim().toLowerCase();
  return HOLIDAY_STAFF.find((entry) => entry.name.toLowerCase() === normalized)?.code || "";
}

function getHolidayStaffName(codeOrName) {
  const normalized = String(codeOrName || "").trim().toLowerCase();
  const match =
    HOLIDAY_STAFF.find((entry) => entry.code.toLowerCase() === normalized) ||
    HOLIDAY_STAFF.find((entry) => entry.name.toLowerCase() === normalized);
  return match?.name || "";
}

function getHolidayStaffEntry(codeOrName) {
  const normalized = String(codeOrName || "").trim().toLowerCase();
  return (
    HOLIDAY_STAFF.find((entry) => entry.code.toLowerCase() === normalized) ||
    HOLIDAY_STAFF.find((entry) => entry.person.toLowerCase() === normalized) ||
    HOLIDAY_STAFF.find((entry) => entry.name.toLowerCase() === normalized) ||
    null
  );
}

function getHolidayStaffPerson(displayName) {
  return HOLIDAY_STAFF.find((entry) => entry.name.toLowerCase() === String(displayName || "").trim().toLowerCase())?.person || "";
}

function requireBoardAccess(request, response) {
  if (canAccessBoard(request.user)) return true;
  response.status(403).json({ error: "Board access required." });
  return false;
}

function requireBoardAdmin(request, response) {
  if (canEditBoard(request.user)) return true;
  response.status(403).json({ error: "Board admin access required." });
  return false;
}

function requireInstallerAccess(request, response) {
  if (canAccessInstaller(request.user)) return true;
  response.status(403).json({ error: "Subcontractor directory access required." });
  return false;
}

function requireInstallerAdmin(request, response) {
  if (canEditInstaller(request.user)) return true;
  response.status(403).json({ error: "Subcontractor directory admin access required." });
  return false;
}

function requirePermissionsManager(request, response) {
  if (canManagePermissions(request.user)) return true;
  response.status(403).json({ error: "Permissions manager access required." });
  return false;
}

function requireHolidayAccess(request, response) {
  if (canAccessHolidays(request.user)) return true;
  response.status(403).json({ error: "Holiday access required." });
  return false;
}

function requireHolidayAdmin(request, response) {
  if (canEditHolidays(request.user)) return true;
  response.status(403).json({ error: "Holiday admin access required." });
  return false;
}

function ensureStoreFile() {
  const file = getDataFile();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  if (!fs.existsSync(file)) {
    fs.writeFileSync(file, `${JSON.stringify({ jobs: [], holidays: [], holidayRequests: [], holidayAllowances: [] }, null, 2)}\n`, "utf8");
  }
}

function readHolidaySeed() {
  if (!fs.existsSync(DEFAULT_HOLIDAY_SEED_FILE)) {
    return { holidays: [], holidayAllowances: [] };
  }

  try {
    const raw = fs.readFileSync(DEFAULT_HOLIDAY_SEED_FILE, "utf8");
    const parsed = JSON.parse(raw);
    return {
      holidays: Array.isArray(parsed.holidays) ? parsed.holidays : [],
      holidayAllowances: Array.isArray(parsed.holidayAllowances) ? parsed.holidayAllowances : []
    };
  } catch (error) {
    console.error("Invalid holiday seed JSON, ignoring seed.", error);
    return { holidays: [], holidayAllowances: [] };
  }
}

function mergeHolidaySeed(store) {
  const seed = readHolidaySeed();
  const nextStore = {
    jobs: Array.isArray(store.jobs) ? store.jobs : [],
    holidays: Array.isArray(store.holidays) ? [...store.holidays] : [],
    holidayRequests: Array.isArray(store.holidayRequests) ? [...store.holidayRequests] : [],
    holidayAllowances: Array.isArray(store.holidayAllowances) ? [...store.holidayAllowances] : []
  };

  seed.holidays.forEach((holiday) => {
    const normalized = sanitizeStaffHoliday(holiday);
    const exists = nextStore.holidays.some(
      (entry) =>
        String(entry.date || "") === normalized.date &&
        String(entry.person || "").trim().toLowerCase() === String(normalized.person || "").trim().toLowerCase() &&
        String(entry.duration || "Full Day").trim().toLowerCase() === String(normalized.duration || "Full Day").trim().toLowerCase()
    );
    if (!exists) {
      nextStore.holidays.push(normalized);
    }
  });

  seed.holidayAllowances.forEach((allowance) => {
    const normalized = sanitizeHolidayAllowance(allowance);
    const exists = nextStore.holidayAllowances.some(
      (entry) =>
        Number(entry.yearStart || 0) === Number(normalized.yearStart || 0) &&
        String(entry.person || "").trim().toLowerCase() === String(normalized.person || "").trim().toLowerCase()
    );
    if (!exists) {
      nextStore.holidayAllowances.push(normalized);
    }
  });

  return nextStore;
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
      return mergeHolidaySeed({ jobs: parsed, holidays: [], holidayRequests: [], holidayAllowances: [] });
    }
    return mergeHolidaySeed({
      jobs: Array.isArray(parsed.jobs) ? parsed.jobs : [],
      holidays: Array.isArray(parsed.holidays) ? parsed.holidays : [],
      holidayRequests: Array.isArray(parsed.holidayRequests) ? parsed.holidayRequests : [],
      holidayAllowances: Array.isArray(parsed.holidayAllowances) ? parsed.holidayAllowances : []
    });
  } catch (error) {
    console.error("Invalid board store JSON, returning empty store.", error);
    return mergeHolidaySeed({ jobs: [], holidays: [], holidayRequests: [], holidayAllowances: [] });
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
    }),
    holidayRequests: [...(store.holidayRequests || [])].sort((left, right) => {
      if (left.startDate !== right.startDate) return String(left.startDate || "").localeCompare(String(right.startDate || ""));
      return String(left.person || "").localeCompare(String(right.person || ""));
    }),
    holidayAllowances: [...(store.holidayAllowances || [])].sort((left, right) => {
      if (left.yearStart !== right.yearStart) return Number(left.yearStart || 0) - Number(right.yearStart || 0);
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
  return {
    start: addDays(today, -7),
    end: addDays(today, 21)
  };
}

function getStartOfMonth(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1));
}

function getEndOfMonth(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0));
}

function addMonths(date, months) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + months, 1));
}

function toMonthId(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
}

function parseMonthId(value) {
  const match = String(value || "").trim().match(/^(\d{4})-(\d{2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  if (!year || month < 1 || month > 12) return null;
  return new Date(Date.UTC(year, month - 1, 1));
}

function buildMonthNavigation(today = getTodayInLondon()) {
  const currentMonth = getStartOfMonth(today);
  const previousMonths = [];
  const futureMonths = [];

  for (let offset = 1; offset <= 6; offset += 1) {
    const previousDate = addMonths(currentMonth, -offset);
    previousMonths.push({
      id: toMonthId(previousDate),
      label: monthFormatter.format(previousDate)
    });

    const futureDate = addMonths(currentMonth, offset);
    futureMonths.push({
      id: toMonthId(futureDate),
      label: monthFormatter.format(futureDate)
    });
  }

  return { previousMonths, futureMonths };
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

function buildBoardRows(jobs, staffHolidays, options = {}) {
  const today = options.today || getTodayInLondon();
  const mode = options.mode === "month" ? "month" : "rolling";
  const start = options.start || (mode === "month" ? getStartOfMonth(today) : getRollingWindow(today).start);
  const end = options.end || (mode === "month" ? getEndOfMonth(today) : getRollingWindow(today).end);
  const weekdayDates = getWeekdaysInRange(start, end);
  const holidayMap = getHolidayMap(start, end);
  const todayIso = toIsoDate(today);
  const displayStaffHolidays = getDisplayStaffHolidays(
    staffHolidays || [],
    toIsoDate(start),
    toIsoDate(end),
    Array.isArray(options.birthdayEntries) ? options.birthdayEntries : []
  );

  const jobsByDate = jobs.reduce((map, job) => {
    if (!isValidIsoDate(job.date)) {
      return map;
    }
    const existing = map.get(job.date) || [];
    existing.push(job);
    map.set(job.date, existing);
    return map;
  }, new Map());

  const unscheduled = jobs
    .filter((job) => !isValidIsoDate(job.date))
    .sort((left, right) => String(left.customerName || "").localeCompare(String(right.customerName || "")));

  const staffHolidaysByDate = displayStaffHolidays.reduce((map, entry) => {
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
  const weekMap = new Map();
  rows.forEach((row) => {
    const rowDate = parseIsoDate(row.isoDate);
    const weekStart = getStartOfWeek(rowDate);
    const weekId = toIsoDate(weekStart);
    if (!weekMap.has(weekId)) {
      weekMap.set(weekId, {
        id: weekId,
        label: formatWeekLabel(weekStart),
        rows: []
      });
      weeks.push(weekMap.get(weekId));
    }
    weekMap.get(weekId).rows.push(row);
  });

  return {
    generatedAt: new Date().toISOString(),
    mode,
    today: todayIso,
    start: toIsoDate(start),
    end: toIsoDate(end),
    rangeLabel:
      mode === "month"
        ? monthFormatter.format(start)
        : `${start.getUTCDate()} ${monthFormatter.format(start).split(" ")[0]} to ${end.getUTCDate()} ${monthFormatter.format(end).split(" ")[0]}`,
    weeks,
    unscheduled
  };
}

function buildBirthdayEntriesForRange(store, startIso, endIso) {
  const startYear = getHolidayYearStartForIsoDate(startIso) || getCurrentHolidayYearStart();
  const endYear = getHolidayYearStartForIsoDate(endIso) || startYear;
  const entries = [];

  for (let yearStart = startYear; yearStart <= endYear; yearStart += 1) {
    const allowanceRows = buildBaseHolidayAllowanceRows(store, yearStart);
    entries.push(...buildBirthdayHolidayEntries(allowanceRows, yearStart));
  }

  return entries;
}

function buildBoardRowsFromStore(store, options = {}) {
  const today = options.today || getTodayInLondon();
  const mode = options.mode === "month" ? "month" : options.mode === "range" ? "range" : "rolling";
  const rangeStart = options.start || (mode === "month" ? getStartOfMonth(today) : getRollingWindow(today).start);
  const rangeEnd = options.end || (mode === "month" ? getEndOfMonth(today) : getRollingWindow(today).end);
  const birthdayEntries = buildBirthdayEntriesForRange(store, toIsoDate(rangeStart), toIsoDate(rangeEnd));
  return buildBoardRows(store.jobs, store.holidays, { ...options, today, mode, start: rangeStart, end: rangeEnd, birthdayEntries });
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
    isPlaceholder:
      payload.isPlaceholder === true ||
      String(payload.isPlaceholder || "").trim().toLowerCase() === "true" ||
      String(payload.isPlaceholder || "").trim() === "1",
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
    type: String(payload.type || "holiday").trim().toLowerCase() || "holiday",
    createdAt: String(payload.createdAt || new Date().toISOString()),
    updatedAt: new Date().toISOString()
  };
}

function sanitizeHolidayRequest(payload) {
  const normalizedStatus = String(payload.status || "pending").trim().toLowerCase();
  const status = ["pending", "approved", "rejected"].includes(normalizedStatus) ? normalizedStatus : "pending";
  const normalizedDuration = String(payload.duration || "Full Day").trim();
  return {
    id: String(payload.id || makeId()),
    person: String(payload.person || "").trim(),
    requestedByUserId: String(payload.requestedByUserId || "").trim(),
    requestedByName: String(payload.requestedByName || "").trim(),
    startDate: String(payload.startDate || "").trim(),
    endDate: String(payload.endDate || payload.startDate || "").trim(),
    duration: normalizedDuration || "Full Day",
    notes: String(payload.notes || "").trim(),
    status,
    reviewedAt: String(payload.reviewedAt || "").trim(),
    reviewedBy: String(payload.reviewedBy || "").trim(),
    createdAt: String(payload.createdAt || new Date().toISOString()),
    updatedAt: new Date().toISOString()
  };
}

function sanitizeHolidayAllowance(payload) {
  function toNumber(value) {
    const numeric = Number(value);
    return Number.isFinite(numeric) ? numeric : 0;
  }

  return {
    id: String(payload.id || makeId()),
    yearStart: Number(payload.yearStart || getCurrentHolidayYearStart()),
    person: String(payload.person || "").trim(),
    workDaysPerWeek: toNumber(payload.workDaysPerWeek),
    standardEntitlement: toNumber(payload.standardEntitlement),
    extraServiceDays: toNumber(payload.extraServiceDays),
    birthDate: String(payload.birthDate || "").trim(),
    christmasDays: toNumber(payload.christmasDays),
    bankHolidayDays: toNumber(payload.bankHolidayDays),
    unpaidDaysBooked: toNumber(payload.unpaidDaysBooked),
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

function enumerateIsoDates(startDate, endDate) {
  const start = parseIsoDate(startDate);
  const end = parseIsoDate(endDate);
  if (!start || !end || start > end) return [];

  const dates = [];
  let cursor = start;
  while (cursor <= end) {
    dates.push(toIsoDate(cursor));
    cursor = addDays(cursor, 1);
  }
  return dates;
}

function isWeekdayIsoDate(isoDate) {
  const parsed = parseIsoDate(isoDate);
  if (!parsed) return false;
  const day = parsed.getUTCDay();
  return day >= 1 && day <= 5;
}

function isBankHolidayIsoDate(isoDate) {
  const parsed = parseIsoDate(isoDate);
  if (!parsed) return false;
  return getUkBankHolidays(parsed.getUTCFullYear()).some((holiday) => holiday.date === isoDate);
}

function findNearestWorkingIsoDate(isoDate, preferredDirection = 0) {
  const base = parseIsoDate(isoDate);
  if (!base) return "";

  const directionalOffsets =
    preferredDirection < 0
      ? [-1, 1]
      : preferredDirection > 0
        ? [1, -1]
        : [-1, 1];

  for (let distance = 1; distance <= 7; distance += 1) {
    for (const direction of directionalOffsets) {
      const candidate = addDays(base, distance * direction);
      const candidateIso = toIsoDate(candidate);
      if (isWeekdayIsoDate(candidateIso) && !isBankHolidayIsoDate(candidateIso)) {
        return candidateIso;
      }
    }
  }

  return "";
}

function getObservedBirthdayIsoDate(staffEntry, yearStart = getCurrentHolidayYearStart()) {
  const birthDate = String(staffEntry?.birthDate || "").trim();
  if (!birthDate || !/^\d{4}-\d{2}-\d{2}$/.test(birthDate)) return "";
  const [, monthString, dayString] = birthDate.split("-");
  const month = Number(monthString);
  const day = Number(dayString);
  if (!month || !day) return "";

  const occurrenceYear = month >= 2 ? yearStart : yearStart + 1;
  const birthday = new Date(Date.UTC(occurrenceYear, month - 1, day));
  const birthdayIso = toIsoDate(birthday);
  const weekday = birthday.getUTCDay();

  if (weekday === 6) {
    return findNearestWorkingIsoDate(birthdayIso, -1) || birthdayIso;
  }

  if (weekday === 0) {
    return findNearestWorkingIsoDate(birthdayIso, 1) || birthdayIso;
  }

  if (isBankHolidayIsoDate(birthdayIso)) {
    return findNearestWorkingIsoDate(birthdayIso, 0) || birthdayIso;
  }

  return birthdayIso;
}

function buildBirthdayHolidayEntries(allowanceRows = [], yearStart = getCurrentHolidayYearStart()) {
  return allowanceRows.map((allowanceEntry) => {
    const observedDate = getObservedBirthdayIsoDate({ birthDate: allowanceEntry.birthDate }, yearStart);
    if (!observedDate) return null;
    return sanitizeStaffHoliday({
      id: `birthday-${yearStart}-${String(allowanceEntry.code || allowanceEntry.person || "").toLowerCase().replace(/[^a-z0-9]+/g, "-")}`,
      date: observedDate,
      person: allowanceEntry.person,
      duration: "Full Day",
      type: "birthday"
    });
  }).filter(Boolean);
}

function getHolidayType(entry) {
  return String(entry?.type || "holiday").trim().toLowerCase();
}

function buildBaseHolidayAllowanceRows(store, yearStart = getCurrentHolidayYearStart()) {
  return HOLIDAY_STAFF.map((staffEntry) => {
    const existing = (store.holidayAllowances || []).find(
      (entry) =>
        Number(entry.yearStart || 0) === yearStart &&
        String(entry.person || "").trim().toLowerCase() === staffEntry.person.toLowerCase()
    );

    const normalized = sanitizeHolidayAllowance({
      yearStart,
      person: staffEntry.person,
      birthDate: staffEntry.birthDate || "",
      ...(existing || {})
    });

    return {
      ...normalized,
      code: staffEntry.code,
      fullName: staffEntry.name
    };
  });
}

function getDisplayStaffHolidays(staffHolidays, startIso, endIso, birthdayEntries = []) {
  const startYear = getHolidayYearStartForIsoDate(startIso) || getCurrentHolidayYearStart();
  const endYear = getHolidayYearStartForIsoDate(endIso) || startYear;
  const scopedBirthdayEntries = birthdayEntries.filter((entry) => {
    const yearStart = getHolidayYearStartForIsoDate(entry.date);
    return yearStart >= startYear && yearStart <= endYear;
  });

  const birthdayMap = new Map(
    scopedBirthdayEntries.map((entry) => [
      `${String(entry.person || "").trim().toLowerCase()}::${entry.date}`,
      entry
    ])
  );

  const normalized = (staffHolidays || [])
    .map((entry) => sanitizeStaffHoliday(entry))
    .filter((entry) => entry.date >= startIso && entry.date <= endIso)
    .map((entry) => {
      const key = `${String(entry.person || "").trim().toLowerCase()}::${entry.date}`;
      if (birthdayMap.has(key)) {
        return {
          ...entry,
          id: birthdayMap.get(key).id,
          type: "birthday"
        };
      }
      return entry;
    });

  const existingKeys = new Set(
    normalized.map((entry) => `${String(entry.person || "").trim().toLowerCase()}::${entry.date}`)
  );

  scopedBirthdayEntries.forEach((entry) => {
    const key = `${String(entry.person || "").trim().toLowerCase()}::${entry.date}`;
    if (!existingKeys.has(key) && entry.date >= startIso && entry.date <= endIso) {
      normalized.push(entry);
    }
  });

  return normalized.sort((left, right) => {
    if (left.date !== right.date) return String(left.date || "").localeCompare(String(right.date || ""));
    return String(left.person || "").localeCompare(String(right.person || ""));
  });
}

function getCurrentHolidayYearStart(today = getTodayInLondon()) {
  const year = today.getUTCFullYear();
  return today.getUTCMonth() >= 1 ? year : year - 1;
}

function getHolidayYearBounds(yearStart = getCurrentHolidayYearStart()) {
  return {
    yearStart,
    start: new Date(Date.UTC(yearStart, 1, 1)),
    end: new Date(Date.UTC(yearStart + 1, 0, 31))
  };
}

function getHolidayYearLabel(yearStart = getCurrentHolidayYearStart()) {
  const endYear = String(yearStart + 1).slice(-2);
  return `${yearStart}-${endYear}`;
}

function getHolidayYearStartForIsoDate(isoDate) {
  const parsed = parseIsoDate(isoDate);
  if (!parsed) return null;
  return parsed.getUTCMonth() >= 1 ? parsed.getUTCFullYear() : parsed.getUTCFullYear() - 1;
}

function getHolidayYearOptions(currentYearStart = getCurrentHolidayYearStart(), yearsAhead = 3) {
  return Array.from({ length: yearsAhead + 1 }, (_, index) => {
    const yearStart = currentYearStart + index;
    return {
      yearStart,
      label: getHolidayYearLabel(yearStart)
    };
  });
}

function calculateHolidayDays(startDate, endDate, duration = "Full Day") {
  const weekdays = enumerateIsoDates(startDate, endDate).filter(isWeekdayIsoDate);
  const factor = duration === "Morning" || duration === "Afternoon" ? 0.5 : 1;
  return weekdays.length * factor;
}

function buildHolidayAllowanceSummaries(store, yearStart = getCurrentHolidayYearStart()) {
  const bounds = getHolidayYearBounds(yearStart);
  const startIso = toIsoDate(bounds.start);
  const endIso = toIsoDate(bounds.end);
  const allowanceRows = buildBaseHolidayAllowanceRows(store, yearStart);
  const birthdayEntries = buildBirthdayHolidayEntries(allowanceRows, yearStart);
  const displayHolidays = getDisplayStaffHolidays(store.holidays || [], startIso, endIso, birthdayEntries);
  const approvedCounts = new Map();

  displayHolidays.forEach((holiday) => {
    if (!holiday?.date || holiday.date < startIso || holiday.date > endIso) return;
    if (getHolidayType(holiday) === "birthday") return;
    const person = String(holiday.person || "").trim();
    const current = approvedCounts.get(person) || 0;
    approvedCounts.set(
      person,
      current + (holiday.duration === "Morning" || holiday.duration === "Afternoon" ? 0.5 : 1)
    );
  });

  return allowanceRows.map((normalized) => {
    const prorataAllowance =
      normalized.standardEntitlement +
      normalized.extraServiceDays;
    const bookedDays = approvedCounts.get(normalized.person) || 0;
    const daysLeft =
      prorataAllowance -
      normalized.christmasDays -
      normalized.bankHolidayDays -
      bookedDays;

    return {
      ...normalized,
      observedBirthdayDate: getObservedBirthdayIsoDate({ birthDate: normalized.birthDate }, yearStart),
      prorataAllowance,
      bookedDays,
      daysLeft
    };
  });
}

function broadcast(event, payload) {
  const body = `event: ${event}\ndata: ${JSON.stringify(payload)}\n\n`;
  for (const client of streamClients) {
    client.write(body);
  }
}

async function getBoardPayload(options = {}) {
  const store = await readStore();
  const today = getTodayInLondon();
  const navigation = buildMonthNavigation(today);
  let boardOptions = { today, mode: "rolling" };

  if (options.start && options.end) {
    boardOptions = {
      today,
      mode: "range",
      start: options.start,
      end: options.end
    };
  } else if (options.mode === "month" && options.monthId) {
    const monthDate = parseMonthId(options.monthId);
    if (monthDate) {
      boardOptions = {
        today,
        mode: "month",
        start: getStartOfMonth(monthDate),
        end: getEndOfMonth(monthDate)
      };
    }
  }

  return {
    jobs: store.jobs,
    holidays: store.holidays,
    board: {
      ...buildBoardRowsFromStore(store, boardOptions),
      ...navigation,
      selectedMonth: options.mode === "month" ? options.monthId || "" : ""
    }
  };
}

async function getHolidayPayload(forUser, yearStart = getCurrentHolidayYearStart()) {
  const store = await readStore();
  const currentPerson = getHolidayStaffPerson(forUser?.displayName);
  const bounds = getHolidayYearBounds(yearStart);
  const startIso = toIsoDate(bounds.start);
  const endIso = toIsoDate(bounds.end);
  const visibleRequests = canEditHolidays(forUser)
    ? (store.holidayRequests || [])
    : (store.holidayRequests || []).filter(
        (request) =>
          String(request.requestedByUserId || "") === String(forUser?.id || "") ||
          String(request.person || "").trim().toLowerCase() === String(currentPerson || "").toLowerCase()
      );

  const yearRequests = visibleRequests.filter((request) => {
    const requestStart = String(request.startDate || "");
    const requestEnd = String(request.endDate || request.startDate || "");
    return requestStart <= endIso && requestEnd >= startIso && String(request.status || "pending").trim().toLowerCase() === "pending";
  });

  const yearHolidays = (store.holidays || []).filter((holiday) => {
    const holidayDate = String(holiday.date || "");
    return holidayDate >= startIso && holidayDate <= endIso;
  });
  const birthdayEntries = buildBirthdayHolidayEntries(buildBaseHolidayAllowanceRows(store, yearStart), yearStart);
  const displayYearHolidays = getDisplayStaffHolidays(store.holidays || [], startIso, endIso, birthdayEntries);

  const allowanceRows = buildHolidayAllowanceSummaries(store, yearStart).filter((entry) =>
    canEditHolidays(forUser)
      ? true
      : String(entry.person || "").trim().toLowerCase() === String(currentPerson || "").toLowerCase()
  );

  return {
    holidays: displayYearHolidays,
    holidayRequests: yearRequests,
    holidayStaff: HOLIDAY_STAFF,
    holidayAllowances: allowanceRows,
    holidayYearStart: yearStart,
    currentHolidayYearStart: getCurrentHolidayYearStart(),
    holidayYearLabel: getHolidayYearLabel(yearStart),
    holidayYearOptions: getHolidayYearOptions(getCurrentHolidayYearStart())
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

function buildCoreBridgeOrderDestinationsUrl(config, orderId) {
  const normalizedBase = config.baseUrl.endsWith("/") ? config.baseUrl : `${config.baseUrl}/`;
  const normalizedPath = config.orderPath.startsWith("/") ? config.orderPath.slice(1) : config.orderPath;
  const detailPath = `${normalizedPath}/${orderId}/orderdestination`;
  const url = new URL(detailPath, normalizedBase);

  url.searchParams.set("apiversion", config.apiVersion);
  url.searchParams.set("contactlevel", "full");

  return url.toString();
}

function buildCoreBridgeDestinationDetailUrl(config, destinationId) {
  const normalizedBase = config.baseUrl.endsWith("/") ? config.baseUrl : `${config.baseUrl}/`;
  const normalizedPath = config.orderPath.startsWith("/") ? config.orderPath.slice(1) : config.orderPath;
  const destinationPath = normalizedPath.replace(/order\/?$/i, `destination/${destinationId}`);
  const url = new URL(destinationPath, normalizedBase);

  url.searchParams.set("apiversion", config.apiVersion);
  url.searchParams.set("contactlevel", "full");

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

function getCoreBridgeRoles(record) {
  if (Array.isArray(record?.ContactRoles)) return record.ContactRoles;
  if (Array.isArray(record?.OrderContactRoles)) return record.OrderContactRoles;
  return [];
}

function listFlatContactRoleIndexes(flatRecord) {
  const indexes = new Set();
  for (const key of Object.keys(flatRecord || {})) {
    const match = key.match(/^(?:contactroles|ordercontactroles)\.(\d+)\./);
    if (match) {
      indexes.add(Number(match[1]));
    }
  }
  return Array.from(indexes).sort((left, right) => left - right);
}

function listFlatRoleLocatorIndexes(flatRecord, rolePrefix, roleIndex) {
  const indexes = new Set();
  const pattern = new RegExp(`^${rolePrefix}\\.${roleIndex}\\.ordercontactrolelocators\\.(\\d+)\\.`);

  for (const key of Object.keys(flatRecord || {})) {
    const match = key.match(pattern);
    if (match) {
      indexes.add(Number(match[1]));
    }
  }

  return Array.from(indexes).sort((left, right) => left - right);
}

function buildFlatContactRoleAddress(flatRecord, rolePrefix, roleIndex) {
  const locatorIndexes = listFlatRoleLocatorIndexes(flatRecord, rolePrefix, roleIndex);

  for (const locatorIndex of locatorIndexes) {
    const locatorType = String(
      flatRecord[`${rolePrefix}.${roleIndex}.ordercontactrolelocators.${locatorIndex}.locatortype`] || ""
    ).trim();
    if (locatorType !== "1") continue;

    const address = buildAddressFromAliases(flatRecord, [
      [`${rolePrefix}.${roleIndex}.ordercontactrolelocators.${locatorIndex}.metadata.street1`],
      [`${rolePrefix}.${roleIndex}.ordercontactrolelocators.${locatorIndex}.metadata.street2`],
      [`${rolePrefix}.${roleIndex}.ordercontactrolelocators.${locatorIndex}.metadata.city`],
      [`${rolePrefix}.${roleIndex}.ordercontactrolelocators.${locatorIndex}.metadata.state`],
      [
        `${rolePrefix}.${roleIndex}.ordercontactrolelocators.${locatorIndex}.metadata.postalcode`,
        `${rolePrefix}.${roleIndex}.ordercontactrolelocators.${locatorIndex}.metadata.postcode`
      ]
    ]);

    if (address) return address;
  }

  return buildAddressFromAliases(flatRecord, [
    [`${rolePrefix}.${roleIndex}.ordercontactrolelocators.0.metadata.street1`],
    [`${rolePrefix}.${roleIndex}.ordercontactrolelocators.0.metadata.street2`],
    [`${rolePrefix}.${roleIndex}.ordercontactrolelocators.0.metadata.city`],
    [`${rolePrefix}.${roleIndex}.ordercontactrolelocators.0.metadata.state`],
    [
      `${rolePrefix}.${roleIndex}.ordercontactrolelocators.0.metadata.postalcode`,
      `${rolePrefix}.${roleIndex}.ordercontactrolelocators.0.metadata.postcode`
    ]
  ]);
}

function pickDestinationAddressFromFlat(flatRecord) {
  const indexes = listFlatContactRoleIndexes(flatRecord);
  const rolePrefixes = ["contactroles", "ordercontactroles"];

  for (const rolePrefix of rolePrefixes) {
    for (const index of indexes) {
      const roleType = String(flatRecord[`${rolePrefix}.${index}.roletype`] || "").toLowerCase();
      const destinationId = String(flatRecord[`${rolePrefix}.${index}.destinationid`] || "").trim();
      const address = buildFlatContactRoleAddress(flatRecord, rolePrefix, index);
      if (roleType === "shipto" && destinationId && address) {
        return address;
      }
    }
  }

  for (const rolePrefix of rolePrefixes) {
    for (const index of indexes) {
      const destinationId = String(flatRecord[`${rolePrefix}.${index}.destinationid`] || "").trim();
      const address = buildFlatContactRoleAddress(flatRecord, rolePrefix, index);
      if (destinationId && address) {
        return address;
      }
    }
  }

  return "";
}

function pickPreferredCoreBridgeContactRole(record) {
  const roles = getCoreBridgeRoles(record);
  if (!roles.length) return null;

  return (
    roles.find((role) => String(role?.RoleType || "").toLowerCase() === "shipto") ||
    roles.find((role) => role?.DestinationID) ||
    roles[0]
  );
}

function pickDestinationCoreBridgeRole(record) {
  const roles = getCoreBridgeRoles(record);
  if (!roles.length) return null;

  const hasAddressMetadata = (role) =>
    Array.isArray(role?.OrderContactRoleLocators) &&
    role.OrderContactRoleLocators.some((locator) => {
      if (Number(locator?.LocatorType) !== 1) return false;
      const metadata = locator?.MetaData || locator?.MetaDataObject || locator?.metadata || {};
      return Boolean(
        String(
          metadata?.Street1 ||
            metadata?.street1 ||
            metadata?.Street2 ||
            metadata?.street2 ||
            metadata?.City ||
            metadata?.city ||
            metadata?.State ||
            metadata?.state ||
            metadata?.PostalCode ||
            metadata?.postalcode ||
            metadata?.Postcode ||
            metadata?.postcode ||
            ""
        ).trim()
      );
    });

  return (
    roles.find(
      (role) =>
        role?.DestinationID &&
        String(role?.RoleType || "").toLowerCase() === "shipto" &&
        hasAddressMetadata(role)
    ) ||
    roles.find((role) => role?.DestinationID && hasAddressMetadata(role)) ||
    null
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
  if (looksLikePhone(locatorValue)) return locatorValue;

  const locators = Array.isArray(role?.OrderContactRoleLocators) ? role.OrderContactRoleLocators : [];
  for (const locator of locators) {
    const metadata = locator?.MetaData || locator?.MetaDataObject || locator?.metadata || {};
    const metadataPhone = String(
      metadata?.BusinessPhone ||
        metadata?.businessPhone ||
        metadata?.businessphone ||
        metadata?.Phone ||
        metadata?.phone ||
        metadata?.Telephone ||
        metadata?.telephone ||
        metadata?.Mobile ||
        metadata?.mobile ||
        ""
    ).trim();
    if (looksLikePhone(metadataPhone)) return metadataPhone;
  }

  return "";
}

function buildDestinationAddressFromRole(role) {
  const locators = Array.isArray(role?.OrderContactRoleLocators) ? role.OrderContactRoleLocators : [];
  const addressLocator = locators.find((locator) => Number(locator?.LocatorType) === 1) || null;
  if (!addressLocator) return "";

  const metadata = addressLocator?.MetaData || addressLocator?.MetaDataObject || addressLocator?.metadata || {};
  const parts = [
    metadata?.Street1 || metadata?.street1 || "",
    metadata?.Street2 || metadata?.street2 || "",
    metadata?.City || metadata?.city || "",
    metadata?.State || metadata?.state || "",
    metadata?.PostalCode || metadata?.postalcode || metadata?.Postcode || metadata?.postcode || ""
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean);

  return parts.join(", ");
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
  const destinationRole = pickDestinationCoreBridgeRole(record);
  const directDescription = String(
    record?.SE_EstimateDescription ||
    record?.EstimateDescription ||
    record?.OrderDescription ||
    record?.Description ||
    ""
  ).trim();
  const orderDestinationAddress = getOrderDestinationAddressFromOrderRecord(record);
  const orderDestinationPhone = getOrderDestinationPhoneFromOrderRecord(record);
  const destinationRoleAddress =
    buildAddressFromRole(destinationRole) ||
    pickDestinationAddressFromFlat(flat);
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
    number: orderDestinationPhone || preferredRolePhone || directRolePhone,
    address: orderDestinationAddress || destinationRoleAddress,
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

function getCoreBridgeDestinationAddressFromRecord(record) {
  if (!record || typeof record !== "object") return "";

  const roles = getCoreBridgeRoles(record);
  const directRoleAddress =
    roles.find((role) => String(role?.RoleType || "").toLowerCase() === "shipto" && buildAddressFromRole(role)) ||
    roles.find((role) => buildAddressFromRole(role)) ||
    null;
  if (directRoleAddress) {
    return buildAddressFromRole(directRoleAddress);
  }

  const flat = flattenRecord(record);
  const objectRoleAddress = buildAddressFromRole(pickDestinationCoreBridgeRole(record));
  if (objectRoleAddress) return objectRoleAddress;
  const flatRoleAddress = pickDestinationAddressFromFlat(flat);
  if (flatRoleAddress) return flatRoleAddress;

  return buildAddressFromAliases(flat, [
    ["toaddress", "shiptoaddress", "address1"],
    ["toaddress2", "shiptoaddress2", "address2"],
    ["tocity", "shiptocity", "city"],
    ["tocounty", "shiptocounty", "state", "county"],
    ["topostcode", "shiptopostcode", "postcode", "postalcode"]
  ]);
}

function pickBestCoreBridgeDestinationRecord(records) {
  if (!Array.isArray(records) || !records.length) return null;

  return (
    records.find((record) => record?.IsDefault) ||
    records.find((record) => {
      const roles = getCoreBridgeRoles(record);
      return roles.some((role) => String(role?.RoleType || "").toLowerCase() === "shipto");
    }) ||
    records[0]
  );
}

function pickBestOrderDestinationRecord(record) {
  const destinations = Array.isArray(record?.Destinations) ? record.Destinations : [];
  if (!destinations.length) return null;

  return destinations.find((destination) => destination?.IsDefault) || destinations[0];
}

function getOrderDestinationAddressFromOrderRecord(record) {
  const destinationRecord = pickBestOrderDestinationRecord(record);
  if (!destinationRecord) return "";
  return getCoreBridgeDestinationAddressFromRecord(destinationRecord);
}

function getCoreBridgeDestinationPhoneFromRecord(record) {
  if (!record || typeof record !== "object") return "";

  const roles = getCoreBridgeRoles(record);
  const matchingRole =
    roles.find((role) => String(role?.RoleType || "").toLowerCase() === "shipto" && buildPhoneFromRole(role)) ||
    roles.find((role) => buildPhoneFromRole(role)) ||
    null;

  if (matchingRole) {
    return buildPhoneFromRole(matchingRole);
  }

  const flat = flattenRecord(record);
  return pickFirstPhone(flat, [
    "contactroles.0.ordercontactrolelocators.1.locator",
    "ordercontactroles.0.ordercontactrolelocators.1.locator",
    "contactroles.0.ordercontactrolelocators.0.metadata.businessphone",
    "ordercontactroles.0.ordercontactrolelocators.0.metadata.businessphone",
    "businessphone",
    "phone",
    "telephone",
    "mobile"
  ]);
}

function getOrderDestinationPhoneFromOrderRecord(record) {
  const destinationRecord = pickBestOrderDestinationRecord(record);
  if (!destinationRecord) return "";
  return getCoreBridgeDestinationPhoneFromRecord(destinationRecord);
}

function pickDestinationLookupId(record) {
  if (!record || typeof record !== "object") return "";

  return String(
    record?.DestinationID ??
      record?.destinationId ??
      record?.ID ??
      record?.id ??
      record?.OrderDestinationID ??
      record?.orderDestinationId ??
      ""
  ).trim();
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

async function fetchCoreBridgeOrderDestinationAddress(config, orderId) {
  const response = await fetch(buildCoreBridgeOrderDestinationsUrl(config, orderId), {
    headers: {
      Authorization: `Bearer ${config.token}`,
      "Ocp-Apim-Subscription-Key": config.subscriptionKey,
      Accept: "application/json"
    }
  });

  if (!response.ok) {
    throw new Error(`Destination lookup failed for ${orderId} (${response.status})`);
  }

  const contentType = String(response.headers.get("content-type") || "");
  const rawBody = await response.text();
  if (contentType.includes("text/html") || /^\s*</.test(rawBody)) {
    throw new Error(`Destination lookup returned HTML for ${orderId}`);
  }

  const body = JSON.parse(rawBody);
  const records = extractCoreBridgeDestinationRecords(body);
  const destinationRecord = pickBestCoreBridgeDestinationRecord(records);
  const inlineAddress = getCoreBridgeDestinationAddressFromRecord(destinationRecord);

  const destinationLookupId = pickDestinationLookupId(destinationRecord);
  if (!destinationLookupId) return inlineAddress;

  const detailResponse = await fetch(buildCoreBridgeDestinationDetailUrl(config, destinationLookupId), {
    headers: {
      Authorization: `Bearer ${config.token}`,
      "Ocp-Apim-Subscription-Key": config.subscriptionKey,
      Accept: "application/json"
    }
  });

  if (!detailResponse.ok) {
    throw new Error(`Destination detail lookup failed for ${destinationLookupId} (${detailResponse.status})`);
  }

  const detailContentType = String(detailResponse.headers.get("content-type") || "");
  const detailRawBody = await detailResponse.text();
  if (detailContentType.includes("text/html") || /^\s*</.test(detailRawBody)) {
    throw new Error(`Destination detail lookup returned HTML for ${destinationLookupId}`);
  }

  const detailBody = JSON.parse(detailRawBody);
  const detailRecords = extractCoreBridgeDestinationRecords(detailBody);
  const detailRecord = detailRecords[0] || detailBody;
  return getCoreBridgeDestinationAddressFromRecord(detailRecord) || inlineAddress;
}

async function fetchCoreBridgeOrderDestinationDebug(config, orderId) {
  const destinationUrl = buildCoreBridgeOrderDestinationsUrl(config, orderId);
  const result = {
    orderId: String(orderId),
    destinationUrl,
    destinationRecords: [],
    chosenDestinationRecord: null,
    inlineAddress: "",
    destinationLookupId: "",
    destinationDetailUrl: "",
    destinationDetailRecord: null,
    destinationDetailAddress: "",
    errors: []
  };

  try {
    const response = await fetch(destinationUrl, {
      headers: {
        Authorization: `Bearer ${config.token}`,
        "Ocp-Apim-Subscription-Key": config.subscriptionKey,
        Accept: "application/json"
      }
    });

    result.destinationStatus = response.status;
    const contentType = String(response.headers.get("content-type") || "");
    const rawBody = await response.text();
    result.destinationBodyPreview = rawBody.slice(0, 4000);

    if (!response.ok) {
      result.errors.push(`Destination lookup failed (${response.status})`);
      return result;
    }

    if (contentType.includes("text/html") || /^\s*</.test(rawBody)) {
      result.errors.push("Destination lookup returned HTML");
      return result;
    }

    const body = JSON.parse(rawBody);
    const records = extractCoreBridgeDestinationRecords(body);
    result.destinationRecords = records.map((record) => ({
      id: pickDestinationLookupId(record),
      destinationId: String(record?.DestinationID ?? record?.destinationId ?? ""),
      destinationNumber: String(record?.DestinationNumber ?? record?.destinationNumber ?? ""),
      isDefault: Boolean(record?.IsDefault),
      inlineAddress: getCoreBridgeDestinationAddressFromRecord(record)
    }));

    const chosenRecord = pickBestCoreBridgeDestinationRecord(records);
    result.chosenDestinationRecord = chosenRecord
      ? {
          id: pickDestinationLookupId(chosenRecord),
          destinationId: String(chosenRecord?.DestinationID ?? chosenRecord?.destinationId ?? ""),
          destinationNumber: String(chosenRecord?.DestinationNumber ?? chosenRecord?.destinationNumber ?? ""),
          isDefault: Boolean(chosenRecord?.IsDefault)
        }
      : null;

    result.inlineAddress = getCoreBridgeDestinationAddressFromRecord(chosenRecord);
    const destinationLookupId = pickDestinationLookupId(chosenRecord);
    result.destinationLookupId = destinationLookupId;

    if (!destinationLookupId) {
      result.errors.push("No destination lookup id found");
      return result;
    }

    const destinationDetailUrl = buildCoreBridgeDestinationDetailUrl(config, destinationLookupId);
    result.destinationDetailUrl = destinationDetailUrl;
    const detailResponse = await fetch(destinationDetailUrl, {
      headers: {
        Authorization: `Bearer ${config.token}`,
        "Ocp-Apim-Subscription-Key": config.subscriptionKey,
        Accept: "application/json"
      }
    });

    result.destinationDetailStatus = detailResponse.status;
    const detailContentType = String(detailResponse.headers.get("content-type") || "");
    const detailRawBody = await detailResponse.text();
    result.destinationDetailBodyPreview = detailRawBody.slice(0, 4000);

    if (!detailResponse.ok) {
      result.errors.push(`Destination detail lookup failed (${detailResponse.status})`);
      return result;
    }

    if (detailContentType.includes("text/html") || /^\s*</.test(detailRawBody)) {
      result.errors.push("Destination detail lookup returned HTML");
      return result;
    }

    const detailBody = JSON.parse(detailRawBody);
    const detailRecords = extractCoreBridgeDestinationRecords(detailBody);
    const detailRecord = detailRecords[0] || detailBody;
    result.destinationDetailRecord = {
      id: pickDestinationLookupId(detailRecord),
      destinationId: String(detailRecord?.DestinationID ?? detailRecord?.destinationId ?? ""),
      destinationNumber: String(detailRecord?.DestinationNumber ?? detailRecord?.destinationNumber ?? "")
    };
    result.destinationDetailAddress = getCoreBridgeDestinationAddressFromRecord(detailRecord);
    return result;
  } catch (error) {
    result.errors.push(error.message || "Unknown destination debug error");
    return result;
  }
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

function looksLikeCoreBridgeDestinationRecord(record) {
  if (!record || typeof record !== "object" || Array.isArray(record)) return false;

  return Boolean(
    record.DestinationID !== undefined ||
    record.destinationId !== undefined ||
    record.DestinationNumber ||
    record.destinationNumber ||
    Array.isArray(record.ContactRoles) ||
    Array.isArray(record.OrderContactRoles) ||
    Array.isArray(record.OrderDestinationItems)
  );
}

function extractCoreBridgeDestinationRecords(payload) {
  if (Array.isArray(payload)) {
    return payload.filter((item) => looksLikeCoreBridgeDestinationRecord(item));
  }

  if (!payload || typeof payload !== "object") return [];

  if (looksLikeCoreBridgeDestinationRecord(payload)) {
    return [payload];
  }

  const directCandidates = [
    payload.data,
    payload.destination,
    payload.destinations,
    payload.result,
    payload.item,
    payload.value,
    payload.rows
  ];

  for (const candidate of directCandidates) {
    if (Array.isArray(candidate)) {
      const matches = candidate.filter((item) => looksLikeCoreBridgeDestinationRecord(item));
      if (matches.length) return matches;
    } else if (looksLikeCoreBridgeDestinationRecord(candidate)) {
      return [candidate];
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
              try {
                const destinationAddress = await fetchCoreBridgeOrderDestinationAddress(config, order.id);
                if (destinationAddress) {
                  detailedOrder.address = destinationAddress;
                }
              } catch (destinationError) {
                attempts.push(`DESTINATION ${order.id} ${destinationError.message}`);
              }
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
    const session = getSessionFromRequest(request);
    if (!session) {
      response.status(401).json({ error: "Login required." });
      return;
    }

    const freshSessionUser = await getFreshSessionUser(session);

    const store = await readUsersStore();
    const users = canManagePermissions(freshSessionUser)
      ? store.users
      : store.users.filter((user) => String(user.id || "") === String(freshSessionUser.id || ""));

    response.json(users.map((user) => ({
      ...sanitizeUser(user),
      canManagePermissions: canManagePermissions(sanitizeUser(user))
    })));
  });

  app.get("/api/auth/me", (request, response) => {
    const session = getSessionFromRequest(request);
    if (!session) {
      response.status(401).json({ error: "Not signed in." });
      return;
    }

    getFreshSessionUser(session)
      .then((freshUser) => {
        sessions.set(session.sessionId, {
          user: freshUser,
          expiresAt: session.expiresAt
        });
        response.json({
          user: {
            ...freshUser,
            canManagePermissions: canManagePermissions(freshUser)
          }
        });
      })
      .catch((error) => {
        response.status(500).json({ error: error.message || "Could not load current user." });
      });
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
    if (!canAccessBoard(sessionUser) && !canAccessInstaller(sessionUser) && !canAccessHolidays(sessionUser)) {
      response.status(403).json({ error: "That account does not have access." });
      return;
    }
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

  app.patch("/api/auth/users/:id/permissions", async (request, response) => {
    const session = getSessionFromRequest(request);
    if (!session) {
      response.status(401).json({ error: "Login required." });
      return;
    }

    request.user = session.user;
    if (!requirePermissionsManager(request, response)) return;

    try {
      const updatedUser = await updateUserPermissions(request.params.id, {
        board: request.body?.board,
        installer: request.body?.installer,
        holidays: request.body?.holidays
      });
      response.json({ user: updatedUser });
    } catch (error) {
      response.status(400).json({ error: error.message || "Could not update permissions." });
    }
  });

  app.use("/api", async (request, response, next) => {
    if (request.path.startsWith("/auth/")) {
      next();
      return;
    }

    const session = getSessionFromRequest(request);
    if (!session) {
      response.status(401).json({ error: "Login required." });
      return;
    }

    const freshUser = await getFreshSessionUser(session);
    sessions.set(session.sessionId, {
      user: freshUser,
      expiresAt: session.expiresAt
    });
    request.user = freshUser;
    next();
  });

  app.get("/healthz", (request, response) => {
    response.json({ ok: true });
  });

  app.get("/api/board", async (request, response) => {
    if (!requireBoardAccess(request, response)) return;
    const mode = String(request.query.mode || "").trim().toLowerCase();
    const monthId = String(request.query.month || "").trim();
    const start = parseIsoDate(String(request.query.start || "").trim());
    const end = parseIsoDate(String(request.query.end || "").trim());
    const payload = await getBoardPayload({
      start,
      end,
      mode: mode === "month" ? "month" : "rolling",
      monthId
    });
    response.json(payload.board);
  });

  app.get("/api/jobs", async (request, response) => {
    if (!requireBoardAccess(request, response)) return;
    const store = await readStore();
    response.json(store.jobs);
  });

  app.get("/api/installers", async (request, response) => {
    if (!requireInstallerAccess(request, response)) return;
    const installers = await readInstallersStore();
    console.log(`Serving ${installers.length} installers from ${getInstallersFile()}`);
    response.json(installers);
  });

  app.get("/api/installers/status", async (request, response) => {
    if (!requireInstallerAccess(request, response)) return;
    const installers = await readInstallersStore();
    const requests = await readRequestsStore();
    response.json({
      installers: installers.length,
      requests: requests.length
    });
  });

  app.get("/api/installers/debug", async (request, response) => {
    if (!requireInstallerAccess(request, response)) return;

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
    if (!requireInstallerAdmin(request, response)) return;
    response.json(await readRequestsStore());
  });

app.get("/api/corebridge/orders", async (request, response) => {
  if (!requireBoardAdmin(request, response)) return;
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
    if (!requireBoardAdmin(request, response)) return;
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
      try {
        const destinationAddress = await fetchCoreBridgeOrderDestinationAddress(config, orderId);
        if (destinationAddress) {
          order.address = destinationAddress;
        }
      } catch (destinationError) {
        if (includeDebug) {
          order._destinationError = destinationError.message;
        }
      }
      response.json(order);
    } catch (error) {
      console.error("CoreBridge detail lookup failed.", error.message);
      response.status(500).json({
        error: "Could not load the CoreBridge order detail.",
        detail: error.message
      });
    }
  });

  app.get("/api/corebridge/orders/:id/destination-debug", async (request, response) => {
    if (!requireBoardAdmin(request, response)) return;
    try {
      const orderId = String(request.params.id || "").trim();
      if (!orderId) {
        response.status(400).json({ error: "An order id is required." });
        return;
      }

      const config = getCoreBridgeConfig();
      if (!config.token || !config.subscriptionKey) {
        response.status(503).json({ error: "CoreBridge is not configured yet." });
        return;
      }

      const payload = await fetchCoreBridgeOrderDestinationDebug(config, orderId);
      response.json(payload);
    } catch (error) {
      console.error("CoreBridge destination debug failed.", error.message);
      response.status(500).json({
        error: "Could not load the CoreBridge destination debug.",
        detail: error.message
      });
    }
  });

  app.post("/api/jobs", async (request, response) => {
    if (!requireBoardAdmin(request, response)) return;
    const nextJob = sanitizeJob(request.body || {});
    if (!nextJob.customerName) {
      response.status(400).json({ error: "Customer name is required." });
      return;
    }
    if (nextJob.date && !isValidIsoDate(nextJob.date)) {
      response.status(400).json({ error: "If a date is supplied, it must be valid." });
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
      board: buildBoardRowsFromStore(savedStore)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.delete("/api/jobs/:id", async (request, response) => {
    if (!requireBoardAdmin(request, response)) return;
    const store = await readStore();
    store.jobs = store.jobs.filter((job) => job.id !== request.params.id);
    const savedStore = await writeStore(store);
    const payload = {
      jobs: savedStore.jobs,
      holidays: savedStore.holidays,
      board: buildBoardRowsFromStore(savedStore)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.post("/api/installers", async (request, response) => {
    if (!requireInstallerAdmin(request, response)) return;
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
    if (!requireInstallerAdmin(request, response)) return;
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
    if (!requireInstallerAdmin(request, response)) return;
    const installers = await readInstallersStore();
    const filtered = installers.filter((installer) => installer.id !== request.params.id);
    response.json(await writeInstallersStore(filtered));
  });

  app.delete("/api/requests/:id", async (request, response) => {
    if (!requireInstallerAdmin(request, response)) return;
    const requests = await readRequestsStore();
    const filtered = requests.filter((requestItem) => requestItem.id !== request.params.id);
    response.json(await writeRequestsStore(filtered));
  });

  app.get("/api/holidays", async (request, response) => {
    if (!requireHolidayAccess(request, response)) return;
    const store = await readStore();
    response.json(store.holidays);
  });

  app.get("/api/holiday-requests", async (request, response) => {
    if (!requireHolidayAccess(request, response)) return;
    const yearStart = Number(request.query.yearStart || getCurrentHolidayYearStart());
    const payload = await getHolidayPayload(request.user, yearStart);
    response.json(payload);
  });

  app.post("/api/holiday-requests", async (request, response) => {
    if (!requireHolidayAccess(request, response)) return;
    const personLabel = getHolidayStaffPerson(request.user?.displayName);
    if (!personLabel && !canEditHolidays(request.user)) {
      response.status(400).json({ error: "This account is not linked to a holiday calendar code." });
      return;
    }

    const nextRequest = sanitizeHolidayRequest({
      ...request.body,
      person: canEditHolidays(request.user) ? request.body?.person || personLabel : personLabel,
      requestedByUserId: request.user?.id || "",
      requestedByName: request.user?.displayName || "",
      status: "pending"
    });

    if (!getHolidayStaffEntry(nextRequest.person)) {
      response.status(400).json({ error: "A valid employee is required." });
      return;
    }

    if (!isValidIsoDate(nextRequest.startDate) || !isValidIsoDate(nextRequest.endDate)) {
      response.status(400).json({ error: "Valid start and end dates are required." });
      return;
    }

    if (parseIsoDate(nextRequest.startDate) > parseIsoDate(nextRequest.endDate)) {
      response.status(400).json({ error: "End date must be on or after the start date." });
      return;
    }

    const requestYearStart = getHolidayYearStartForIsoDate(nextRequest.startDate);
    const requestEndYearStart = getHolidayYearStartForIsoDate(nextRequest.endDate);
    if (!requestYearStart || requestYearStart !== requestEndYearStart) {
      response.status(400).json({ error: "Holiday requests must fit within a single holiday year." });
      return;
    }

    const store = await readStore();
    if (!canEditHolidays(request.user)) {
      const allowanceSummary = buildHolidayAllowanceSummaries(store, requestYearStart).find(
        (entry) => String(entry.person || "").trim().toLowerCase() === String(nextRequest.person || "").trim().toLowerCase()
      );
      const requestedDays = calculateHolidayDays(nextRequest.startDate, nextRequest.endDate, nextRequest.duration);
      if (allowanceSummary && requestedDays > allowanceSummary.daysLeft) {
        response.status(400).json({
          error: `Only ${allowanceSummary.daysLeft} holiday days remain in ${getHolidayYearLabel(requestYearStart)}.`
        });
        return;
      }
    }

    store.holidayRequests = Array.isArray(store.holidayRequests) ? store.holidayRequests : [];
    store.holidayRequests.unshift(nextRequest);
    const savedStore = await writeStore(store);
    const visiblePayload = await getHolidayPayload(request.user, requestYearStart);
    broadcast("board-updated", buildBoardRowsFromStore(savedStore));
    response.json(visiblePayload);
  });

  app.patch("/api/holiday-requests/:id", async (request, response) => {
    if (!requireHolidayAdmin(request, response)) return;
    const status = String(request.body?.status || "").trim().toLowerCase();
    if (!["approved", "rejected"].includes(status)) {
      response.status(400).json({ error: "Status must be approved or rejected." });
      return;
    }

    const store = await readStore();
    store.holidayRequests = Array.isArray(store.holidayRequests) ? store.holidayRequests : [];
    const requestIndex = store.holidayRequests.findIndex((item) => String(item.id || "") === String(request.params.id || ""));
    if (requestIndex === -1) {
      response.status(404).json({ error: "Holiday request not found." });
      return;
    }

    const holidayRequest = {
      ...store.holidayRequests[requestIndex],
      status,
      reviewedAt: new Date().toISOString(),
      reviewedBy: request.user?.displayName || "",
      updatedAt: new Date().toISOString()
    };
    store.holidayRequests[requestIndex] = holidayRequest;

    if (status === "approved") {
      const requestDates = enumerateIsoDates(holidayRequest.startDate, holidayRequest.endDate).filter(isWeekdayIsoDate);
      for (const date of requestDates) {
        const holidayEntry = sanitizeStaffHoliday({
          id: makeId(),
          date,
          person: holidayRequest.person,
          duration: holidayRequest.duration
        });

        const existingIndex = store.holidays.findIndex(
          (entry) => String(entry.date || "") === date && String(entry.person || "").trim().toLowerCase() === String(holidayRequest.person || "").trim().toLowerCase()
        );

        if (existingIndex >= 0) {
          holidayEntry.id = store.holidays[existingIndex].id;
          holidayEntry.createdAt = store.holidays[existingIndex].createdAt || holidayEntry.createdAt;
          store.holidays[existingIndex] = holidayEntry;
        } else {
          store.holidays.unshift(holidayEntry);
        }
      }
    }

    const savedStore = await writeStore(store);
    const payloadYearStart = getHolidayYearStartForIsoDate(holidayRequest.startDate) || getCurrentHolidayYearStart();
    const visiblePayload = await getHolidayPayload(request.user, payloadYearStart);
    broadcast("board-updated", buildBoardRowsFromStore(savedStore));
    response.json(visiblePayload);
  });

  app.post("/api/holiday-allowances", async (request, response) => {
    if (!requireHolidayAdmin(request, response)) return;

    const nextAllowance = sanitizeHolidayAllowance(request.body || {});
    if (!getHolidayStaffEntry(nextAllowance.person)) {
      response.status(400).json({ error: "A valid employee is required." });
      return;
    }

    const store = await readStore();
    store.holidayAllowances = Array.isArray(store.holidayAllowances) ? store.holidayAllowances : [];
    const existingIndex = store.holidayAllowances.findIndex(
      (entry) =>
        Number(entry.yearStart || 0) === Number(nextAllowance.yearStart || 0) &&
        String(entry.person || "").trim().toLowerCase() === String(nextAllowance.person || "").trim().toLowerCase()
    );

    if (existingIndex >= 0) {
      nextAllowance.id = store.holidayAllowances[existingIndex].id;
      nextAllowance.createdAt = store.holidayAllowances[existingIndex].createdAt || nextAllowance.createdAt;
      store.holidayAllowances[existingIndex] = nextAllowance;
    } else {
      store.holidayAllowances.unshift(nextAllowance);
    }

    await writeStore(store);
    const payload = await getHolidayPayload(request.user, nextAllowance.yearStart);
    response.json(payload);
  });

  app.post("/api/holidays", async (request, response) => {
    if (!requireBoardAdmin(request, response)) return;
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
      board: buildBoardRowsFromStore(savedStore)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.delete("/api/holidays/:id", async (request, response) => {
    if (!requireBoardAdmin(request, response)) return;
    const store = await readStore();
    store.holidays = store.holidays.filter((item) => item.id !== request.params.id);
    const savedStore = await writeStore(store);
    const payload = {
      jobs: savedStore.jobs,
      holidays: savedStore.holidays,
      board: buildBoardRowsFromStore(savedStore)
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
