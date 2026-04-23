const express = require("express");
const cors = require("cors");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const zlib = require("node:zlib");
const {
  bootstrapPasswordsFromEnv,
  createUser,
  deleteUser,
  ensureUsersFile,
  readUsersStore,
  sanitizeUser,
  setUserPasswordById,
  updateUserAttendanceProfile,
  updateUserPermissions,
  updateUserProfile,
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
const DEFAULT_SOCIAL_TONE_FILE = path.join(__dirname, "..", "data", "social-tone-matt-rutlidge.xlsx");
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
  { code: "MR", name: "Matt Rutlidge", person: "Matt R", birthDate: "" },
  { code: "DD", name: "Dawn Dewhurst", person: "Dawn D", birthDate: "" },
  { code: "TVB", name: "Tom Van-Boyd", person: "Tom V-B", birthDate: "" },
  { code: "AH", name: "Amber Hardman", person: "Amber H", birthDate: "" },
  { code: "ED", name: "Eddy D'Antonio", person: "Eddy D'A", birthDate: "" },
  { code: "PM", name: "Paul Morris", person: "Paul M", birthDate: "" },
  { code: "KW", name: "Kyle Wright", person: "Kyle W", birthDate: "" },
  { code: "MC", name: "Matt Carroll", person: "Matt C", birthDate: "" },
  { code: "KC", name: "Keilan Curtis", person: "Keilan C", birthDate: "" }
];
const HOLIDAY_RESET_VERSION = 1;

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

function getJobUploadsDir() {
  return path.join(path.dirname(getDataFile()), "job-photos");
}

async function ensureJobUploadsDir() {
  const directory = getJobUploadsDir();
  await fsp.mkdir(directory, { recursive: true });
  return directory;
}

function getRamsUploadsDir() {
  return path.join(path.dirname(getDataFile()), "rams-pdfs");
}

async function ensureRamsUploadsDir() {
  const directory = getRamsUploadsDir();
  await fsp.mkdir(directory, { recursive: true });
  return directory;
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

function canAccessAttendance(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "attendance", user?.role === "host" ? "admin" : "user") !== "none";
}

function canEditAttendance(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "attendance", user?.role === "host" ? "admin" : "user") === "admin";
}

function canAccessMileage(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "mileage", user?.role === "host" ? "admin" : "user") !== "none";
}

function canEditMileage(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "mileage", user?.role === "host" ? "admin" : "user") === "admin";
}

function canAccessVanEstimator(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "vanEstimator", "none") !== "none";
}

function canAccessRams(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "rams", user?.role === "host" ? "admin" : "none") !== "none";
}

function canAccessSocialPost(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "socialPost", user?.role === "host" ? "admin" : "none") !== "none";
}

function canEditSocialPost(user) {
  if (canManagePermissions(user)) return true;
  return getUserPermission(user, "socialPost", user?.role === "host" ? "admin" : "none") === "admin";
}

function toPublicRamsProfile(user = {}) {
  const safeUser = sanitizeUser(user);
  return {
    id: safeUser.id,
    displayName: safeUser.displayName,
    jobTitle: safeUser.jobTitle || "",
    phoneNumber: safeUser.phoneNumber || "",
    qualifications: Array.isArray(safeUser.qualifications) ? safeUser.qualifications : [],
    photoDataUrl: safeUser.photoDataUrl || ""
  };
}

function canManagePermissions(user) {
  return String(user?.displayName || "").trim().toLowerCase() === "matt rutlidge";
}

function toHolidayPersonLabel(displayName) {
  const normalized = String(displayName || "").trim();
  if (!normalized) return "";
  const parts = normalized.split(/\s+/).filter(Boolean);
  if (parts.length >= 2) {
    return `${parts[0]} ${String(parts[1] || "").charAt(0)}${parts.length > 2 && String(parts[1] || "").length <= 2 ? parts.slice(2).map((part) => `-${String(part).charAt(0)}`).join("") : ""}`.trim();
  }
  return normalized;
}

function toHolidayCode(displayName) {
  return String(displayName || "")
    .trim()
    .split(/[\s'-]+/)
    .filter(Boolean)
    .map((part) => String(part).charAt(0).toUpperCase())
    .join("")
    .slice(0, 4);
}

function buildHolidayStaffList(users = [], allowanceEntries = []) {
  const merged = [];
  const seen = new Set();
  const activeUserNames = new Set(
    users.map((user) => String(user.displayName || "").trim().toLowerCase()).filter(Boolean)
  );

  function addEntry(entry) {
    const normalizedName = String(entry?.name || entry?.fullName || "").trim();
    const normalizedPerson = String(entry?.person || "").trim();
    const normalizedCode = String(entry?.code || "").trim();
    if (!normalizedName || seen.has(normalizedName.toLowerCase())) return;
    merged.push({
      ...entry,
      name: normalizedName,
      fullName: entry.fullName || normalizedName,
      person: normalizedPerson || toHolidayPersonLabel(normalizedName)
    });
    seen.add(normalizedName.toLowerCase());
    if (normalizedPerson) seen.add(normalizedPerson.toLowerCase());
    if (normalizedCode) seen.add(normalizedCode.toLowerCase());
  }

  HOLIDAY_STAFF.forEach((entry) => {
    if (activeUserNames.has(String(entry.name || "").trim().toLowerCase())) {
      addEntry({ ...entry, fullName: entry.name });
    }
  });

  const sourceNames = users.map((user) => user.displayName);

  sourceNames.forEach((displayName) => {
    const normalizedName = String(displayName || "").trim();
    if (!normalizedName) return;
    if (seen.has(normalizedName.toLowerCase())) return;

    const code = toHolidayCode(normalizedName) || normalizedName.slice(0, 2).toUpperCase();
    const person = toHolidayPersonLabel(normalizedName);
    const entry = {
      code,
      name: normalizedName,
      fullName: normalizedName,
      person,
      birthDate: ""
    };
    merged.push(entry);
    seen.add(normalizedName.toLowerCase());
    seen.add(code.toLowerCase());
    seen.add(person.toLowerCase());
  });

  return merged;
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
  return HOLIDAY_STAFF.find((entry) => entry.name.toLowerCase() === String(displayName || "").trim().toLowerCase())?.person || toHolidayPersonLabel(displayName);
}

function formatHolidayRequestDateRange(startDate, endDate) {
  const start = String(startDate || "").trim();
  const end = String(endDate || startDate || "").trim();
  if (!start) return "";
  return end && end !== start ? `${start} to ${end}` : start;
}

function getHolidayNotificationRecipients(users, excludedUserIds = []) {
  const excluded = new Set(excludedUserIds.map((value) => String(value || "")));
  return users.filter((user) => canEditHolidays(sanitizeUser(user)) && !excluded.has(String(user.id || "")));
}

function getBoardNotificationRecipients(users) {
  return users.filter((user) => canAccessBoard(sanitizeUser(user)));
}

function getBoardEditorNotificationRecipients(users) {
  return users.filter((user) => canEditBoard(sanitizeUser(user)));
}

function getBoardLinkForUser(user, jobId = "") {
  const basePath = canEditBoard(sanitizeUser(user)) ? "/board" : "/client/board";
  if (!jobId) return basePath;
  const params = new URLSearchParams({ job: String(jobId) });
  return `${basePath}?${params.toString()}`;
}

function formatBoardNotificationDate(isoDate) {
  const parsed = parseIsoDate(isoDate);
  if (!parsed) return String(isoDate || "");
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    timeZone: TIME_ZONE
  }).format(parsed);
}

function getJobNotificationLabel(job) {
  return String(job?.orderReference || "").trim() || String(job?.customerName || "").trim() || "Job";
}

function getJobNotificationSummary(job) {
  const parts = [
    String(job?.orderReference || "").trim(),
    String(job?.customerName || "").trim(),
    String(job?.description || "").trim()
  ].filter(Boolean);
  return parts.join(" - ") || getJobNotificationLabel(job);
}

function pushBoardNotification(store, users, buildPayload) {
  store.notifications = Array.isArray(store.notifications) ? store.notifications : [];
  getBoardNotificationRecipients(users).forEach((user) => {
    const payload = buildPayload(user) || {};
    store.notifications.unshift(
      createNotification({
        userId: user.id,
        link: getBoardLinkForUser(user, payload.jobId),
        ...payload
      })
    );
  });
}

function pushBoardEditorNotification(store, users, buildPayload) {
  store.notifications = Array.isArray(store.notifications) ? store.notifications : [];
  getBoardEditorNotificationRecipients(users).forEach((user) => {
    const payload = buildPayload(user) || {};
    store.notifications.unshift(
      createNotification({
        userId: user.id,
        link: getBoardLinkForUser(user, payload.jobId),
        ...payload
      })
    );
  });
}

function getHolidayStaffIdentityKey(value) {
  const match = getHolidayStaffEntry(value);
  if (match?.code) return match.code.toLowerCase();
  return String(value || "").trim().toLowerCase();
}

function requireBoardAccess(request, response) {
  if (canAccessBoard(request.user)) return true;
  response.status(403).json({ error: "Board access required." });
  return false;
}

function requireBoardOrRamsAccess(request, response) {
  if (canAccessBoard(request.user) || canAccessRams(request.user)) return true;
  response.status(403).json({ error: "Board or RAMS access required." });
  return false;
}

function requireRamsAccess(request, response) {
  if (canAccessRams(request.user)) return true;
  response.status(403).json({ error: "RAMS access required." });
  return false;
}

function requireSocialPostAccess(request, response) {
  if (canAccessSocialPost(request.user)) return true;
  response.status(403).json({ error: "Social Post access required." });
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

function requireAttendanceAccess(request, response) {
  if (canAccessAttendance(request.user)) return true;
  response.status(403).json({ error: "Attendance access required." });
  return false;
}

function requireAttendanceAdmin(request, response) {
  if (canEditAttendance(request.user)) return true;
  response.status(403).json({ error: "Attendance admin access required." });
  return false;
}

function requireMileageAccess(request, response) {
  if (canAccessMileage(request.user)) return true;
  response.status(403).json({ error: "Mileage access required." });
  return false;
}

function requireMileageAdmin(request, response) {
  if (canEditMileage(request.user)) return true;
  response.status(403).json({ error: "Mileage admin access required." });
  return false;
}

function ensureStoreFile() {
  const file = getDataFile();
  fs.mkdirSync(path.dirname(file), { recursive: true });
    if (!fs.existsSync(file)) {
      fs.writeFileSync(
        file,
        `${JSON.stringify({ jobs: [], holidays: [], holidayRequests: [], holidayAllowances: [], holidayEvents: [], notifications: [], attendanceEntries: [], mileageClaims: [], socialPostToneVoices: [], socialPostDeletedToneVoiceIds: [] }, null, 2)}\n`,
        "utf8"
      );
    }
  }

function readHolidaySeed() {
    if (!fs.existsSync(DEFAULT_HOLIDAY_SEED_FILE)) {
    return { holidays: [], holidayAllowances: [], holidayEvents: [] };
  }

  try {
    const raw = fs.readFileSync(DEFAULT_HOLIDAY_SEED_FILE, "utf8");
    const parsed = JSON.parse(raw);
      return {
        holidays: Array.isArray(parsed.holidays) ? parsed.holidays : [],
      holidayAllowances: Array.isArray(parsed.holidayAllowances) ? parsed.holidayAllowances : [],
      holidayEvents: Array.isArray(parsed.holidayEvents) ? parsed.holidayEvents : []
      };
  } catch (error) {
    console.error("Invalid holiday seed JSON, ignoring seed.", error);
    return { holidays: [], holidayAllowances: [], holidayEvents: [] };
  }
}

function mergeHolidaySeed(store) {
  const seed = readHolidaySeed();
      const nextStore = {
        jobs: Array.isArray(store.jobs) ? store.jobs : [],
        holidays: Array.isArray(store.holidays) ? [...store.holidays] : [],
        holidayRequests: Array.isArray(store.holidayRequests) ? [...store.holidayRequests] : [],
        holidayAllowances: Array.isArray(store.holidayAllowances) ? [...store.holidayAllowances] : [],
        holidayEvents: Array.isArray(store.holidayEvents) ? [...store.holidayEvents] : [],
        notifications: Array.isArray(store.notifications) ? [...store.notifications] : [],
        attendanceEntries: Array.isArray(store.attendanceEntries) ? [...store.attendanceEntries] : [],
        mileageClaims: Array.isArray(store.mileageClaims) ? [...store.mileageClaims] : [],
        socialPostToneVoices: Array.isArray(store.socialPostToneVoices) ? [...store.socialPostToneVoices] : [],
        socialPostDeletedToneVoiceIds: Array.isArray(store.socialPostDeletedToneVoiceIds) ? [...store.socialPostDeletedToneVoiceIds] : []
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

function applyHolidayResetMigration(store) {
  const currentVersion = Number(store?.holidayResetVersion || 0);
  if (currentVersion >= HOLIDAY_RESET_VERSION) {
    return {
      ...store,
      holidayResetVersion: currentVersion
    };
  }

    return {
      ...store,
      holidays: [],
      holidayRequests: [],
      holidayAllowances: Array.isArray(store.holidayAllowances)
      ? store.holidayAllowances.map((entry) => ({
          ...entry,
          birthDate: ""
        }))
        : [],
      holidayEvents: [],
      notifications: Array.isArray(store.notifications) ? store.notifications : [],
      holidayResetVersion: HOLIDAY_RESET_VERSION
    };
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
          const migrated = applyHolidayResetMigration({
            jobs: parsed,
            holidays: [],
            holidayRequests: [],
            holidayAllowances: [],
            holidayEvents: [],
            notifications: [],
            attendanceEntries: [],
            mileageClaims: [],
            socialPostToneVoices: [],
            socialPostDeletedToneVoiceIds: []
          });
      if (Number(migrated.holidayResetVersion || 0) !== 0) {
        await writeStore(migrated);
      }
      return mergeHolidaySeed(migrated);
    }
        const migrated = applyHolidayResetMigration({
          jobs: Array.isArray(parsed.jobs) ? parsed.jobs : [],
          holidays: Array.isArray(parsed.holidays) ? parsed.holidays : [],
          holidayRequests: Array.isArray(parsed.holidayRequests) ? parsed.holidayRequests : [],
          holidayAllowances: Array.isArray(parsed.holidayAllowances) ? parsed.holidayAllowances : [],
          holidayEvents: Array.isArray(parsed.holidayEvents) ? parsed.holidayEvents : [],
          notifications: Array.isArray(parsed.notifications) ? parsed.notifications : [],
          attendanceEntries: Array.isArray(parsed.attendanceEntries) ? parsed.attendanceEntries : [],
          mileageClaims: Array.isArray(parsed.mileageClaims) ? parsed.mileageClaims : [],
          socialPostToneVoices: Array.isArray(parsed.socialPostToneVoices) ? parsed.socialPostToneVoices : [],
          socialPostDeletedToneVoiceIds: Array.isArray(parsed.socialPostDeletedToneVoiceIds) ? parsed.socialPostDeletedToneVoiceIds : [],
          holidayResetVersion: Number(parsed.holidayResetVersion || 0)
        });
    if (Number(migrated.holidayResetVersion || 0) !== Number(parsed.holidayResetVersion || 0)) {
      await writeStore(migrated);
    }
    return mergeHolidaySeed(migrated);
  } catch (error) {
    console.error("Invalid board store JSON, returning empty store.", error);
        const migrated = applyHolidayResetMigration({
          jobs: [],
          holidays: [],
          holidayRequests: [],
          holidayAllowances: [],
          holidayEvents: [],
          notifications: [],
          attendanceEntries: [],
          mileageClaims: [],
          socialPostToneVoices: [],
          socialPostDeletedToneVoiceIds: []
        });
    await writeStore(migrated);
    return mergeHolidaySeed(migrated);
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
      }),
      holidayEvents: [...(store.holidayEvents || [])].sort((left, right) => {
        if (left.date !== right.date) return String(left.date || "").localeCompare(String(right.date || ""));
        return String(left.title || "").localeCompare(String(right.title || ""));
      }),
      attendanceEntries: [...(store.attendanceEntries || [])]
        .map((entry) => sanitizeAttendanceEntry(entry))
        .sort((left, right) => {
          if (left.date !== right.date) return String(left.date || "").localeCompare(String(right.date || ""));
          return String(left.person || "").localeCompare(String(right.person || ""));
        }),
      mileageClaims: [...(store.mileageClaims || [])]
        .map((entry) => sanitizeMileageClaim(entry))
        .sort((left, right) => {
          if (left.monthId !== right.monthId) return String(right.monthId || "").localeCompare(String(left.monthId || ""));
          return String(left.userName || "").localeCompare(String(right.userName || ""));
        }),
      socialPostToneVoices: Array.isArray(store.socialPostToneVoices) ? store.socialPostToneVoices : [],
      socialPostDeletedToneVoiceIds: Array.isArray(store.socialPostDeletedToneVoiceIds) ? store.socialPostDeletedToneVoiceIds : [],
      notifications: [...(store.notifications || [])].sort((left, right) =>
        String(right.createdAt || "").localeCompare(String(left.createdAt || ""))
      ),
      holidayResetVersion: Number(store.holidayResetVersion || HOLIDAY_RESET_VERSION)
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

function getCurrentMonthId() {
  return toMonthId(getStartOfMonth(getTodayInLondon()));
}

function formatMileageMonthLabel(monthId) {
  const parsed = parseMonthId(monthId);
  return parsed ? monthFormatter.format(parsed) : String(monthId || "");
}

function extractUkPostcode(value) {
  const match = String(value || "")
    .toUpperCase()
    .match(/\b([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})\b/);
  return match ? match[1].replace(/\s+/g, "") : "";
}

async function geocodeMileagePostcode(query) {
  const postcode = extractUkPostcode(query);
  if (!postcode) return null;
  try {
    const response = await fetch(`https://api.postcodes.io/postcodes/${encodeURIComponent(postcode)}`, {
      headers: { Accept: "application/json" }
    });
    if (!response.ok) return null;
    const payload = await response.json();
    const lat = Number(payload?.result?.latitude);
    const lon = Number(payload?.result?.longitude);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    return { lat, lon, label: payload?.result?.postcode || postcode };
  } catch (error) {
    return null;
  }
}

async function geocodeMileageLocation(query) {
  const normalized = String(query || "").trim();
  if (!normalized) return null;

  const postcodeMatch = await geocodeMileagePostcode(normalized);
  if (postcodeMatch) return postcodeMatch;

  const queryAttempts = [
    normalized,
    `${normalized}, UK`,
    `${normalized}, United Kingdom`
  ];

  for (const attempt of [...new Set(queryAttempts)]) {
    try {
      const url = new URL("https://nominatim.openstreetmap.org/search");
      url.searchParams.set("format", "jsonv2");
      url.searchParams.set("addressdetails", "1");
      url.searchParams.set("limit", "1");
      url.searchParams.set("countrycodes", "gb");
      url.searchParams.set("q", attempt);

      const response = await fetch(url, {
        headers: {
          Accept: "application/json",
          "Accept-Language": "en-GB,en;q=0.9",
          "User-Agent": "sxpreston-mileage/1.0 (https://www.sxpreston.com)"
        }
      });
      if (!response.ok) continue;
      const payload = await response.json();
      const first = Array.isArray(payload) ? payload[0] : null;
      const lat = Number(first?.lat);
      const lon = Number(first?.lon);
      if (Number.isFinite(lat) && Number.isFinite(lon)) {
        return { lat, lon, label: first?.display_name || attempt };
      }
    } catch (error) {
      continue;
    }
  }

  return null;
}

async function estimateDrivingMiles(from, to) {
  const [origin, destination] = await Promise.all([
    geocodeMileageLocation(from),
    geocodeMileageLocation(to)
  ]);

  if (!origin || !destination) {
    return { miles: 0, resolved: false, message: "Could not resolve one of those addresses." };
  }

  const routeUrl = new URL(
    `https://router.project-osrm.org/route/v1/driving/${origin.lon},${origin.lat};${destination.lon},${destination.lat}`
  );
  routeUrl.searchParams.set("overview", "false");
  const response = await fetch(routeUrl, { headers: { Accept: "application/json" } });
  if (!response.ok) {
    return { miles: 0, resolved: false, message: "Could not calculate a driving route." };
  }
  const payload = await response.json();
  const distanceMeters = Number(payload?.routes?.[0]?.distance);
  if (!Number.isFinite(distanceMeters)) {
    return { miles: 0, resolved: false, message: "Could not calculate a driving route." };
  }

  return {
    miles: Math.round((distanceMeters / 1609.344) * 10) / 10,
    resolved: true,
    origin: origin.label,
    destination: destination.label
  };
}

function getDistanceKmBetween(left, right) {
  const earthRadiusKm = 6371;
  const leftLat = Number(left?.lat);
  const leftLon = Number(left?.lon);
  const rightLat = Number(right?.lat);
  const rightLon = Number(right?.lon);
  if (![leftLat, leftLon, rightLat, rightLon].every(Number.isFinite)) return 0;
  const toRadians = (value) => (value * Math.PI) / 180;
  const dLat = toRadians(rightLat - leftLat);
  const dLon = toRadians(rightLon - leftLon);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRadians(leftLat)) * Math.cos(toRadians(rightLat)) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return earthRadiusKm * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function formatHospitalAddress(tags = {}) {
  const explicit = String(tags["addr:full"] || tags.address || tags["contact:address"] || "").trim();
  if (explicit) return explicit;
  return [
    tags["addr:housename"],
    tags["addr:housenumber"],
    tags["addr:street"],
    tags["addr:place"],
    tags["addr:suburb"],
    tags["addr:city"],
    tags["addr:county"],
    tags["addr:postcode"]
  ].filter(Boolean).join(", ");
}

async function reverseGeocodeHospitalAddress(lat, lon) {
  if (!Number.isFinite(lat) || !Number.isFinite(lon)) return "";
  try {
    const url = new URL("https://nominatim.openstreetmap.org/reverse");
    url.searchParams.set("format", "jsonv2");
    url.searchParams.set("lat", String(lat));
    url.searchParams.set("lon", String(lon));
    url.searchParams.set("zoom", "18");
    url.searchParams.set("addressdetails", "1");
    const response = await fetch(url, {
      headers: {
        Accept: "application/json",
        "User-Agent": "sxpreston-rams/1.0 (https://www.sxpreston.com)"
      }
    });
    if (!response.ok) return "";
    const payload = await response.json();
    const address = payload?.address || {};
    const parts = [
      address.road || address.pedestrian || address.footway,
      address.suburb || address.neighbourhood,
      address.city || address.town || address.village,
      address.county,
      address.postcode
    ].filter(Boolean);
    return parts.join(", ") || String(payload?.display_name || "").trim();
  } catch (error) {
    return "";
  }
}

async function findNearestHospitals(address) {
  const origin = await geocodeMileageLocation(address);
  if (!origin) {
    return { origin: null, hospitals: [], message: "Could not resolve the installation address." };
  }

  const query = `
    [out:json][timeout:25];
    (
      node["amenity"="hospital"](around:40000,${origin.lat},${origin.lon});
      way["amenity"="hospital"](around:40000,${origin.lat},${origin.lon});
      relation["amenity"="hospital"](around:40000,${origin.lat},${origin.lon});
    );
    out center tags;
  `;

  const response = await fetch("https://overpass-api.de/api/interpreter", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
      Accept: "application/json",
      "User-Agent": "sxpreston-rams/1.0 (https://www.sxpreston.com)"
    },
    body: new URLSearchParams({ data: query })
  });
  if (!response.ok) {
    return { origin, hospitals: [], message: "Could not load nearby hospitals." };
  }

  const payload = await response.json();
  const seen = new Set();
  const hospitals = await Promise.all((Array.isArray(payload?.elements) ? payload.elements : [])
    .map((element) => {
      const tags = element.tags || {};
      const lat = Number(element.lat ?? element.center?.lat);
      const lon = Number(element.lon ?? element.center?.lon);
      const name = String(tags.name || tags.operator || "Hospital").trim();
      if (!name || !Number.isFinite(lat) || !Number.isFinite(lon)) return null;
      const dedupeKey = `${name.toLowerCase()}-${Math.round(lat * 1000)}-${Math.round(lon * 1000)}`;
      if (seen.has(dedupeKey)) return null;
      seen.add(dedupeKey);
      const distanceKm = getDistanceKmBetween(origin, { lat, lon });
      const addressLabel = formatHospitalAddress(tags);
      return {
        id: String(element.id || dedupeKey),
        name,
        address: addressLabel,
        lat,
        lon,
        distanceKm: Math.round(distanceKm * 10) / 10
      };
    })
    .filter(Boolean));

  const nearestHospitalCandidates = hospitals
    .sort((left, right) => left.distanceKm - right.distanceKm)
    .slice(0, 5);

  const enrichedHospitals = await Promise.all(nearestHospitalCandidates.map(async (hospital) => {
    const addressLabel = hospital.address || await reverseGeocodeHospitalAddress(hospital.lat, hospital.lon);
    const distanceLabel = Number.isFinite(hospital.distanceKm) ? ` (${hospital.distanceKm} km)` : "";
    return {
      id: hospital.id,
      name: hospital.name,
      address: addressLabel,
      distanceKm: hospital.distanceKm,
      label: `${hospital.name}${addressLabel ? ` - ${addressLabel}` : ""}${distanceLabel}`
    };
  }));

  return { origin, hospitals: enrichedHospitals };
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
  const displayHolidayEvents = (options.holidayEvents || []).filter((event) => {
    const eventDate = String(event?.date || "");
    return eventDate >= toIsoDate(start) && eventDate <= toIsoDate(end);
  });

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
  const holidayEventsByDate = displayHolidayEvents.reduce((map, entry) => {
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
      const sameDayHolidayEvents = (holidayEventsByDate.get(isoDate) || []).sort((left, right) =>
        String(left.title || "").localeCompare(String(right.title || ""))
      );

      return {
        isoDate,
      dayLabel: weekdayFormatter.format(date).toUpperCase(),
      dayNumber: String(date.getUTCDate()).padStart(2, "0"),
      fullDateLabel: longDateFormatter.format(date),
        bankHoliday: holidayMap.get(isoDate) || "",
        staffHolidays: sameDayStaffHolidays,
        holidayEvents: sameDayHolidayEvents,
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
  return buildBoardRows(store.jobs, store.holidays, {
    ...options,
    today,
    mode,
    start: rangeStart,
    end: rangeEnd,
    birthdayEntries,
    holidayEvents: store.holidayEvents || []
  });
}

function sanitizeJobPhoto(payload) {
  return {
    id: String(payload.id || makeId()),
    storageName: String(payload.storageName || "").trim(),
    fileName: String(payload.fileName || "job-photo.jpg").trim() || "job-photo.jpg",
    contentType: String(payload.contentType || "image/jpeg").trim() || "image/jpeg",
    size: Math.max(0, Number(payload.size) || 0),
    width: Math.max(0, Number(payload.width) || 0),
    height: Math.max(0, Number(payload.height) || 0),
    uploadedAt: String(payload.uploadedAt || new Date().toISOString()),
    uploadedByName: String(payload.uploadedByName || "").trim()
  };
}

function sanitizeRamsDocument(payload = {}) {
  return {
    id: String(payload.id || makeId()),
    reference: String(payload.reference || "").trim(),
    createdAt: String(payload.createdAt || new Date().toISOString()),
    updatedAt: String(payload.updatedAt || new Date().toISOString()),
    createdByName: String(payload.createdByName || "").trim(),
    pdfStorageName: String(payload.pdfStorageName || "").trim(),
    pdfFileName: String(payload.pdfFileName || "").trim(),
    pdfSize: Math.max(0, Number(payload.pdfSize) || 0),
    questions: payload.questions && typeof payload.questions === "object" ? payload.questions : {},
    cardOrder: Array.isArray(payload.cardOrder) ? payload.cardOrder.map(String).filter(Boolean) : [],
    edits: payload.edits && typeof payload.edits === "object" ? payload.edits : {},
    viewPayload: payload.viewPayload && typeof payload.viewPayload === "object" ? payload.viewPayload : {}
  };
}

function buildRamsPdfFileName(job, documentId, existingDocuments = []) {
  const baseReference = String(job?.orderReference || job?.id || "RAMS").trim() || "RAMS";
  const safeBase = `${baseReference} RAMS`.replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, " ").trim();
  const used = new Set(
    existingDocuments
      .filter((document) => String(document?.id || "") !== String(documentId || ""))
      .map((document) => String(document?.pdfFileName || "").trim().toLowerCase())
      .filter(Boolean)
  );
  let candidate = `${safeBase}.pdf`;
  let suffix = 1;
  while (used.has(candidate.toLowerCase())) {
    candidate = `${safeBase} (${suffix}).pdf`;
    suffix += 1;
  }
  return candidate;
}

function toPublicJob(job) {
  const normalized = sanitizeJob(job);
  return {
    ...normalized,
    photos: normalized.photos.map((photo) => ({
      ...photo,
      url: `/api/jobs/${encodeURIComponent(normalized.id)}/photos/${encodeURIComponent(photo.id)}`
    })),
    ramsDocuments: normalized.ramsDocuments
  };
}

function escapePdfText(value) {
  return String(value || "")
    .replace(/\\/g, "\\\\")
    .replace(/\(/g, "\\(")
    .replace(/\)/g, "\\)");
}

function wrapPdfText(value, maxLength = 78) {
  const words = String(value || "").split(/\s+/).filter(Boolean);
  if (!words.length) return [""];
  const lines = [];
  let current = "";
  words.forEach((word) => {
    const next = current ? `${current} ${word}` : word;
    if (next.length > maxLength && current) {
      lines.push(current);
      current = word;
    } else {
      current = next;
    }
  });
  if (current) lines.push(current);
  return lines;
}

function formatDateTimeForPdf(value) {
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return String(value || "");
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    timeZone: TIME_ZONE
  }).format(parsed);
}

function formatJobDateForPdf(value) {
  const parsed = parseIsoDate(value);
  if (!parsed) return String(value || "");
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    timeZone: TIME_ZONE
  }).format(parsed);
}

function sanitizePdfLine(value, fallback = "-") {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  return text || fallback;
}

function buildSimpleTextPdf(title, lines = []) {
  const pageWidth = 595.28;
  const pageHeight = 841.89;
  const margin = 42;
  const bottomMargin = 42;
  const lineHeight = 12;
  const objects = [];
  const pages = [];
  let currentLines = [];
  let y = pageHeight - margin;

  function addObject(body) {
    objects.push(body);
    return objects.length;
  }

  function makeStream(bodyText) {
    const content = Buffer.from(String(bodyText || ""), "utf8");
    return Buffer.concat([
      Buffer.from(`<< /Length ${content.length} >>\nstream\n`, "utf8"),
      content,
      Buffer.from("\nendstream", "utf8")
    ]);
  }

  function addPage() {
    pages.push(currentLines);
    currentLines = [];
    y = pageHeight - margin;
  }

  function writeLine(text, options = {}) {
    const size = options.size || 9;
    const font = options.bold ? "F2" : "F1";
    const gap = options.gap ?? 3;
    const maxLength = options.maxLength || (size >= 13 ? 60 : 88);
    const segments = wrapPdfText(text, maxLength);
    segments.forEach((segment) => {
      if (y < bottomMargin) addPage();
      currentLines.push({ text: segment, x: margin + (options.indent || 0), y, size, font });
      y -= lineHeight;
    });
    y -= gap;
  }

  writeLine(title || "RAMS Document", { bold: true, size: 16, gap: 8, maxLength: 58 });
  lines.forEach((line) => {
    if (line === "__GAP__") {
      y -= 8;
      return;
    }
    if (typeof line === "string") {
      writeLine(line);
      return;
    }
    writeLine(line.text, line);
  });
  if (currentLines.length) addPage();

  const pagesId = addObject("");
  const catalogId = addObject(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`);
  const fontRegularId = addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
  const fontBoldId = addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>");
  const pageIds = [];

  pages.forEach((pageLines, pageIndex) => {
    const commands = [
      "BT",
      "/F2 8 Tf",
      `${pageWidth - margin - 122} ${pageHeight - margin + 10} Td`,
      "(SX SIGNS EXPRESS) Tj",
      "ET"
    ];
    pageLines.forEach((line) => {
      commands.push("BT");
      commands.push(`/${line.font} ${line.size} Tf`);
      commands.push(`${line.x} ${line.y} Td`);
      commands.push(`(${escapePdfText(line.text)}) Tj`);
      commands.push("ET");
    });
    commands.push("BT");
    commands.push("/F1 7 Tf");
    commands.push(`${pageWidth - margin - 45} ${bottomMargin - 20} Td`);
    commands.push(`(Page ${pageIndex + 1} of ${pages.length}) Tj`);
    commands.push("ET");
    const contentId = addObject(makeStream(commands.join("\n")));
    const pageId = addObject(
      `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /Font << /F1 ${fontRegularId} 0 R /F2 ${fontBoldId} 0 R >> >> /Contents ${contentId} 0 R >>`
    );
    pageIds.push(pageId);
  });

  objects[pagesId - 1] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pageIds.length} >>`;
  objects[catalogId - 1] = `<< /Type /Catalog /Pages ${pagesId} 0 R >>`;

  const chunks = [Buffer.from("%PDF-1.4\n", "utf8")];
  const offsets = [0];
  objects.forEach((object, index) => {
    offsets.push(Buffer.concat(chunks).length);
    chunks.push(Buffer.from(`${index + 1} 0 obj\n`, "utf8"));
    chunks.push(Buffer.isBuffer(object) ? object : Buffer.from(String(object), "utf8"));
    chunks.push(Buffer.from("\nendobj\n", "utf8"));
  });
  const xrefOffset = Buffer.concat(chunks).length;
  chunks.push(Buffer.from(`xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`, "utf8"));
  offsets.slice(1).forEach((offset) => {
    chunks.push(Buffer.from(`${String(offset).padStart(10, "0")} 00000 n \n`, "utf8"));
  });
  chunks.push(Buffer.from(`trailer\n<< /Size ${objects.length + 1} /Root ${catalogId} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`, "utf8"));
  return Buffer.concat(chunks);
}

function buildRamsPdfDocument(job, document, payload = {}) {
  const lines = [];
  const meta = Array.isArray(payload.meta) ? payload.meta : [];
  const site = Array.isArray(payload.site) ? payload.site : [];
  const arrangements = Array.isArray(payload.arrangements) ? payload.arrangements : [];
  const accessMethods = Array.isArray(payload.accessMethods) ? payload.accessMethods : [];
  const tools = Array.isArray(payload.tools) ? payload.tools : [];
  const installers = Array.isArray(payload.installers) ? payload.installers : [];
  const ppe = Array.isArray(payload.ppe) ? payload.ppe : [];
  const siteHazards = sanitizePdfLine(payload.siteHazards || "N/A");
  const firstAid = payload.firstAid && typeof payload.firstAid === "object" ? payload.firstAid : {};
  const emergencyContacts = Array.isArray(payload.emergencyContacts) ? payload.emergencyContacts : [];
  const officeAddress = sanitizePdfLine(payload.officeAddress || "Unit 3, Sherdley Road, Lostock Hall, Preston PR5 5LP");
  const risks = Array.isArray(payload.risks) ? payload.risks : [];
  const methods = Array.isArray(payload.methods) ? payload.methods : [];

  lines.push({ text: `Reference: ${sanitizePdfLine(document.reference || payload.reference || job.orderReference)}`, bold: true, size: 10 });
  meta.forEach((item) => lines.push(`${sanitizePdfLine(item.label)}: ${sanitizePdfLine(item.value)}`));
  lines.push("__GAP__");
  lines.push({ text: "Site Details", bold: true, size: 12 });
  site.forEach((item) => lines.push(`${sanitizePdfLine(item.label)}: ${sanitizePdfLine(item.value)}`));
  lines.push("__GAP__");
  lines.push({ text: "Access, Tools and PPE", bold: true, size: 12 });
  if (installers.length) {
    lines.push(`Installers: ${installers.map((installer) => {
      const name = sanitizePdfLine(installer?.name, "");
      const title = sanitizePdfLine(installer?.jobTitle, "");
      const qualifications = sanitizePdfLine(installer?.qualifications, "");
      return [name, title, qualifications].filter(Boolean).join(" - ");
    }).filter(Boolean).join("; ")}`);
  }
  lines.push(`Tools: ${tools.map((item) => sanitizePdfLine(item, "")).filter(Boolean).join(", ") || "-"}`);
  lines.push(`Access: ${accessMethods.map((item) => sanitizePdfLine(item, "")).filter(Boolean).join(", ") || "-"}`);
  lines.push(`Site Specific Hazards or Information: ${siteHazards}`);
  lines.push(`PPE: ${ppe.map((item) => sanitizePdfLine(item, "")).filter(Boolean).join(", ") || "-"}`);
  lines.push(`First Aid Facilities: ${sanitizePdfLine(firstAid.facility)}`);
  lines.push(`First Aid Box Location: ${sanitizePdfLine(firstAid.boxLocation || "Signs Express Van")}`);
  lines.push("__GAP__");
  lines.push({ text: "Arrangements", bold: true, size: 12 });
  arrangements.forEach((item) => lines.push(`${sanitizePdfLine(item.label)}: ${sanitizePdfLine(item.value)}`));
  lines.push("__GAP__");
  lines.push({ text: "Risk Assessment", bold: true, size: 12 });
  risks.forEach((risk) => {
    lines.push({ text: sanitizePdfLine(risk.title), bold: true, size: 10 });
    lines.push(`Who may be harmed: ${sanitizePdfLine(risk.whoAtRisk)}`);
    lines.push(`Initial L/C/R: ${sanitizePdfLine(risk.initialL)}/${sanitizePdfLine(risk.initialC)}/${sanitizePdfLine(risk.initialR)}   Residual L/C/R: ${sanitizePdfLine(risk.residualL)}/${sanitizePdfLine(risk.residualC)}/${sanitizePdfLine(risk.residualR)}   Risk: ${sanitizePdfLine(risk.risk)}`);
    lines.push(`Responsibility: ${sanitizePdfLine(risk.responsibility)}`);
    (Array.isArray(risk.controls) ? risk.controls : []).forEach((control) => lines.push({ text: `- ${sanitizePdfLine(control, "")}`, indent: 12 }));
    lines.push("__GAP__");
  });
  lines.push({ text: "Method Statement", bold: true, size: 12 });
  methods.forEach((method) => {
    lines.push({ text: sanitizePdfLine(method.title), bold: true, size: 10 });
    (Array.isArray(method.lines) ? method.lines : []).forEach((line) => lines.push({ text: `- ${sanitizePdfLine(line, "")}`, indent: 12 }));
    lines.push("__GAP__");
  });
  lines.push({ text: "Emergency Contacts", bold: true, size: 12 });
  emergencyContacts.forEach((contact) => {
    lines.push(`${sanitizePdfLine(contact?.label || "Contact")}: ${[contact?.name, contact?.jobTitle, contact?.phone].map((item) => sanitizePdfLine(item, "")).filter(Boolean).join(" - ")}`);
  });
  lines.push(`Signs Express Office: ${officeAddress}`);
  lines.push("__GAP__");

  return buildSimpleTextPdf(payload.title || "Risk Assessment and Method Statement", lines);
}

function parseJpegDimensions(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 4 || buffer[0] !== 0xff || buffer[1] !== 0xd8) {
    return { width: 0, height: 0 };
  }

  let offset = 2;
  while (offset < buffer.length) {
    while (offset < buffer.length && buffer[offset] !== 0xff) offset += 1;
    while (offset < buffer.length && buffer[offset] === 0xff) offset += 1;
    if (offset >= buffer.length) break;
    const marker = buffer[offset];
    offset += 1;

    if (marker === 0xd9 || marker === 0xda) break;
    if (offset + 1 >= buffer.length) break;

    const segmentLength = buffer.readUInt16BE(offset);
    if (segmentLength < 2 || offset + segmentLength > buffer.length) break;

    if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
      if (offset + 7 >= buffer.length) break;
      return {
        height: buffer.readUInt16BE(offset + 3),
        width: buffer.readUInt16BE(offset + 5)
      };
    }

    offset += segmentLength;
  }

  return { width: 0, height: 0 };
}

function buildPdfDocument(job, photoAssets = []) {
  const pageWidth = 595.28;
  const pageHeight = 841.89;
  const margin = 48;
  const lineHeight = 13;
  const brandRight = pageWidth - margin;
  const objects = [];

  function addObject(body) {
    objects.push(body);
    return objects.length;
  }

  function makeStream(bodyText) {
    const content = Buffer.from(String(bodyText || ""), "utf8");
    return Buffer.concat([
      Buffer.from(`<< /Length ${content.length} >>\nstream\n`, "utf8"),
      content,
      Buffer.from("\nendstream", "utf8")
    ]);
  }

  function makeBinaryStream(dict, bodyBuffer) {
    return Buffer.concat([
      Buffer.from(`<< ${dict} /Length ${bodyBuffer.length} >>\nstream\n`, "utf8"),
      bodyBuffer,
      Buffer.from("\nendstream", "utf8")
    ]);
  }

  const pagesId = addObject("");
  const catalogId = addObject(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`);
  const fontRegularId = addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
  const fontBoldId = addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>");

  const pageIds = [];
  const uploadedBy = [...new Set(photoAssets.map((asset) => String(asset.uploadedByName || "").trim()).filter(Boolean))].join(", ") || "-";
  const completionDate = job.date ? formatJobDateForPdf(job.date) : "-";
  const detailLines = [
    { text: job.customerName || "Job Export", font: "bold", size: 16, gap: 4 },
    ...(job.description ? [{ text: job.description, font: "regular", size: 10, gap: 10 }] : []),
    { text: `Order Ref: ${job.orderReference || "-"}`, font: "regular", size: 9, gap: 0 },
    { text: `Completion Date: ${completionDate}`, font: "regular", size: 9, gap: 0 },
    { text: `Job Type: ${job.jobType === "Other" ? job.customJobType || "Other" : job.jobType || "-"}`, font: "regular", size: 9, gap: 0 },
    {
      text: `Installers: ${[
        ...(Array.isArray(job.installers) ? job.installers.filter((entry) => entry !== "Custom") : []),
        ...(job.installers?.includes?.("Custom") && job.customInstaller ? [job.customInstaller] : [])
      ].join(", ") || "-"}`,
      font: "regular",
      size: 9,
      gap: 0
    },
    { text: `Contact: ${job.contact || "-"}`, font: "regular", size: 9, gap: 0 },
    { text: `Number: ${job.number || "-"}`, font: "regular", size: 9, gap: 0 },
    { text: `Address: ${job.address || "-"}`, font: "regular", size: 9, gap: 0 },
    { text: `Photos Uploaded By: ${uploadedBy}`, font: "regular", size: 9, gap: 0 },
    { text: `Photos: ${photoAssets.length}`, font: "regular", size: 9, gap: 0 }
  ];

  let y = pageHeight - margin - 8;
  const detailCommands = [];
  detailCommands.push("BT");
  detailCommands.push("/F2 12 Tf");
  detailCommands.push(`${brandRight - 118} ${pageHeight - margin + 4} Td`);
  detailCommands.push("(SX SIGNS EXPRESS) Tj");
  detailCommands.push("ET");
  detailCommands.push("BT");
  detailCommands.push("/F1 6 Tf");
  detailCommands.push(`${brandRight - 121} ${pageHeight - margin - 8} Td`);
  detailCommands.push("(CENTRAL LANCASHIRE AND SOUTHPORT) Tj");
  detailCommands.push("ET");

  detailLines.forEach((line) => {
    wrapPdfText(line.text, 92).forEach((segment) => {
      detailCommands.push("BT");
      detailCommands.push(`/${line.font === "bold" ? "F2" : "F1"} ${line.size} Tf`);
      detailCommands.push(`${margin} ${y} Td`);
      detailCommands.push(`(${escapePdfText(segment)}) Tj`);
      detailCommands.push("ET");
      y -= lineHeight;
    });
    y -= line.gap || 8;
  });

  const firstPageImageIds = photoAssets.slice(0, 4).map((asset) =>
    addObject(
      makeBinaryStream(
        `/Type /XObject /Subtype /Image /Width ${asset.width || 1} /Height ${asset.height || 1} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode`,
        asset.buffer
      )
    )
  );

  const firstPageSlots = [
    { x: margin, y: 290, width: 230, height: 180 },
    { x: pageWidth - margin - 230, y: 290, width: 230, height: 180 },
    { x: margin, y: 70, width: 230, height: 180 },
    { x: pageWidth - margin - 230, y: 70, width: 230, height: 180 }
  ];

  firstPageImageIds.forEach((imageId, index) => {
    const asset = photoAssets[index];
    const slot = firstPageSlots[index];
    if (!asset || !slot) return;
    const ratio = Math.min(
      slot.width / Math.max(1, asset.width || 1),
      slot.height / Math.max(1, asset.height || 1),
      1
    );
    const drawWidth = Math.max(1, Math.round((asset.width || 1) * ratio));
    const drawHeight = Math.max(1, Math.round((asset.height || 1) * ratio));
    const drawX = slot.x + (slot.width - drawWidth) / 2;
    const drawY = slot.y + (slot.height - drawHeight) / 2;
    detailCommands.push("q");
    detailCommands.push(`${drawWidth} 0 0 ${drawHeight} ${drawX} ${drawY} cm`);
    detailCommands.push(`/Im${index + 1} Do`);
    detailCommands.push("Q");
  });

  const detailContentId = addObject(makeStream(detailCommands.join("\n")));
  const detailPageId = addObject(
    `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /Font << /F1 ${fontRegularId} 0 R /F2 ${fontBoldId} 0 R >> /XObject << ${firstPageImageIds.map((id, index) => `/Im${index + 1} ${id} 0 R`).join(" ")} >> >> /Contents ${detailContentId} 0 R >>`
  );
  pageIds.push(detailPageId);

  const remainingPhotos = photoAssets.slice(4);
  for (let pageIndex = 0; pageIndex < remainingPhotos.length; pageIndex += 4) {
    const pagePhotos = remainingPhotos.slice(pageIndex, pageIndex + 4);
    const imageIds = pagePhotos.map((asset) =>
      addObject(
        makeBinaryStream(
          `/Type /XObject /Subtype /Image /Width ${asset.width || 1} /Height ${asset.height || 1} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode`,
          asset.buffer
        )
      )
    );
    const slots = [
      { x: margin, y: 430, width: 230, height: 160 },
      { x: pageWidth - margin - 230, y: 430, width: 230, height: 160 },
      { x: margin, y: 160, width: 230, height: 160 },
      { x: pageWidth - margin - 230, y: 160, width: 230, height: 160 }
    ];
    const photoCommands = [];
    imageIds.forEach((imageId, index) => {
      const asset = pagePhotos[index];
      const slot = slots[index];
      const ratio = Math.min(
        slot.width / Math.max(1, asset.width || 1),
        slot.height / Math.max(1, asset.height || 1),
        1
      );
      const drawWidth = Math.max(1, Math.round((asset.width || 1) * ratio));
      const drawHeight = Math.max(1, Math.round((asset.height || 1) * ratio));
      const drawX = slot.x + (slot.width - drawWidth) / 2;
      const drawY = slot.y + (slot.height - drawHeight) / 2;
      photoCommands.push("q");
      photoCommands.push(`${drawWidth} 0 0 ${drawHeight} ${drawX} ${drawY} cm`);
      photoCommands.push(`/Im${index + 1} Do`);
      photoCommands.push("Q");
    });
    const photoContentId = addObject(makeStream(photoCommands.join("\n")));
    const photoPageId = addObject(
      `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /XObject << ${imageIds.map((id, index) => `/Im${index + 1} ${id} 0 R`).join(" ")} >> >> /Contents ${photoContentId} 0 R >>`
    );
    pageIds.push(photoPageId);
  }

  objects[pagesId - 1] = `<< /Type /Pages /Count ${pageIds.length} /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] >>`;

  let pdf = Buffer.from("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n", "binary");
  const offsets = [0];

  objects.forEach((object, index) => {
    offsets.push(pdf.length);
    const header = Buffer.from(`${index + 1} 0 obj\n`, "utf8");
    const body = Buffer.isBuffer(object) ? object : Buffer.from(String(object || ""), "utf8");
    const footer = Buffer.from("\nendobj\n", "utf8");
    pdf = Buffer.concat([pdf, header, body, footer]);
  });

  const xrefOffset = pdf.length;
  let xref = `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  offsets.slice(1).forEach((offset) => {
    xref += `${String(offset).padStart(10, "0")} 00000 n \n`;
  });
  const trailer = `trailer\n<< /Size ${objects.length + 1} /Root ${catalogId} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  pdf = Buffer.concat([pdf, Buffer.from(xref + trailer, "utf8")]);
  return pdf;
}

function sanitizeJob(payload) {
  const rawInstallers = Array.isArray(payload.installers)
    ? payload.installers.map(String)
    : typeof payload.installers === "string" && payload.installers.trim()
      ? payload.installers.split(/[,/]+/).map((item) => item.trim()).filter(Boolean)
      : [];
  const rawPhotos = Array.isArray(payload.photos) ? payload.photos.map(sanitizeJobPhoto) : [];
  const rawRamsDocuments = Array.isArray(payload.ramsDocuments)
    ? payload.ramsDocuments.map(sanitizeRamsDocument)
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
    isCompleted:
      payload.isCompleted === true ||
      String(payload.isCompleted || "").trim().toLowerCase() === "true" ||
      String(payload.isCompleted || "").trim() === "1",
    isSnagging:
      payload.isSnagging === true ||
      String(payload.isSnagging || "").trim().toLowerCase() === "true" ||
      String(payload.isSnagging || "").trim() === "1",
    completedAt: String(payload.completedAt || "").trim(),
    completedByUserId: String(payload.completedByUserId || "").trim(),
    completedByName: String(payload.completedByName || "").trim(),
    snaggingAt: String(payload.snaggingAt || "").trim(),
    snaggingByUserId: String(payload.snaggingByUserId || "").trim(),
    snaggingByName: String(payload.snaggingByName || "").trim(),
    photos: rawPhotos,
    ramsDocuments: rawRamsDocuments,
    notes: String(payload.notes || "").trim(),
    createdAt: String(payload.createdAt || new Date().toISOString()),
    updatedAt: new Date().toISOString()
  };
}

function toPublicJobs(jobs = []) {
  return jobs.map((job) => toPublicJob(job));
}

async function deleteJobPhotoFiles(job) {
  const photos = Array.isArray(job?.photos) ? job.photos : [];
  if (!photos.length) return;
  const uploadDir = getJobUploadsDir();
  await Promise.all(
    photos
      .map((photo) => String(photo?.storageName || "").trim())
      .filter(Boolean)
      .map(async (storageName) => {
        try {
          await fsp.unlink(path.join(uploadDir, storageName));
        } catch (error) {
          if (error?.code !== "ENOENT") {
            console.error(`Could not remove photo file ${storageName}.`, error);
          }
        }
      })
  );
}

async function deleteRamsPdfFiles(job) {
  const documents = Array.isArray(job?.ramsDocuments) ? job.ramsDocuments : [];
  if (!documents.length) return;
  const uploadDir = getRamsUploadsDir();
  await Promise.all(
    documents
      .map((document) => String(document?.pdfStorageName || "").trim())
      .filter(Boolean)
      .map(async (storageName) => {
        try {
          await fsp.unlink(path.join(uploadDir, storageName));
        } catch (error) {
          if (error?.code !== "ENOENT") {
            console.error("Could not delete RAMS PDF.", error);
          }
        }
      })
  );
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
  const normalizedAction = String(payload.action || "book").trim().toLowerCase();
  const action = ["book", "cancel"].includes(normalizedAction) ? normalizedAction : "book";
  return {
    id: String(payload.id || makeId()),
    person: String(payload.person || "").trim(),
    requestedByUserId: String(payload.requestedByUserId || "").trim(),
    requestedByName: String(payload.requestedByName || "").trim(),
    startDate: String(payload.startDate || "").trim(),
    endDate: String(payload.endDate || payload.startDate || "").trim(),
    duration: normalizedDuration || "Full Day",
    notes: String(payload.notes || "").trim(),
    action,
    targetRequestId: String(payload.targetRequestId || "").trim(),
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

function sanitizeAttendanceTime(value) {
  const raw = String(value || "").trim();
  const match = raw.match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return "";
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (!Number.isInteger(hours) || !Number.isInteger(minutes)) return "";
  if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) return "";
  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function sanitizeAttendanceEntry(payload) {
  return {
    id: String(payload.id || makeId()),
    person: String(payload.person || "").trim(),
    date: String(payload.date || "").trim(),
    clockIn: sanitizeAttendanceTime(payload.clockIn),
    clockOut: sanitizeAttendanceTime(payload.clockOut),
    source: String(payload.source || "manual").trim().toLowerCase() || "manual",
    adminNote: String(payload.adminNote || "").trim(),
    employeeNote: String(payload.employeeNote || "").trim(),
    missingNotificationSentAt: String(payload.missingNotificationSentAt || "").trim(),
    missingNotificationResolvedAt: String(payload.missingNotificationResolvedAt || "").trim(),
    createdAt: String(payload.createdAt || new Date().toISOString()),
    updatedAt: new Date().toISOString()
  };
}

function sanitizeMileageLine(payload) {
  const miles = Number(payload?.miles);
  return {
    id: String(payload?.id || makeId()),
    date: String(payload?.date || "").trim(),
    from: String(payload?.from || "").trim(),
    to: String(payload?.to || "").trim(),
    note: String(payload?.note || "").trim(),
    miles: Number.isFinite(miles) ? Math.max(0, Math.round(miles * 10) / 10) : 0
  };
}

function sanitizeMileageClaim(payload) {
  const lines = Array.isArray(payload?.lines)
    ? payload.lines.map((line) => sanitizeMileageLine(line)).filter((line) => line.from || line.to || line.note || line.miles)
    : [];
  const totalMiles = Math.round(lines.reduce((sum, line) => sum + Number(line.miles || 0), 0) * 10) / 10;
  return {
    id: String(payload?.id || makeId()),
    userId: String(payload?.userId || "").trim(),
    userName: String(payload?.userName || "").trim(),
    monthId: String(payload?.monthId || "").trim(),
    lines,
    totalMiles,
    status: String(payload?.status || "submitted").trim().toLowerCase() || "submitted",
    submittedAt: String(payload?.submittedAt || new Date().toISOString()),
    createdAt: String(payload?.createdAt || new Date().toISOString()),
    updatedAt: new Date().toISOString()
  };
}

function buildMileageHistory(claims) {
  const monthGroups = new Map();

  (Array.isArray(claims) ? claims : []).map((claim) => sanitizeMileageClaim(claim)).forEach((claim) => {
    claim.lines.forEach((line) => {
      const lineDate = parseIsoDate(line.date);
      const lineMonthId = lineDate ? toMonthId(getStartOfMonth(lineDate)) : claim.monthId;
      if (!lineMonthId) return;

      const existing = monthGroups.get(lineMonthId) || {
        id: `mileage-${lineMonthId}`,
        monthId: lineMonthId,
        monthLabel: formatMileageMonthLabel(lineMonthId),
        totalMiles: 0,
        lineCount: 0,
        lines: [],
        updatedAt: "",
        submittedAt: ""
      };

      const nextLine = {
        ...line,
        claimMonthId: claim.monthId
      };
      existing.lines.push(nextLine);
      existing.totalMiles = Math.round((existing.totalMiles + Number(line.miles || 0)) * 10) / 10;
      existing.lineCount += 1;
      existing.updatedAt = String(claim.updatedAt || "") > String(existing.updatedAt || "") ? claim.updatedAt : existing.updatedAt;
      existing.submittedAt = String(claim.submittedAt || "") > String(existing.submittedAt || "") ? claim.submittedAt : existing.submittedAt;
      monthGroups.set(lineMonthId, existing);
    });
  });

  return [...monthGroups.values()]
    .map((group) => ({
      ...group,
      lines: group.lines.sort((left, right) => {
        if (left.date !== right.date) return String(right.date || "").localeCompare(String(left.date || ""));
        return String(right.id || "").localeCompare(String(left.id || ""));
      })
    }))
    .sort((left, right) => String(right.monthId || "").localeCompare(String(left.monthId || "")));
}

function getMileageLineMonthId(line, fallbackMonthId = "") {
  const lineDate = parseIsoDate(line?.date);
  return lineDate ? toMonthId(getStartOfMonth(lineDate)) : String(fallbackMonthId || "").trim();
}

function buildMileageAdminOverview(claims, users, monthId) {
  const targetMonthId = parseMonthId(monthId) ? monthId : getCurrentMonthId();
  const mileageUsers = (Array.isArray(users) ? users : [])
    .map((user) => sanitizeUser(user))
    .filter((user) => canAccessMileage(user))
    .filter((user) => String(user.displayName || "").trim().toLowerCase() !== "matt rutlidge")
    .sort((left, right) => String(left.displayName || "").localeCompare(String(right.displayName || "")));

  const claimList = (Array.isArray(claims) ? claims : []).map((claim) => sanitizeMileageClaim(claim));
  const usersWithClaims = new Map(mileageUsers.map((user) => [String(user.id || ""), user]));

  const userRows = [...usersWithClaims.values()]
    .map((user) => {
      const userClaims = claimList.filter((claim) => String(claim.userId || "") === String(user.id || ""));
      const journeys = [];

      userClaims.forEach((claim) => {
        claim.lines.forEach((line) => {
          const lineMonthId = getMileageLineMonthId(line, claim.monthId);
          if (lineMonthId !== targetMonthId) return;
          journeys.push({
            ...line,
            claimMonthId: claim.monthId,
            userId: user.id,
            userName: user.displayName || claim.userName || "Unknown user"
          });
        });
      });

      journeys.sort((left, right) => {
        if (left.date !== right.date) return String(right.date || "").localeCompare(String(left.date || ""));
        return String(right.id || "").localeCompare(String(left.id || ""));
      });

      const totalMiles = Math.round(journeys.reduce((sum, line) => sum + Number(line.miles || 0), 0) * 10) / 10;
      return {
        userId: String(user.id || ""),
        userName: user.displayName || "Unknown user",
        totalMiles,
        lineCount: journeys.length,
        journeys
      };
    })
    .sort((left, right) => {
      if (right.totalMiles !== left.totalMiles) return right.totalMiles - left.totalMiles;
      return String(left.userName || "").localeCompare(String(right.userName || ""));
    });

  return {
    monthId: targetMonthId,
    monthLabel: formatMileageMonthLabel(targetMonthId),
    totalMiles: Math.round(userRows.reduce((sum, user) => sum + Number(user.totalMiles || 0), 0) * 10) / 10,
    lineCount: userRows.reduce((sum, user) => sum + Number(user.lineCount || 0), 0),
    userCount: userRows.length,
    submittedUserCount: userRows.filter((user) => Number(user.lineCount || 0) > 0).length,
    users: userRows
  };
}

function sanitizeHolidayEvent(payload) {
  return {
    id: String(payload.id || makeId()),
    date: String(payload.date || "").trim(),
    title: String(payload.title || "").trim(),
    createdAt: String(payload.createdAt || new Date().toISOString()),
    updatedAt: new Date().toISOString()
  };
}

function sanitizeNotification(payload) {
  return {
    id: String(payload.id || makeId()),
    userId: String(payload.userId || "").trim(),
    title: String(payload.title || "").trim(),
    message: String(payload.message || "").trim(),
    link: String(payload.link || "").trim(),
    type: String(payload.type || "general").trim().toLowerCase(),
    read: Boolean(payload.read),
    createdAt: String(payload.createdAt || new Date().toISOString()),
    updatedAt: new Date().toISOString()
  };
}

function createNotification(payload) {
  return sanitizeNotification({
    ...payload,
    read: false
  });
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

function isHalfDayHoliday(entry) {
  const duration = String(entry?.duration || "").trim().toLowerCase();
  return duration === "morning" || duration === "afternoon";
}

function buildBaseHolidayAllowanceRows(store, yearStart = getCurrentHolidayYearStart(), holidayStaffList = HOLIDAY_STAFF) {
  return holidayStaffList.map((staffEntry) => {
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
        `${getHolidayStaffIdentityKey(entry.person)}::${entry.date}`,
        entry
      ])
  );

  const normalized = (staffHolidays || [])
    .map((entry) => sanitizeStaffHoliday(entry))
    .filter((entry) => entry.date >= startIso && entry.date <= endIso)
      .map((entry) => {
        const key = `${getHolidayStaffIdentityKey(entry.person)}::${entry.date}`;
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
      normalized.map((entry) => `${getHolidayStaffIdentityKey(entry.person)}::${entry.date}`)
  );

  scopedBirthdayEntries.forEach((entry) => {
      const key = `${getHolidayStaffIdentityKey(entry.person)}::${entry.date}`;
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

function buildHolidayAllowanceSummaries(store, yearStart = getCurrentHolidayYearStart(), holidayStaffList = HOLIDAY_STAFF) {
  const bounds = getHolidayYearBounds(yearStart);
  const startIso = toIsoDate(bounds.start);
  const endIso = toIsoDate(bounds.end);
  const allowanceRows = buildBaseHolidayAllowanceRows(store, yearStart, holidayStaffList);
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
  const usersStore = await readUsersStore();
  const holidayStaffList = buildHolidayStaffList(usersStore.users || [], store.holidayAllowances || []);
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

  const requestsInYear = visibleRequests.filter((request) => {
    const requestStart = String(request.startDate || "");
    const requestEnd = String(request.endDate || request.startDate || "");
    return requestStart <= endIso && requestEnd >= startIso;
  });
  const yearRequests = requestsInYear.filter((request) => {
    const requestStatus = String(request.status || "pending").trim().toLowerCase();
    return canEditHolidays(forUser) ? requestStatus === "pending" : requestStatus === "pending";
  });
  const approvedHolidayRequests = canEditHolidays(forUser)
    ? []
    : requestsInYear.filter((request) => {
        const requestStatus = String(request.status || "pending").trim().toLowerCase();
        const requestAction = String(request.action || "book").trim().toLowerCase();
        return requestStatus === "approved" && requestAction === "book";
      });

  const yearHolidays = (store.holidays || []).filter((holiday) => {
    const holidayDate = String(holiday.date || "");
    return holidayDate >= startIso && holidayDate <= endIso;
  });
  const birthdayEntries = buildBirthdayHolidayEntries(buildBaseHolidayAllowanceRows(store, yearStart, holidayStaffList), yearStart);
  const displayYearHolidays = getDisplayStaffHolidays(store.holidays || [], startIso, endIso, birthdayEntries);

  const allowanceRows = buildHolidayAllowanceSummaries(store, yearStart, holidayStaffList).filter((entry) =>
    canEditHolidays(forUser)
      ? true
      : String(entry.person || "").trim().toLowerCase() === String(currentPerson || "").toLowerCase()
  );

    return {
      holidays: displayYearHolidays,
      holidayRequests: yearRequests,
      approvedHolidayRequests,
      holidayStaff: holidayStaffList,
      holidayAllowances: allowanceRows,
      holidayEvents: (store.holidayEvents || []).filter((event) => {
        const eventDate = String(event.date || "");
        return eventDate >= startIso && eventDate <= endIso;
      }),
      holidayYearStart: yearStart,
      currentHolidayYearStart: getCurrentHolidayYearStart(),
      holidayYearLabel: getHolidayYearLabel(yearStart),
    holidayYearOptions: getHolidayYearOptions(getCurrentHolidayYearStart())
  };
}

function parseMonthId(value) {
  const match = String(value || "").trim().match(/^(\d{4})-(\d{2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  if (!year || !month || month < 1 || month > 12) return null;
  return new Date(Date.UTC(year, month - 1, 1));
}

function toMonthId(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
}

function getAttendanceMonthBounds(monthId = "") {
  const target = parseMonthId(monthId) || getStartOfMonth(getTodayInLondon());
  return {
    monthId: toMonthId(target),
    start: getStartOfMonth(target),
    end: getEndOfMonth(target)
  };
}

function getAttendanceDisplayLabel(entry, bankHolidayLabel) {
  if (entry) {
    const type = getHolidayType(entry);
    if (type === "birthday") return "Birthday";
    if (type === "unpaid" || type === "unpaid-holiday" || type === "unpaid holiday") return "Unpaid holiday";
    if (type === "absence" || type === "absent") return "Absence";
    if (isHalfDayHoliday(entry)) return "";
    return "Holiday";
  }
  if (bankHolidayLabel) return bankHolidayLabel;
  return "";
}

function getAttendanceHalfDayLabel(entry) {
  if (!entry || !isHalfDayHoliday(entry)) return "";
  const type = getHolidayType(entry);
  if (type === "birthday") return "";
  const duration = String(entry.duration || "").trim();
  return `${duration} holiday`;
}

function getWeekdayKeyFromIso(isoDate) {
  const parsed = parseIsoDate(isoDate);
  const day = parsed?.getUTCDay();
  return ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"][day] || "";
}

function getAttendanceLinkForUser(user, date = "") {
  const basePath = "/attendance";
  if (!date) return basePath;
  const params = new URLSearchParams({ date: String(date) });
  return `${basePath}?${params.toString()}`;
}

function syncAttendanceMissingNotification(store, users, attendanceEntry) {
  const personKey = getHolidayStaffIdentityKey(attendanceEntry.person);
  const matchingUser = (users || []).find((user) => {
    if (!canAccessAttendance(sanitizeUser(user))) return false;
    return getHolidayStaffIdentityKey(getHolidayStaffPerson(user.displayName) || user.displayName) === personKey;
  });
  if (!matchingUser) return;

  store.notifications = Array.isArray(store.notifications) ? store.notifications : [];
  const existingNotification = store.notifications.find((notification) =>
    String(notification.type || "") === "attendance-missing" &&
    String(notification.userId || "") === String(matchingUser.id || "") &&
    String(notification.link || "") === getAttendanceLinkForUser(matchingUser, attendanceEntry.date)
  );

  const hasMissingClock = Boolean(
    (attendanceEntry.clockIn && !attendanceEntry.clockOut) ||
    (!attendanceEntry.clockIn && attendanceEntry.clockOut)
  );

  if (hasMissingClock) {
    const title = "Missing clocking data";
    const message = `${attendanceEntry.person} has missing attendance data on ${formatBoardNotificationDate(attendanceEntry.date)}. Please add a note to explain it.`;
    if (existingNotification) {
      existingNotification.read = false;
      existingNotification.title = title;
      existingNotification.message = message;
      existingNotification.updatedAt = new Date().toISOString();
      return;
    }
    store.notifications.unshift(
      createNotification({
        userId: matchingUser.id,
        type: "attendance-missing",
        title,
        message,
        link: getAttendanceLinkForUser(matchingUser, attendanceEntry.date)
      })
    );
    return;
  }

  if (existingNotification) {
    existingNotification.read = true;
    existingNotification.updatedAt = new Date().toISOString();
  }
}

async function getAttendancePayload(forUser, monthId = "") {
  const store = await readStore();
  const usersStore = await readUsersStore();
  const holidayStaffList = buildHolidayStaffList(usersStore.users || [], store.holidayAllowances || []);
  const { monthId: resolvedMonthId, start, end } = getAttendanceMonthBounds(monthId);
  const startIso = toIsoDate(start);
  const endIso = toIsoDate(end);
  const todayIso = toIsoDate(getTodayInLondon());
  const displayHolidays = getDisplayStaffHolidays(
    store.holidays || [],
    startIso,
    endIso,
    buildBirthdayEntriesForRange(store, startIso, endIso)
  );
  const holidayByKey = new Map(
    displayHolidays.map((entry) => [`${getHolidayStaffIdentityKey(entry.person)}::${entry.date}`, entry])
  );
  const bankHolidayMap = getHolidayMap(start, end);
  const attendanceEntries = (store.attendanceEntries || []).map((entry) => sanitizeAttendanceEntry(entry));
  const entryByKey = new Map(
    attendanceEntries
      .filter((entry) => entry.date >= startIso && entry.date <= endIso)
      .map((entry) => [`${getHolidayStaffIdentityKey(entry.person)}::${entry.date}`, entry])
  );

  const attendanceUsers = (usersStore.users || [])
    .map((user) => sanitizeUser(user))
    .filter((user) => canAccessAttendance(user));

  const attendanceUserByPersonKey = new Map(
    attendanceUsers.map((user) => [
      getHolidayStaffIdentityKey(getHolidayStaffPerson(user.displayName) || user.displayName),
      user
    ])
  );

  const staffEntries = holidayStaffList.filter((staffEntry) =>
    attendanceUsers.some(
      (user) => getHolidayStaffIdentityKey(getHolidayStaffPerson(user.displayName) || user.displayName) === getHolidayStaffIdentityKey(staffEntry.person)
    )
  );

  const currentPersonKey = getHolidayStaffIdentityKey(getHolidayStaffPerson(forUser?.displayName) || forUser?.displayName);
  const visibleStaff = canEditAttendance(forUser)
    ? staffEntries
    : staffEntries.filter((entry) => getHolidayStaffIdentityKey(entry.person) === currentPersonKey);
  const filteredVisibleStaff = visibleStaff.filter((entry) => {
    const user = attendanceUserByPersonKey.get(getHolidayStaffIdentityKey(entry.person));
    const mode = String(user?.attendanceProfile?.mode || "required").trim().toLowerCase();
    return mode !== "exempt";
  });

  const rows = enumerateIsoDates(startIso, endIso).map((isoDate) => {
    const parsed = parseIsoDate(isoDate);
    const weekday = parsed?.getUTCDay() ?? 0;
    const isWeekend = weekday === 0 || weekday === 6;
    const bankHolidayLabel = bankHolidayMap.get(isoDate) || "";
    return {
      isoDate,
      dateLabel: formatBoardNotificationDate(isoDate),
      weekdayLabel: parsed
        ? parsed.toLocaleDateString("en-GB", { weekday: "short", timeZone: TIME_ZONE })
        : "",
      isToday: isoDate === todayIso,
      cells: filteredVisibleStaff.map((staffEntry) => {
        const key = `${getHolidayStaffIdentityKey(staffEntry.person)}::${isoDate}`;
        const holidayEntry = holidayByKey.get(key);
        const attendanceEntry = entryByKey.get(key) || null;
        const user = attendanceUserByPersonKey.get(getHolidayStaffIdentityKey(staffEntry.person));
        const attendanceProfile = user?.attendanceProfile || { mode: "required", contractedHours: {} };
        const attendanceMode = String(attendanceProfile.mode || "required").trim().toLowerCase();
        const weekdayKey = getWeekdayKeyFromIso(isoDate);
        const contractedHours = attendanceProfile.contractedHours?.[weekdayKey] || { in: "", out: "", off: false };
        let displayLabel = getAttendanceDisplayLabel(holidayEntry, bankHolidayLabel);
        const halfDayHolidayLabel = getAttendanceHalfDayLabel(holidayEntry);
        const isWorkingDay = !displayLabel && !isWeekend;
        const isContractedOff = Boolean(contractedHours.off) && isWorkingDay;
        const usesContractedHours =
          attendanceMode === "fixed" &&
          isWorkingDay &&
          !isContractedOff &&
          contractedHours.in &&
          contractedHours.out;
        if (isContractedOff) {
          displayLabel = "Off";
        }
        if (usesContractedHours) {
          displayLabel = "";
        }
        const hasMissingClock = Boolean(
          attendanceEntry &&
          ((attendanceEntry.clockIn && !attendanceEntry.clockOut) || (!attendanceEntry.clockIn && attendanceEntry.clockOut))
        );

        return {
          person: staffEntry.person,
          code: staffEntry.code,
          fullName: staffEntry.fullName || staffEntry.name || staffEntry.person,
          attendanceMode,
          displayLabel: displayLabel || (isWeekend ? "Weekend" : ""),
          halfDayHolidayLabel,
          isHoliday: Boolean(displayLabel),
          isWeekend,
          isWorkingDay,
          clockIn: attendanceEntry?.clockIn || (usesContractedHours ? contractedHours.in : ""),
          clockOut: attendanceEntry?.clockOut || (usesContractedHours ? contractedHours.out : ""),
          adminNote: attendanceEntry?.adminNote || "",
          employeeNote: attendanceEntry?.employeeNote || "",
          hasMissingClock: usesContractedHours || isContractedOff ? false : hasMissingClock,
          entryId: attendanceEntry?.id || "",
          canExplain:
            !usesContractedHours &&
            !isContractedOff &&
            !canEditAttendance(forUser) &&
            getHolidayStaffIdentityKey(staffEntry.person) === currentPersonKey &&
            hasMissingClock
        };
      })
    };
  });

  const missingEntries = rows
    .flatMap((row) => row.cells.map((cell) => ({ ...cell, isoDate: row.isoDate, dateLabel: row.dateLabel })))
    .filter((cell) => cell.canExplain);

  return {
    monthId: resolvedMonthId,
    monthLabel: start.toLocaleDateString("en-GB", { month: "long", year: "numeric", timeZone: TIME_ZONE }),
    today: todayIso,
    staff: filteredVisibleStaff.map((entry) => ({
      person: entry.person,
      code: entry.code,
      fullName: entry.fullName || entry.name || entry.person,
      attendanceMode:
        attendanceUserByPersonKey.get(getHolidayStaffIdentityKey(entry.person))?.attendanceProfile?.mode || "required"
    })),
    rows,
    adminMode: canEditAttendance(forUser),
    missingEntries
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

function filterCoreBridgeOrders(orders, searchTerm = "", options = {}) {
  const normalizedSearch = String(searchTerm || "").trim().toLowerCase();
  const includeClosed = options.includeClosed === true;
  const filteredByStatus = includeClosed
    ? orders
    : orders.filter((order) => {
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

async function fetchCoreBridgeOrders(searchTerm = "", includeDebug = false, options = {}) {
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
        normalizedSearch,
        options
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

function sanitizeSocialToneVoice(voice = {}) {
  const examples = Array.isArray(voice.examples)
    ? voice.examples
        .map((example) => ({
          reference: String(example?.reference || "").replace(/\s+/g, " ").trim().slice(0, 80),
          post: String(example?.post || "").replace(/\u0000/g, "").trim().slice(0, 8000)
        }))
        .filter((example) => example.reference && example.post)
        .slice(0, 250)
    : [];
  return {
    id: String(voice.id || makeId()),
    name: String(voice.name || "LinkedIn").replace(/\s+/g, " ").trim().slice(0, 80) || "LinkedIn",
    fileName: String(voice.fileName || "").replace(/\s+/g, " ").trim().slice(0, 160),
    content: String(voice.content || "").replace(/\u0000/g, "").trim().slice(0, 100000),
    supportingText: String(voice.supportingText || "").replace(/\u0000/g, "").trim().slice(0, 30000),
    examples,
    createdAt: String(voice.createdAt || new Date().toISOString()),
    seeded: voice.seeded === true
  };
}

function decodeXmlEntities(value = "") {
  return String(value)
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#160;/g, " ")
    .replace(/&#xA0;/gi, " ");
}

function readZipEntries(buffer) {
  const entries = new Map();
  const eocdSignature = 0x06054b50;
  let eocdOffset = -1;
  for (let index = buffer.length - 22; index >= Math.max(0, buffer.length - 70000); index -= 1) {
    if (buffer.readUInt32LE(index) === eocdSignature) {
      eocdOffset = index;
      break;
    }
  }
  if (eocdOffset < 0) return entries;

  const entryCount = buffer.readUInt16LE(eocdOffset + 10);
  const centralDirectoryOffset = buffer.readUInt32LE(eocdOffset + 16);
  let offset = centralDirectoryOffset;
  for (let index = 0; index < entryCount; index += 1) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50) break;
    const method = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localHeaderOffset = buffer.readUInt32LE(offset + 42);
    const fileName = buffer.slice(offset + 46, offset + 46 + fileNameLength).toString("utf8");
    const localFileNameLength = buffer.readUInt16LE(localHeaderOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localHeaderOffset + 28);
    const dataStart = localHeaderOffset + 30 + localFileNameLength + localExtraLength;
    const compressed = buffer.slice(dataStart, dataStart + compressedSize);
    try {
      entries.set(fileName, method === 8 ? zlib.inflateRawSync(compressed) : compressed);
    } catch (error) {
      entries.set(fileName, Buffer.alloc(0));
    }
    offset += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

function parseXlsxToneData(buffer) {
  const entries = readZipEntries(buffer);
  if (!entries.size) return { content: "", examples: [] };
  const sharedXml = entries.get("xl/sharedStrings.xml")?.toString("utf8") || "";
  const sharedStrings = [...sharedXml.matchAll(/<si\b[\s\S]*?<\/si>/g)].map((match) =>
    decodeXmlEntities(
      [...match[0].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)]
        .map((textMatch) => textMatch[1].replace(/<[^>]+>/g, ""))
        .join("")
    )
  );

  const sheetTexts = [];
  const examples = [];
  for (const [fileName, content] of entries.entries()) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/i.test(fileName)) continue;
    const xml = content.toString("utf8");
    const rows = [...xml.matchAll(/<row\b[\s\S]*?<\/row>/g)].map((rowMatch) => {
      const cellsByColumn = {};
      [...rowMatch[0].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)].forEach((cellMatch) => {
        const attrs = cellMatch[1] || "";
        const body = cellMatch[2] || "";
        const ref = attrs.match(/\br="([A-Z]+)\d+"/i)?.[1]?.toUpperCase() || "";
        const value = body.match(/<v>([\s\S]*?)<\/v>/)?.[1] || body.match(/<t[^>]*>([\s\S]*?)<\/t>/)?.[1] || "";
        const decoded = /\bt="s"/.test(attrs)
          ? sharedStrings[Number(value)] || ""
          : decodeXmlEntities(value.replace(/<[^>]+>/g, ""));
        if (ref && decoded) cellsByColumn[ref] = decoded.trim();
      });
      const reference = String(cellsByColumn.A || "").trim();
      const post = String(cellsByColumn.B || "").trim();
      if (/^(ord|inv)-?\d+/i.test(reference) && post.length > 30) {
        examples.push({ reference, post });
      }
      const cells = Object.keys(cellsByColumn)
        .sort()
        .map((key) => cellsByColumn[key])
        .filter(Boolean);
      return cells.filter(Boolean).join(" | ");
    });
    sheetTexts.push(rows.filter(Boolean).join("\n"));
  }
  return {
    content: sheetTexts.join("\n\n").replace(/\s+\|/g, " |").trim(),
    examples
  };
}

function parseXlsxText(buffer) {
  return parseXlsxToneData(buffer).content;
}

function parseToneVoiceUpload({ name, fileName, dataUrl, text, supportingText }) {
  const rawName = String(name || "").trim();
  const rawFileName = String(fileName || "").trim();
  let content = String(text || "").trim();

  if (!content && dataUrl) {
    const [meta, base64 = ""] = String(dataUrl).split(",");
    const buffer = Buffer.from(base64, "base64");
    const lowerName = rawFileName.toLowerCase();
    if (lowerName.endsWith(".xlsx") || meta.includes("spreadsheetml")) {
      const parsed = parseXlsxToneData(buffer);
      content = parsed.content;
      return sanitizeSocialToneVoice({
        name: rawName || rawFileName.replace(/\.[^.]+$/, "") || "LinkedIn tone",
        fileName: rawFileName,
        content,
        supportingText,
        examples: parsed.examples
      });
    } else {
      content = buffer.toString("utf8");
    }
  }

  return sanitizeSocialToneVoice({
    name: rawName || rawFileName.replace(/\.[^.]+$/, "") || "LinkedIn tone",
    fileName: rawFileName,
    content,
    supportingText,
    examples: []
  });
}

let defaultSocialToneVoiceCache = null;

function getDefaultSocialToneVoice() {
  if (defaultSocialToneVoiceCache) return defaultSocialToneVoiceCache;

  try {
    if (fs.existsSync(DEFAULT_SOCIAL_TONE_FILE)) {
      const parsed = parseXlsxToneData(fs.readFileSync(DEFAULT_SOCIAL_TONE_FILE));
      defaultSocialToneVoiceCache = sanitizeSocialToneVoice({
        id: "matt-rutlidge-default",
        name: "Matt Rutlidge",
        fileName: "Tone of Voice - Linkedin - Corebridge.xlsx",
        content: parsed.content,
        supportingText: "",
        examples: parsed.examples,
        createdAt: "2026-04-22T00:00:00.000Z",
        seeded: true
      });
      return defaultSocialToneVoiceCache;
    }
  } catch (error) {
    console.error("Could not load the default Social Post tone workbook.", error.message);
  }

  defaultSocialToneVoiceCache = sanitizeSocialToneVoice({
    id: "matt-rutlidge-default",
    name: "Matt Rutlidge",
    fileName: "",
    content: "Friendly, practical LinkedIn posts for completed signage work. Use natural hooks, short paragraphs, occasional emojis and plain language.",
    supportingText: "",
    examples: [],
    createdAt: "2026-04-22T00:00:00.000Z",
    seeded: true
  });
  return defaultSocialToneVoiceCache;
}

function getSocialPostVoices(store = {}) {
  const deletedIds = new Set(
    (Array.isArray(store.socialPostDeletedToneVoiceIds) ? store.socialPostDeletedToneVoiceIds : [])
      .map((id) => String(id || ""))
      .filter(Boolean)
  );
  const savedVoices = (Array.isArray(store.socialPostToneVoices) ? store.socialPostToneVoices : []).map(sanitizeSocialToneVoice);
  const hasMattVoice = savedVoices.some(
    (voice) => String(voice.id) === "matt-rutlidge-default" || voice.name.toLowerCase() === "matt rutlidge"
  );
  const voices = hasMattVoice ? savedVoices : [getDefaultSocialToneVoice(), ...savedVoices];
  return voices.filter((voice) => !deletedIds.has(String(voice.id)));
}

function getSocialPostToneSummary(voice) {
  const text = String(voice?.content || "").replace(/\s+/g, " ").trim();
  return text.slice(0, 20000);
}

function getSocialPostToneTraits(voice) {
  return String(voice?.supportingText || "").replace(/\s+/g, " ").trim().slice(0, 12000);
}

function getSocialPostTransformationExamples(voice) {
  return (Array.isArray(voice?.examples) ? voice.examples : [])
    .map((example) => ({
      reference: example.reference,
      finishedPost: String(example.post || "").slice(0, 4000)
    }))
    .slice(0, 40);
}

function classifySocialCustomer(customerName = "") {
  const name = String(customerName || "").toLowerCase();
  if (/\b(council|borough|district council|county council|city council|parish council|local authority)\b/i.test(name)) {
    return { type: "council", label: "council", avoid: ["business", "company", "brand"] };
  }
  if (/\b(leisure|leisure trust|leisure centre|sports centre|community centre)\b/i.test(name)) {
    return { type: "leisure organisation", label: "leisure organisation", avoid: ["business", "company"] };
  }
  if (/\b(school|academy|college|university|nursery)\b/i.test(name)) {
    return { type: "education", label: "school or education provider", avoid: ["business", "company"] };
  }
  if (/\b(nhs|hospital|health|medical centre|surgery|clinic)\b/i.test(name)) {
    return { type: "healthcare", label: "healthcare organisation", avoid: ["business", "company"] };
  }
  if (/\b(charity|foundation|trust|community interest|cic|association)\b/i.test(name)) {
    return { type: "charity or trust", label: "organisation", avoid: ["business", "company"] };
  }
  if (/\b(ltd|limited|plc|llp|group|services|solutions|studio|agency|retail|restaurant|bar|cafe|shop)\b/i.test(name)) {
    return { type: "business", label: "business", avoid: [] };
  }
  return { type: "organisation", label: "organisation", avoid: ["company"] };
}

function getSocialPostAiStatus() {
  const apiKey = String(process.env.OPENAI_API_KEY || "").trim();
  return {
    configured: Boolean(apiKey),
    model: String(process.env.SOCIAL_POST_AI_MODEL || "gpt-4o-mini").trim(),
    keyLength: apiKey ? apiKey.length : 0
  };
}

const SOCIAL_POST_DESCRIPTION_FINGERPRINTS = [
  "5500mm (w) x 650mm (h) x 85mm (d) 3mm Folded Aluminium Composite Sign Tray"
];

function isUsefulSocialText(value = "") {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  if (text.length < 12) return false;
  if (/^(vat|tax|sales tax|discount|subtotal|total|balance|paid|installation|delivery|courier|0|none|n\/a)$/i.test(text)) return false;
  if (/^[\d\s.,£€$-]+$/.test(text)) return false;
  if (/^\d+\s*:\s*(vat|tax)$/i.test(text)) return false;
  return /[a-z]{4,}/i.test(text);
}

function matchesKnownSocialDescription(value = "") {
  const normalized = String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
  return SOCIAL_POST_DESCRIPTION_FINGERPRINTS.some((fingerprint) =>
    normalized.includes(String(fingerprint || "").replace(/\s+/g, " ").trim().toLowerCase())
  );
}

function isSocialDescriptionKey(key = "") {
  const lowerKey = String(key || "").toLowerCase();
  if (/(vat|tax|price|cost|amount|total|subtotal|balance|locator|phone|email|postcode|zip)/i.test(lowerKey)) return false;
  return (
    lowerKey.includes("customerdescription") ||
    lowerKey.includes("descriptiontext") ||
    lowerKey.includes("orderdescription") ||
    lowerKey.includes("estimatedescription") ||
    lowerKey.includes("productdescription") ||
    lowerKey.includes("lineitemdescription") ||
    lowerKey.includes("itemdescription") ||
    lowerKey.includes("description") ||
    lowerKey.includes("specification") ||
    lowerKey.includes("notes") ||
    lowerKey.includes("memo") ||
    lowerKey.includes("product")
  );
}

function normalizeSocialText(value = "") {
  return decodeXmlEntities(value)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/\u00a0/g, " ")
    .replace(/\r/g, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\n\s+/g, "\n")
    .trim();
}

function getSocialDescriptionCandidates(order = {}) {
  const fields = Array.isArray(order.debugFields) ? order.debugFields : [];
  const candidates = [];
  const pushCandidate = (source, value, key = "") => {
    const text = normalizeSocialText(value);
    if (!isUsefulSocialText(text)) return;
    if (candidates.some((candidate) => candidate.text.toLowerCase() === text.toLowerCase())) return;
    candidates.push({ source, key, text, score: scoreSocialDescription(key, text) });
  };

  pushCandidate("normalized order.description", order.description, "order.description");
  fields.forEach((field) => {
    const key = String(field.key || "");
    const value = field.value;
    if (isSocialDescriptionKey(key) || matchesKnownSocialDescription(value)) {
      pushCandidate("Corebridge field", value, key);
    }
  });

  return candidates
    .sort((left, right) => right.score - left.score)
    .slice(0, 20);
}

function scoreSocialDescription(key = "", text = "") {
  const lowerKey = String(key || "").toLowerCase();
  const lowerText = String(text || "").toLowerCase();
  let score = Math.min(String(text || "").length, 900) / 30;
  if (lowerKey.includes("customerdescription")) score += 35;
  if (lowerKey.includes("itemdescription") || lowerKey.includes("lineitemdescription")) score += 25;
  if (lowerKey.includes("description")) score += 20;
  if (lowerKey.includes("orderdestinationitem")) score += 15;
  if (lowerKey.includes("productdescription") || lowerKey.includes("estimatedescription")) score += 12;
  if (/(mm|aluminium|acrylic|vinyl|graphics|installed|illumina|led|tray|wall|floor|fascia|sign)/i.test(text)) score += 20;
  if (/(stonework|external|internal|folded|printed|raised|opal|returns|fret cut)/i.test(text)) score += 12;
  if (matchesKnownSocialDescription(text)) score += 200;
  if (lowerText.includes("vat") || lowerText.includes("tax")) score -= 55;
  if (lowerText.length < 30) score -= 15;
  return score;
}

function extractSocialOrderItems(order = {}) {
  const fields = Array.isArray(order.debugFields) ? order.debugFields : [];
  const groups = new Map();
  fields.forEach((field) => {
    const key = String(field.key || "").toLowerCase();
    const value = normalizeSocialText(field.value);
    if (!isUsefulSocialText(value)) return;
    if (/(vat|tax|price|cost|amount|total|subtotal|balance)/i.test(key)) return;
    const match = key.match(/(?:items?|orderdestinationitems?|estimateitems?|lineitems?)\.(\d+)\.(.+)$/i);
    if (!match) return;
    const group = groups.get(match[1]) || {};
    const leaf = match[2];
    if (leaf.includes("category") && !/vat|tax/i.test(value)) group.category = value;
    if ((leaf.includes("description") || leaf.includes("customerdescription")) && isUsefulSocialText(value)) group.description = value;
    if (!group.description && (leaf.includes("name") || leaf.includes("title") || leaf.includes("product")) && value.length > 22) group.description = value;
    if (leaf.includes("quantity")) group.quantity = value;
    groups.set(match[1], group);
  });

  const items = [...groups.values()]
    .filter((item) => isUsefulSocialText(item.description) || (isUsefulSocialText(item.category) && !/vat|tax/i.test(item.category)))
    .map((item) => ({
      ...item,
      category: /vat|tax/i.test(item.category || "") ? "" : item.category,
      description: item.description || item.category || ""
    }))
    .slice(0, 12);
  if (items.length) return items;
  const candidates = getSocialDescriptionCandidates(order);
  return candidates.slice(0, 3).map((candidate) => ({ category: "", description: candidate.text, quantity: "" }));
}

function getSocialLineItemBriefs(items = []) {
  return items.map((item, index) => {
    const description = cleanSocialPostSpec(item.description || "");
    const category = cleanSocialPostSpec(item.category || "");
    const combined = `${category} ${description}`.toLowerCase();
    let focusScore = Math.min(description.length / 80, 6);
    if (/(fascia|sign tray|illuminat|led|wall|floor|window|wrap|plinth|prize wheel|display|graphics|acrylic|aluminium)/i.test(combined)) focusScore += 3;
    if (/(delivery|install|installation|attend site|survey|artwork|setup labour|bought in)/i.test(combined)) focusScore -= 2;
    return {
      itemNumber: index + 1,
      category: item.category || "",
      quantity: item.quantity || "",
      description,
      suggestedTreatment: focusScore >= 4 ? "feature or mention separately" : "skim, group lightly or omit if it makes the post too long"
    };
  });
}

function buildSocialPostBrief(order, voice) {
  const descriptionCandidates = getSocialDescriptionCandidates(order);
  const items = extractSocialOrderItems(order);
  const lineItems = getSocialLineItemBriefs(items);
  const mainDescription = descriptionCandidates[0]?.text || order.description || items[0]?.description || "";
  const jobEvidence = getSocialJobEvidence(order, items, descriptionCandidates);
  const customerClassification = classifySocialCustomer(order.customerName);
  const sourceFields = (Array.isArray(order.debugFields) ? order.debugFields : [])
    .filter((field) => isSocialDescriptionKey(field.key) || /(?:items?|orderdestinationitems?|estimateitems?|lineitems?)\.\d+/i.test(String(field.key || "")))
    .map((field) => ({ key: field.key, value: normalizeSocialText(field.value).slice(0, 1000) }))
    .filter((field) => field.value)
    .slice(0, 80);
  const exactDescriptionMatches = (Array.isArray(order.debugFields) ? order.debugFields : [])
    .filter((field) => matchesKnownSocialDescription(field.value))
    .map((field) => ({ key: field.key, value: normalizeSocialText(field.value).slice(0, 1500) }));
  return {
    orderReference: order.orderReference || "",
    customerName: order.customerName || "",
    customerClassification,
    primaryDescription: mainDescription,
    description: mainDescription,
    descriptionCandidates,
    ...jobEvidence,
    address: order.address || "",
    contact: order.contact || "",
    items,
    lineItemCount: items.length,
    lineItems,
    toneName: voice?.name || "LinkedIn",
    toneSummary: getSocialPostToneSummary(voice),
    toneTraits: getSocialPostToneTraits(voice),
    transformationExamples: getSocialPostTransformationExamples(voice),
    debug: {
      chosenDescription: mainDescription,
      customerClassification,
      descriptionCandidates,
      itemCandidates: items,
      lineItems,
      exactDescriptionMatches,
      jobEvidence,
      sourceFields,
      toneName: voice?.name || "Matt Rutlidge",
      toneExcerpt: getSocialPostToneSummary(voice).slice(0, 3000),
      transformationExamples: getSocialPostTransformationExamples(voice).slice(0, 8),
      sourceFieldCount: Array.isArray(order.debugFields) ? order.debugFields.length : 0
    }
  };
}

function getSocialLookupReferences(reference = "") {
  const trimmed = String(reference || "").trim();
  const references = [trimmed];
  if (/^inv-/i.test(trimmed)) references.push(trimmed.replace(/^inv-/i, "ORD-"));
  if (/^ord-/i.test(trimmed)) references.push(trimmed.replace(/^ord-/i, "INV-"));
  return [...new Set(references.filter(Boolean))];
}

function cleanSocialPostSpec(value = "") {
  return normalizeSocialText(value)
    .replace(/\b\d{2,5}\s*mm\b/gi, "")
    .replace(/\b\d+(?:\.\d+)?\s*(?:mm|m)\b/gi, "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/\b\d+mm\b/gi, "")
    .replace(/\b3mm\b/gi, "")
    .replace(/\b5mm\b/gi, "")
    .replace(/\s*[xX]\s*/g, " ")
    .replace(/\s+/g, " ")
    .replace(/\s+([,.])/g, "$1")
    .trim();
}

function getSocialPostFeaturePhrases(brief) {
  const text = [brief.primaryDescription, brief.description, ...(brief.items || []).map((item) => item.description)]
    .filter(Boolean)
    .map(cleanSocialPostSpec)
    .join(" ")
    .toLowerCase();
  const features = [];
  if (/poster kit|poster holder|cable poster/.test(text)) {
    features.push("clean cable poster display kits");
  }
  if (/fascia|sign tray|tray|fret cut|push through|opal|illuminat|led/.test(text)) {
    features.push("a digitally printed, stencil-cut illuminated fascia sign");
  }
  if (/window|birch|plywood|3d|dimensional|layer/.test(text)) {
    features.push("layered window-box graphics with dimensional details");
  }
  if (/wall|floor|wrap|ceiling/.test(text)) {
    features.push("wall, ceiling or interior graphic wraps");
  }
  if (brief.hasInstallEvidence && /install|site|stonework|external/.test(text)) {
    features.push("careful external installation on site");
  }
  return [...new Set(features)].slice(0, 4);
}

function getSocialJobEvidence(order = {}, items = [], descriptionCandidates = []) {
  const combined = [
    order.description,
    ...(items || []).map((item) => item.description),
    ...(descriptionCandidates || []).map((candidate) => candidate.text)
  ].filter(Boolean).join(" ").toLowerCase();
  const hasInstallEvidence = /\b(install|installed|installation|fit|fitted|attend site|on site|site visit)\b/i.test(combined);
  const hasDeliveryOnlyEvidence = /\b(delivery|deliver|courier|collection|supplied only|supply only)\b/i.test(combined) && !hasInstallEvidence;
  return {
    hasInstallEvidence,
    hasDeliveryOnlyEvidence,
    evidenceSummary: hasInstallEvidence
      ? "Corebridge includes installation or site attendance wording."
      : hasDeliveryOnlyEvidence
        ? "Corebridge suggests this is supply/delivery rather than installation."
        : "Corebridge does not clearly say installation happened."
  };
}

function sanitizeGeneratedSocialPost(post = "") {
  let text = normalizeSocialText(post)
    .replace(/\bexciting news\b[!,.:\s]*/i, "")
    .replace(/\bpizzazz\b/gi, "impact")
    .replace(/\bjazz(?:ed)?\s+up\b/gi, "made stand out")
    .replace(/\btalk about a glow up!?/gi, "it made a proper difference")
    .replace(/[‐‑‒–—―-]/g, " ")
    .replace(/[ \t]{2,}/g, " ")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
  if (!text) return text;
  if (/^we'?re thrilled\b/i.test(text)) {
    text = text.replace(/^we'?re thrilled[^.\n]*[.\n]*/i, "Want to get noticed?\n\n");
  }
  return text
    .split(/\n{2,}|\n(?=\S)/)
    .map((paragraph) => paragraph.replace(/\s+/g, " ").trim())
    .filter(Boolean)
    .join("\n\n");
}

function generateFallbackSocialPost(brief) {
  const customer = brief.customerName || "a client";
  const features = getSocialPostFeaturePhrases(brief);
  const featureSentence = features.length
    ? `For ${customer}, we helped with ${features.join(", ")}.`
    : `For ${customer}, we helped turn a detailed signage brief into something bold, polished and built to be noticed.`;
  return [
    "Want to get noticed? Start with signage people actually stop and look at.",
    "",
    "In a world full of safe choices, the brands that stand out are usually the ones brave enough to show a bit of personality.",
    "",
    featureSentence,
    "",
    "The Corebridge spec might be full of materials, fixings and production detail, but the end result is much simpler: impact, visibility and a finish that feels properly considered.",
    "",
    "At Signs Express (Central Lancashire) & Signs Express (Southport), we bring bold ideas to life.",
    "",
    "Ready to be the business people talk about? Let's make it happen.",
    "",
    "#SignsExpress #Signage #Branding #VehicleGraphics #BusinessBranding"
  ].join("\n");
}

async function generateSocialPostWithAi(brief) {
  const aiStatus = getSocialPostAiStatus();
  const apiKey = String(process.env.OPENAI_API_KEY || "").trim();
  if (!apiKey) {
    return {
      post: sanitizeGeneratedSocialPost(generateFallbackSocialPost(brief)),
      source: "template",
      warning: "OPENAI_API_KEY is not configured, so this used the local fallback rather than reading the tone file with AI.",
      aiStatus
    };
  }

  const model = aiStatus.model;
  const prompt = [
    "Write one LinkedIn post for Signs Express Central Lancashire in Matt Rutlidge's style.",
    "",
    "IMPORTANT TRANSFORMATION:",
    "- The Corebridge data is raw production language. Do not repeat it as a specification list.",
    "- Read primaryDescription, descriptionCandidates, items and lineItems to understand what was supplied, produced or installed.",
    "- Treat each entry in lineItems as a separate quotation line unless the wording explicitly says one item is physically part of another.",
    "- Do not blend two separate line items into one invented product. For example, if one line is a prize wheel and another line is a plinth, describe them as separate pieces of the same project, not as a prize wheel wrapped around a plinth.",
    "- Decide what deserves space in the post. Feature visually interesting or higher-value lines, lightly mention supporting pieces, and quietly ignore boring low-value/admin/delivery lines if the post would become too long.",
    "- The tone file may contain spreadsheet rows where column A is a Corebridge job reference and column B is Matt's finished LinkedIn post for that exact job.",
    "- If toneTraits is present, treat it as the explicit personality and writing rules for the chosen person. Use it alongside the examples, not instead of them.",
    "- Read transformationExamples as paired before/after training examples. For each row, mentally ask: what did Matt take from the Corebridge job, what did he ignore, what did he simplify, what did he fluff up, and what hook style did he use?",
    "- Find repeatable patterns across all transformationExamples, then apply those patterns to the new Corebridge job.",
    "- Read toneSummary as additional historic LinkedIn examples. Infer the structure, rhythm, emoji use, hooks, line breaks, calls to action and level of technical simplification.",
    "- Before writing, internally build a style map from toneSummary and transformationExamples: hook formulas, joke patterns, self-deprecating lines, playful misdirection, favourite emojis, emoji count, emoji placement, sentence length, paragraph length and sign-off habits.",
    "- Move much closer to Matt's actual mannerisms than generic marketing copy. Matt often opens with a funny or slightly self-deprecating observation, a rhetorical question, or a playful comparison before revealing the job.",
    "- Example hook flavours to emulate when appropriate: Did I spend hours with a pencil case full of Sharpies and half-chewed crayons, or is this actually one-piece wallpaper? / Simple, effective, and looks tidy. Enough about me though, what do you think of this completed job?",
    "- Use the examples to infer which emojis Matt likes and where they normally sit. Match the usual emoji density from the tone file rather than adding random emojis.",
    "- Do not over-polish the humour. It should feel like Matt's version of funny: dry, playful, slightly cheeky, human and not corporate.",
    "- Avoid words and phrases Matt would not use, including pizzazz, jazzed up, glow up, game changer, elevate your brand, stunning solution and proud to announce.",
    "- Check customerClassification before describing the customer. Do not call a council, leisure trust, school, charity, NHS or public sector organisation a business or company. Use organisation, council, venue, team, site or facility instead.",
    "- Learn hook style from the examples. Do not start with generic AI phrases like Exciting news, We are thrilled, We are delighted, In today's fast-paced world, or Transform your space.",
    "- Turn overcomplicated production wording into a natural post about impact, branding, visibility, design and the finished result.",
    "- Use a hook, short paragraphs, light emoji use where it fits, and a conversational call to action.",
    "- Mention the customer/project if available. Mention designer credit if the source data or tone examples clearly support it, otherwise do not invent names.",
    "- It is fine to mention simplified product phrases such as illuminated fascia sign, window graphics, wall wraps or layered displays.",
    "- Only say installed, installation or on site if hasInstallEvidence is true. If hasDeliveryOnlyEvidence is true, treat it as supplied or delivered.",
    "- Do not include raw dimensions, material thicknesses, internal installation caveats, VAT/tax/price/admin lines, HTML entities, or a list of every Corebridge line item.",
    "- Do not use hyphen or dash characters anywhere in the post.",
    "- Put a blank line between paragraphs so it is easy to read on LinkedIn.",
    "- Keep it LinkedIn-ready and flowing naturally, not like a quote summary.",
    "- Include 3 to 6 relevant hashtags at the end.",
    "",
    JSON.stringify(brief, null, 2)
  ].join("\n");

  try {
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model,
        messages: [
          { role: "system", content: "You are a senior LinkedIn copywriter for a UK signage company. You convert technical signage job descriptions into warm, bold, human social posts in the user's proven tone of voice." },
          { role: "user", content: prompt }
        ],
        temperature: 0.82
      })
    });
    const payload = await response.json();
    const post = payload?.choices?.[0]?.message?.content?.trim();
    if (!response.ok || !post) throw new Error(payload?.error?.message || "AI generation failed.");
    return { post: sanitizeGeneratedSocialPost(post), source: "ai", aiStatus };
  } catch (error) {
    return { post: sanitizeGeneratedSocialPost(generateFallbackSocialPost(brief)), source: "template", warning: error.message, aiStatus };
  }
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
  app.use(express.json({ limit: "15mb" }));
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
    if (!canAccessBoard(sessionUser) && !canAccessInstaller(sessionUser) && !canAccessHolidays(sessionUser) && !canAccessAttendance(sessionUser) && !canAccessMileage(sessionUser) && !canAccessVanEstimator(sessionUser) && !canAccessRams(sessionUser) && !canAccessSocialPost(sessionUser)) {
      response.status(403).json({ error: "That account does not have access." });
      return;
    }
    const { sessionId, expiresAt } = createSession(sessionUser);
    response.setHeader("Set-Cookie", serializeSessionCookie(sessionId, { expiresAt }));
    response.json({ user: { ...sessionUser, canManagePermissions: canManagePermissions(sessionUser) } });
  });

  app.post("/api/auth/users", async (request, response) => {
    const session = getSessionFromRequest(request);
    if (!session) {
      response.status(401).json({ error: "Login required." });
      return;
    }

    request.user = session.user;
    if (!requirePermissionsManager(request, response)) return;

    try {
      const user = await createUser({
        displayName: request.body?.displayName,
        role: request.body?.role,
        password: request.body?.password
      });

      const store = await readStore();
      const holidayPerson = getHolidayStaffPerson(user.displayName);
      const yearStart = getCurrentHolidayYearStart();
      const allowanceExists = (store.holidayAllowances || []).some(
        (entry) =>
          Number(entry.yearStart || 0) === yearStart &&
          String(entry.person || "").trim().toLowerCase() === String(holidayPerson || "").trim().toLowerCase()
      );

      if (!allowanceExists && holidayPerson) {
        store.holidayAllowances = Array.isArray(store.holidayAllowances) ? store.holidayAllowances : [];
        store.holidayAllowances.unshift(
          sanitizeHolidayAllowance({
            yearStart,
            person: holidayPerson
          })
        );
        await writeStore(store);
      }

      response.json({ user });
    } catch (error) {
      response.status(400).json({ error: error.message || "Could not create user." });
    }
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
          holidays: request.body?.holidays,
          attendance: request.body?.attendance,
          mileage: request.body?.mileage,
          vanEstimator: request.body?.vanEstimator,
          rams: request.body?.rams,
          socialPost: request.body?.socialPost
        });
      response.json({ user: updatedUser });
    } catch (error) {
      response.status(400).json({ error: error.message || "Could not update permissions." });
    }
  });

  app.patch("/api/auth/users/:id/attendance-profile", async (request, response) => {
    const session = getSessionFromRequest(request);
    if (!session) {
      response.status(401).json({ error: "Login required." });
      return;
    }

    request.user = session.user;
    if (!canManagePermissions(request.user)) {
      response.status(403).json({ error: "Only Matt Rutlidge can change attendance settings." });
      return;
    }

    try {
      const user = await updateUserAttendanceProfile(request.params.id, request.body || {});
      response.json({ user });
    } catch (error) {
      console.error(error);
      response.status(400).json({ error: error.message || "Could not update attendance settings." });
    }
  });

  app.patch("/api/auth/users/:id/profile", async (request, response) => {
    const session = getSessionFromRequest(request);
    if (!session) {
      response.status(401).json({ error: "Login required." });
      return;
    }

    request.user = session.user;
    if (!requirePermissionsManager(request, response)) return;

    try {
      const user = await updateUserProfile(request.params.id, request.body || {});
      response.json({ user });
    } catch (error) {
      console.error(error);
      response.status(400).json({ error: error.message || "Could not update user profile." });
    }
  });

  app.patch("/api/auth/users/:id/password", async (request, response) => {
    const session = getSessionFromRequest(request);
    if (!session) {
      response.status(401).json({ error: "Login required." });
      return;
    }

    request.user = session.user;
    if (!requirePermissionsManager(request, response)) return;

    try {
      const user = await setUserPasswordById(request.params.id, request.body?.password);
      response.json({ user });
    } catch (error) {
      response.status(400).json({ error: error.message || "Could not update password." });
    }
  });

    app.delete("/api/auth/users/:id", async (request, response) => {
      const session = getSessionFromRequest(request);
      if (!session) {
        response.status(401).json({ error: "Login required." });
        return;
    }

      request.user = session.user;
      if (!requirePermissionsManager(request, response)) return;

      try {
        const deletedUser = await deleteUser(request.params.id);
        const store = await readStore();
        const deletedDisplayName = String(deletedUser?.displayName || "").trim();
        const deletedPersonKey = getHolidayStaffIdentityKey(getHolidayStaffPerson(deletedDisplayName) || deletedDisplayName);

        store.holidays = (store.holidays || []).filter((entry) => {
          if (getHolidayType(entry) === "birthday") return true;
          return getHolidayStaffIdentityKey(entry.person) !== deletedPersonKey;
        });
        store.holidayRequests = (store.holidayRequests || []).filter((entry) => {
          const samePerson = getHolidayStaffIdentityKey(entry.person) === deletedPersonKey;
          const sameRequester = String(entry.requestedByUserId || "") === String(deletedUser.id || "");
          return !samePerson && !sameRequester;
        });
          store.holidayAllowances = (store.holidayAllowances || []).filter(
            (entry) => getHolidayStaffIdentityKey(entry.person) !== deletedPersonKey
          );
          store.attendanceEntries = (store.attendanceEntries || []).filter(
            (entry) => getHolidayStaffIdentityKey(entry.person) !== deletedPersonKey
          );
          store.mileageClaims = (store.mileageClaims || []).filter(
            (entry) => String(entry.userId || "") !== String(deletedUser.id || "")
          );
          store.notifications = (store.notifications || []).filter(
            (entry) => String(entry.userId || "") !== String(deletedUser.id || "")
          );

        await writeStore(store);
        response.json({ user: deletedUser });
      } catch (error) {
        response.status(400).json({ error: error.message || "Could not delete user." });
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
    if (!requireBoardOrRamsAccess(request, response)) return;
    const store = await readStore();
    response.json(toPublicJobs(store.jobs));
  });

  app.get("/api/rams/hospitals", async (request, response) => {
    if (!requireRamsAccess(request, response)) return;
    const address = String(request.query.address || "").trim();
    if (!address) {
      response.status(400).json({ error: "Installation address is required." });
      return;
    }
    try {
      response.json(await findNearestHospitals(address));
    } catch (error) {
      console.error("RAMS hospital lookup failed.", error);
      response.status(500).json({ error: "Could not load nearby hospitals." });
    }
  });

  app.get("/api/rams/profiles", async (request, response) => {
    if (!requireBoardOrRamsAccess(request, response)) return;
    const usersStore = await readUsersStore();
    response.json((usersStore.users || []).map(toPublicRamsProfile));
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

  app.get("/api/social-post/voices", async (request, response) => {
    if (!requireSocialPostAccess(request, response)) return;
    const store = await readStore();
    response.json({
      voices: getSocialPostVoices(store)
    });
  });

  app.get("/api/social-post/status", async (request, response) => {
    if (!requireSocialPostAccess(request, response)) return;
    const store = await readStore();
    const voices = getSocialPostVoices(store);
    response.json({
      ai: getSocialPostAiStatus(),
      voices: voices.map((voice) => ({
        id: voice.id,
        name: voice.name,
        fileName: voice.fileName,
        contentLength: voice.content.length,
        supportingTextLength: voice.supportingText.length,
        exampleCount: voice.examples.length,
        createdAt: voice.createdAt,
        seeded: voice.seeded
      }))
    });
  });

  app.post("/api/social-post/voices", async (request, response) => {
    if (!canEditSocialPost(request.user)) {
      response.status(403).json({ error: "Only admins can upload tone of voice files." });
      return;
    }

    try {
      const voice = parseToneVoiceUpload(request.body || {});
      if (!voice.content) {
        response.status(400).json({ error: "The tone file did not contain readable text." });
        return;
      }
      const store = await readStore();
      const voices = Array.isArray(store.socialPostToneVoices) ? store.socialPostToneVoices : [];
      store.socialPostDeletedToneVoiceIds = (store.socialPostDeletedToneVoiceIds || []).filter((id) => String(id) !== String(voice.id));
      store.socialPostToneVoices = [voice, ...voices.filter((entry) => String(entry.id) !== String(voice.id))].slice(0, 20);
      const savedStore = await writeStore(store);
      response.json({ voices: getSocialPostVoices(savedStore), voice });
    } catch (error) {
      response.status(400).json({ error: error.message || "Could not upload tone of voice." });
    }
  });

  app.patch("/api/social-post/voices/:id", async (request, response) => {
    if (!canEditSocialPost(request.user)) {
      response.status(403).json({ error: "Only admins can edit tone of voice files." });
      return;
    }

    const store = await readStore();
    const voices = Array.isArray(store.socialPostToneVoices) ? store.socialPostToneVoices : [];
    const index = voices.findIndex((voice) => String(voice.id) === String(request.params.id));
    const editingDefaultVoice = String(request.params.id) === String(getDefaultSocialToneVoice().id);
    if (index === -1 && !editingDefaultVoice) {
      response.status(404).json({ error: "Tone of voice not found." });
      return;
    }

    const existing = sanitizeSocialToneVoice(index === -1 ? getDefaultSocialToneVoice() : voices[index]);
    const nextVoice = sanitizeSocialToneVoice({
      ...existing,
      name: request.body?.name ?? existing.name,
      content: request.body?.content ?? existing.content,
      supportingText: request.body?.supportingText ?? existing.supportingText,
      examples: request.body?.examples ?? existing.examples,
      fileName: request.body?.fileName ?? existing.fileName,
      createdAt: existing.createdAt,
      seeded: false
    });
    if (!nextVoice.content) {
      response.status(400).json({ error: "Tone content cannot be empty." });
      return;
    }

    store.socialPostToneVoices = index === -1
      ? [nextVoice, ...voices].slice(0, 20)
      : voices.map((voice, voiceIndex) => (voiceIndex === index ? nextVoice : voice));
    const savedStore = await writeStore(store);
    response.json({ voices: getSocialPostVoices(savedStore), voice: nextVoice });
  });

  app.delete("/api/social-post/voices/:id", async (request, response) => {
    if (!canEditSocialPost(request.user)) {
      response.status(403).json({ error: "Only admins can delete tone of voice files." });
      return;
    }
    const store = await readStore();
    const deletedId = String(request.params.id || "");
    store.socialPostToneVoices = (store.socialPostToneVoices || []).filter((voice) => String(voice.id) !== deletedId);
    store.socialPostDeletedToneVoiceIds = [
      ...new Set([...(store.socialPostDeletedToneVoiceIds || []).map(String), deletedId].filter(Boolean))
    ];
    const savedStore = await writeStore(store);
    response.json({ voices: getSocialPostVoices(savedStore) });
  });

  app.post("/api/social-post/generate", async (request, response) => {
    if (!requireSocialPostAccess(request, response)) return;
    try {
      const orderReference = String(request.body?.orderReference || "").trim();
      if (!orderReference) {
        response.status(400).json({ error: "Enter a Corebridge order reference." });
        return;
      }
      const store = await readStore();
      const voices = getSocialPostVoices(store);
      const selectedVoice = voices.find((voice) => String(voice.id) === String(request.body?.voiceId)) || voices[0] || sanitizeSocialToneVoice({
        name: "Matt Rutlidge",
        content: "Friendly, practical LinkedIn posts for completed signage work. Mention the customer, what was produced, and the installation or finish where relevant."
      });
      const lookupReferences = getSocialLookupReferences(orderReference);
      let lookup = null;
      let order = null;
      const lookupAttempts = [];
      for (const reference of lookupReferences) {
        try {
          const candidateLookup = await fetchCoreBridgeOrders(reference, true, { includeClosed: true });
          const candidateOrder = (candidateLookup.orders || [])[0];
          lookupAttempts.push({ reference, found: Boolean(candidateOrder), sourceUrl: candidateLookup.sourceUrl || "" });
          if (candidateOrder) {
            lookup = candidateLookup;
            order = candidateOrder;
            break;
          }
        } catch (lookupError) {
          lookupAttempts.push({ reference, found: false, error: lookupError.message || "Lookup failed" });
        }
      }
      if (!order) {
        response.status(404).json({ error: "No Corebridge order found for that reference." });
        return;
      }
      const brief = buildSocialPostBrief(order, selectedVoice);
      brief.lookupAttempts = lookupAttempts;
      brief.debug.lookupAttempts = lookupAttempts;
      brief.debug.selectedVoice = {
        id: selectedVoice.id,
        name: selectedVoice.name,
        fileName: selectedVoice.fileName,
        contentLength: selectedVoice.content.length,
        supportingTextLength: selectedVoice.supportingText.length,
        exampleCount: selectedVoice.examples.length,
        isFallback: !voices.length
      };
      const generated = await generateSocialPostWithAi(brief);
      response.json({
        order,
        lookup,
        brief,
        voice: selectedVoice,
        post: generated.post,
        source: generated.source,
        warning: generated.warning || "",
        ai: generated.aiStatus || getSocialPostAiStatus(),
        debug: brief.debug
      });
    } catch (error) {
      response.status(error.statusCode || 500).json({
        error: error.statusCode === 503 ? "Corebridge is not configured yet." : "Could not generate the social post.",
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
    const existingJob = index >= 0 ? sanitizeJob(store.jobs[index]) : null;
    if (index >= 0) {
      nextJob.createdAt = store.jobs[index].createdAt || nextJob.createdAt;
      store.jobs[index] = nextJob;
    } else {
      store.jobs.unshift(nextJob);
    }

    const usersStore = await readUsersStore();
    if (!existingJob && nextJob.date) {
      const jobSummary = getJobNotificationSummary(nextJob);
      const bookedDate = formatBoardNotificationDate(nextJob.date);
        pushBoardNotification(store, usersStore.users || [], () => ({
          jobId: nextJob.id,
          type: "job-added",
          title: "Job added to board",
          message: `${jobSummary} was added to the installation board for ${bookedDate} by ${request.user?.displayName || "a user"}.`
        }));
    } else if (existingJob && existingJob.date !== nextJob.date && nextJob.date) {
      const jobSummary = getJobNotificationSummary(nextJob);
      const fromLabel = existingJob.date ? formatBoardNotificationDate(existingJob.date) : "Unscheduled";
      const toLabel = formatBoardNotificationDate(nextJob.date);
        pushBoardNotification(store, usersStore.users || [], () => ({
          jobId: nextJob.id,
          type: "job-moved",
          title: existingJob.date ? "Job moved on board" : "Job added to board",
          message: existingJob.date
            ? `${jobSummary} moved from ${fromLabel} to ${toLabel} by ${request.user?.displayName || "a user"}.`
          : `${jobSummary} was added to the installation board for ${toLabel} by ${request.user?.displayName || "a user"}.`
      }));
    }

    const savedStore = await writeStore(store);
    const payload = {
      jobs: toPublicJobs(savedStore.jobs),
      holidays: savedStore.holidays,
      board: buildBoardRowsFromStore(savedStore),
      job: toPublicJob(nextJob)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.post("/api/jobs/:id/rams", async (request, response) => {
    if (!requireRamsAccess(request, response)) return;
    const store = await readStore();
    const index = store.jobs.findIndex((job) => String(job.id || "") === String(request.params.id || ""));
    if (index === -1) {
      response.status(404).json({ error: "Job not found." });
      return;
    }

    const existing = sanitizeJob(store.jobs[index]);
    const incoming = sanitizeRamsDocument({
      ...(request.body || {}),
      createdByName: request.body?.createdByName || request.user?.displayName || ""
    });
    const existingRams = existing.ramsDocuments.find((entry) => String(entry.id || "") === String(incoming.id || ""));
    const nextDocument = sanitizeRamsDocument({
      ...incoming,
      viewPayload: request.body?.viewPayload || request.body?.pdf || incoming.viewPayload || {},
      createdAt: existingRams?.createdAt || incoming.createdAt || new Date().toISOString(),
      createdByName: existingRams?.createdByName || incoming.createdByName || request.user?.displayName || "",
      updatedAt: new Date().toISOString(),
      pdfStorageName: "",
      pdfFileName: "",
      pdfSize: 0
    });
    const ramsDocuments = existing.ramsDocuments.some((entry) => String(entry.id || "") === String(nextDocument.id || ""))
      ? existing.ramsDocuments.map((entry) => String(entry.id || "") === String(nextDocument.id || "") ? nextDocument : entry)
      : [nextDocument, ...existing.ramsDocuments];
    const nextJob = sanitizeJob({
      ...existing,
      ramsDocuments
    });

    store.jobs[index] = nextJob;
    const savedStore = await writeStore(store);
    const payload = {
      jobs: toPublicJobs(savedStore.jobs),
      holidays: savedStore.holidays,
      board: buildBoardRowsFromStore(savedStore),
      job: toPublicJob(nextJob),
      ramsDocument: nextDocument
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.post("/api/jobs/:jobId/rams/:ramsId/amendments", async (request, response) => {
    if (!requireBoardAccess(request, response)) return;
    const note = String(request.body?.note || "").trim();
    if (!note) {
      response.status(400).json({ error: "Add a note explaining the amendment required." });
      return;
    }

    const store = await readStore();
    const job = store.jobs.find((entry) => String(entry.id || "") === String(request.params.jobId || ""));
    const document = Array.isArray(job?.ramsDocuments)
      ? job.ramsDocuments.map(sanitizeRamsDocument).find((entry) => String(entry.id || "") === String(request.params.ramsId || ""))
      : null;
    if (!job || !document) {
      response.status(404).json({ error: "RAMS document not found." });
      return;
    }

    const usersStore = await readUsersStore();
    const creatorName = String(document.createdByName || "").trim().toLowerCase();
    const creator = (usersStore.users || []).find((user) => String(user.displayName || "").trim().toLowerCase() === creatorName);
    const recipients = creator
      ? [creator]
      : getBoardEditorNotificationRecipients(usersStore.users || []);
    const jobSummary = getJobNotificationSummary(job);
    store.notifications = Array.isArray(store.notifications) ? store.notifications : [];
    recipients.forEach((user) => {
      store.notifications.unshift(
        createNotification({
          userId: user.id,
          type: "rams-amendment",
          title: "RAMS amendment requested",
          message: `${request.user?.displayName || "A user"} requested a RAMS amendment for ${jobSummary}: ${note}`,
          link: `/rams?jobId=${encodeURIComponent(job.id)}&ramsId=${encodeURIComponent(document.id)}`
        })
      );
    });

    const savedStore = await writeStore(store);
    response.json({
      ok: true,
      notifications: savedStore.notifications.filter((entry) => String(entry.userId || "") === String(request.user?.id || ""))
    });
  });

  app.get("/api/jobs/:jobId/rams/:ramsId/pdf", async (request, response) => {
    if (!requireBoardAccess(request, response)) return;
    const store = await readStore();
    const job = store.jobs.find((entry) => String(entry.id || "") === String(request.params.jobId || ""));
    const document = Array.isArray(job?.ramsDocuments)
      ? job.ramsDocuments.map(sanitizeRamsDocument).find((entry) => String(entry.id || "") === String(request.params.ramsId || ""))
      : null;
    if (!job || !document?.pdfStorageName) {
      response.status(404).json({ error: "RAMS PDF not found." });
      return;
    }
    const filePath = path.join(getRamsUploadsDir(), document.pdfStorageName);
    if (!fs.existsSync(filePath)) {
      response.status(404).json({ error: "RAMS PDF file not found." });
      return;
    }
    response.setHeader("Content-Type", "application/pdf");
    response.setHeader("Content-Disposition", `inline; filename="${(document.pdfFileName || "RAMS.pdf").replace(/"/g, "")}"`);
    fs.createReadStream(filePath).pipe(response);
  });

  app.delete("/api/jobs/:jobId/rams/:ramsId", async (request, response) => {
    if (!requireRamsAccess(request, response)) return;
    const store = await readStore();
    const index = store.jobs.findIndex((entry) => String(entry.id || "") === String(request.params.jobId || ""));
    if (index === -1) {
      response.status(404).json({ error: "Job not found." });
      return;
    }
    const existing = sanitizeJob(store.jobs[index]);
    const document = existing.ramsDocuments.find((entry) => String(entry.id || "") === String(request.params.ramsId || ""));
    if (!document) {
      response.status(404).json({ error: "RAMS document not found." });
      return;
    }
    if (document.pdfStorageName) {
      try {
        await fsp.unlink(path.join(getRamsUploadsDir(), document.pdfStorageName));
      } catch (error) {
        if (error?.code !== "ENOENT") throw error;
      }
    }
    const nextJob = sanitizeJob({
      ...existing,
      ramsDocuments: existing.ramsDocuments.filter((entry) => String(entry.id || "") !== String(request.params.ramsId || ""))
    });
    store.jobs[index] = nextJob;
    const savedStore = await writeStore(store);
    const payload = {
      jobs: toPublicJobs(savedStore.jobs),
      holidays: savedStore.holidays,
      board: buildBoardRowsFromStore(savedStore),
      job: toPublicJob(nextJob)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.post("/api/jobs/:id/complete", async (request, response) => {
    if (!requireBoardAccess(request, response)) return;
    const store = await readStore();
    const index = store.jobs.findIndex((job) => String(job.id || "") === String(request.params.id || ""));
    if (index === -1) {
      response.status(404).json({ error: "Job not found." });
      return;
    }

    const existing = sanitizeJob(store.jobs[index]);
    const nextJob = sanitizeJob({
      ...existing,
      isCompleted: true,
      isSnagging: false,
      completedAt: existing.completedAt || new Date().toISOString(),
      completedByUserId: request.user?.id || existing.completedByUserId,
      completedByName: request.user?.displayName || existing.completedByName,
      snaggingAt: "",
      snaggingByUserId: "",
      snaggingByName: ""
    });

    store.jobs[index] = nextJob;
    if (!existing.isCompleted) {
      const usersStore = await readUsersStore();
      const jobSummary = getJobNotificationSummary(nextJob);
      const photoCount = Array.isArray(nextJob.photos) ? nextJob.photos.length : 0;
      const photoSummary =
        photoCount === 0
          ? "No photos were uploaded."
          : photoCount === 1
            ? "1 photo was uploaded."
            : `${photoCount} photos were uploaded.`;
        pushBoardNotification(store, usersStore.users || [], () => ({
          jobId: nextJob.id,
          type: "job-completed",
          title: "Job marked complete",
          message: `${jobSummary} was marked complete by ${request.user?.displayName || "a user"}. ${photoSummary}`
        }));
    }
    const savedStore = await writeStore(store);
    const payload = {
      jobs: toPublicJobs(savedStore.jobs),
      holidays: savedStore.holidays,
      board: buildBoardRowsFromStore(savedStore),
      job: toPublicJob(nextJob)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.post("/api/jobs/:id/snagging", async (request, response) => {
    if (!requireBoardAdmin(request, response)) return;
    const store = await readStore();
    const index = store.jobs.findIndex((job) => String(job.id || "") === String(request.params.id || ""));
    if (index === -1) {
      response.status(404).json({ error: "Job not found." });
      return;
    }

    const existing = sanitizeJob(store.jobs[index]);
    const nextJob = sanitizeJob({
      ...existing,
      isSnagging: true,
      snaggingAt: existing.snaggingAt || new Date().toISOString(),
      snaggingByUserId: request.user?.id || existing.snaggingByUserId,
      snaggingByName: request.user?.displayName || existing.snaggingByName
    });

    store.jobs[index] = nextJob;
    if (!existing.isSnagging) {
      const usersStore = await readUsersStore();
      const jobSummary = getJobNotificationSummary(nextJob);
      pushBoardEditorNotification(store, usersStore.users || [], () => ({
        jobId: nextJob.id,
        type: "job-snagging",
        title: "Snagging raised",
        message: `${jobSummary} was marked as snagging by ${request.user?.displayName || "a user"}.`
      }));
    }

    const savedStore = await writeStore(store);
    const payload = {
      jobs: toPublicJobs(savedStore.jobs),
      holidays: savedStore.holidays,
      board: buildBoardRowsFromStore(savedStore),
      job: toPublicJob(nextJob)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.post("/api/jobs/:id/unsnagging", async (request, response) => {
    if (!requireBoardAdmin(request, response)) return;
    const store = await readStore();
    const index = store.jobs.findIndex((job) => String(job.id || "") === String(request.params.id || ""));
    if (index === -1) {
      response.status(404).json({ error: "Job not found." });
      return;
    }

    const existing = sanitizeJob(store.jobs[index]);
    const nextJob = sanitizeJob({
      ...existing,
      isSnagging: false,
      snaggingAt: "",
      snaggingByUserId: "",
      snaggingByName: ""
    });

    store.jobs[index] = nextJob;
    const savedStore = await writeStore(store);
    const payload = {
      jobs: toPublicJobs(savedStore.jobs),
      holidays: savedStore.holidays,
      board: buildBoardRowsFromStore(savedStore),
      job: toPublicJob(nextJob)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.post("/api/jobs/:id/uncomplete", async (request, response) => {
    if (!requireBoardAccess(request, response)) return;
    const store = await readStore();
    const index = store.jobs.findIndex((job) => String(job.id || "") === String(request.params.id || ""));
    if (index === -1) {
      response.status(404).json({ error: "Job not found." });
      return;
    }

    const existing = sanitizeJob(store.jobs[index]);
    const nextJob = sanitizeJob({
      ...existing,
      isCompleted: false,
      completedAt: "",
      completedByUserId: "",
      completedByName: ""
    });

    store.jobs[index] = nextJob;
    const savedStore = await writeStore(store);
    const payload = {
      jobs: toPublicJobs(savedStore.jobs),
      holidays: savedStore.holidays,
      board: buildBoardRowsFromStore(savedStore),
      job: toPublicJob(nextJob)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.post("/api/jobs/:id/photos", async (request, response) => {
    if (!requireBoardAccess(request, response)) return;

    const store = await readStore();
    const index = store.jobs.findIndex((job) => String(job.id || "") === String(request.params.id || ""));
    if (index === -1) {
      response.status(404).json({ error: "Job not found." });
      return;
    }

    const dataUrl = String(request.body?.dataUrl || "").trim();
    const fileName = String(request.body?.fileName || "job-photo.jpg").trim() || "job-photo.jpg";
    const match = dataUrl.match(/^data:(image\/jpeg|image\/jpg);base64,(.+)$/i);
    if (!match) {
      response.status(400).json({ error: "Photos must be uploaded as JPEG images." });
      return;
    }

    let buffer;
    try {
      buffer = Buffer.from(match[2], "base64");
    } catch (error) {
      response.status(400).json({ error: "Invalid image data." });
      return;
    }

    if (!buffer.length) {
      response.status(400).json({ error: "Image data is empty." });
      return;
    }

    const photoId = makeId();
    const storageName = `${String(request.params.id || "").trim()}-${photoId}.jpg`;
    const uploadDir = await ensureJobUploadsDir();
    await fsp.writeFile(path.join(uploadDir, storageName), buffer);
    const dimensions = parseJpegDimensions(buffer);

    const existing = sanitizeJob(store.jobs[index]);
    const photos = Array.isArray(existing.photos) ? existing.photos.map(sanitizeJobPhoto) : [];
    photos.unshift(
      sanitizeJobPhoto({
        id: photoId,
        storageName,
        fileName,
        contentType: "image/jpeg",
        size: buffer.length,
        width: Number(request.body?.width) || dimensions.width,
        height: Number(request.body?.height) || dimensions.height,
        uploadedByName: request.user?.displayName || ""
      })
    );

    const nextJob = sanitizeJob({
      ...existing,
      photos
    });

    store.jobs[index] = nextJob;
    const savedStore = await writeStore(store);
    const payload = {
      jobs: toPublicJobs(savedStore.jobs),
      holidays: savedStore.holidays,
      board: buildBoardRowsFromStore(savedStore),
      job: toPublicJob(nextJob)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.get("/api/jobs/:jobId/photos/:photoId", async (request, response) => {
    if (!requireBoardAccess(request, response)) return;
    const store = await readStore();
    const job = store.jobs.find((entry) => String(entry.id || "") === String(request.params.jobId || ""));
    const photo = Array.isArray(job?.photos)
      ? job.photos.map(sanitizeJobPhoto).find((entry) => String(entry.id || "") === String(request.params.photoId || ""))
      : null;

    if (!job || !photo?.storageName) {
      response.status(404).json({ error: "Photo not found." });
      return;
    }

    const filePath = path.join(getJobUploadsDir(), photo.storageName);
    if (!fs.existsSync(filePath)) {
      response.status(404).json({ error: "Photo file not found." });
      return;
    }

    response.setHeader("Content-Type", photo.contentType || "image/jpeg");
    response.sendFile(filePath);
  });

  app.delete("/api/jobs/:jobId/photos/:photoId", async (request, response) => {
    if (!requireBoardAccess(request, response)) return;
    const store = await readStore();
    const index = store.jobs.findIndex((entry) => String(entry.id || "") === String(request.params.jobId || ""));
    if (index === -1) {
      response.status(404).json({ error: "Job not found." });
      return;
    }

    const existing = sanitizeJob(store.jobs[index]);
    const photo = existing.photos.find((entry) => String(entry.id || "") === String(request.params.photoId || ""));
    if (!photo) {
      response.status(404).json({ error: "Photo not found." });
      return;
    }

    try {
      if (photo.storageName) {
        await fsp.unlink(path.join(getJobUploadsDir(), photo.storageName));
      }
    } catch (error) {
      if (error?.code !== "ENOENT") {
        throw error;
      }
    }

    const nextJob = sanitizeJob({
      ...existing,
      photos: existing.photos.filter((entry) => String(entry.id || "") !== String(request.params.photoId || ""))
    });
    store.jobs[index] = nextJob;
    const savedStore = await writeStore(store);
    const payload = {
      jobs: toPublicJobs(savedStore.jobs),
      holidays: savedStore.holidays,
      board: buildBoardRowsFromStore(savedStore),
      job: toPublicJob(nextJob)
    };
    broadcast("board-updated", payload.board);
    response.json(payload);
  });

  app.get("/api/jobs/:id/export", async (request, response) => {
    if (!requireBoardAccess(request, response)) return;
    const store = await readStore();
    const job = store.jobs.find((entry) => String(entry.id || "") === String(request.params.id || ""));
    if (!job) {
      response.status(404).json({ error: "Job not found." });
      return;
    }

    const normalizedJob = sanitizeJob(job);
    const photoAssets = [];
    for (const photo of normalizedJob.photos) {
      const filePath = path.join(getJobUploadsDir(), photo.storageName);
      if (!photo.storageName || !fs.existsSync(filePath)) continue;
      const buffer = await fsp.readFile(filePath);
      const dimensions = parseJpegDimensions(buffer);
      photoAssets.push({
        ...photo,
        buffer,
        width: photo.width || dimensions.width || 1,
        height: photo.height || dimensions.height || 1
      });
    }

    const pdfBuffer = buildPdfDocument(normalizedJob, photoAssets);
    const safeName = String(normalizedJob.orderReference || normalizedJob.customerName || "job-export")
      .replace(/[^a-z0-9_-]+/gi, "-")
      .replace(/^-+|-+$/g, "") || "job-export";
    response.setHeader("Content-Type", "application/pdf");
    response.setHeader("Content-Disposition", `attachment; filename=\"${safeName}.pdf\"`);
    response.send(pdfBuffer);
  });

  app.delete("/api/jobs/:id", async (request, response) => {
    if (!requireBoardAdmin(request, response)) return;
    const store = await readStore();
    const removedJob = store.jobs.find((job) => String(job.id || "") === String(request.params.id || ""));
    store.jobs = store.jobs.filter((job) => job.id !== request.params.id);
    if (removedJob) {
      await deleteJobPhotoFiles(removedJob);
      await deleteRamsPdfFiles(removedJob);
    }
    const savedStore = await writeStore(store);
    const payload = {
      jobs: toPublicJobs(savedStore.jobs),
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

  app.get("/api/notifications", async (request, response) => {
    const store = await readStore();
    const userId = String(request.user?.id || "");
    const notifications = (store.notifications || []).filter(
      (entry) => String(entry.userId || "") === userId
    );
    response.json(notifications);
  });

  app.patch("/api/notifications/:id/read", async (request, response) => {
    const store = await readStore();
    const userId = String(request.user?.id || "");
    const notificationIndex = (store.notifications || []).findIndex(
      (entry) =>
        String(entry.id || "") === String(request.params.id || "") &&
        String(entry.userId || "") === userId
    );

    if (notificationIndex === -1) {
      response.status(404).json({ error: "Notification not found." });
      return;
    }

    store.notifications[notificationIndex] = sanitizeNotification({
      ...store.notifications[notificationIndex],
      read: true
    });

    const savedStore = await writeStore(store);
    response.json(
      savedStore.notifications.filter((entry) => String(entry.userId || "") === userId)
    );
  });

  app.post("/api/notifications/read-all", async (request, response) => {
    const store = await readStore();
    const userId = String(request.user?.id || "");
    store.notifications = (store.notifications || []).map((entry) =>
      String(entry.userId || "") === userId
        ? sanitizeNotification({ ...entry, read: true })
        : entry
    );
    const savedStore = await writeStore(store);
    response.json(
      savedStore.notifications.filter((entry) => String(entry.userId || "") === userId)
    );
  });

  app.get("/api/mileage/admin", async (request, response) => {
    if (!requireMileageAdmin(request, response)) return;
    const requestedMonth = String(request.query.month || "").trim();
    const monthId = parseMonthId(requestedMonth) ? requestedMonth : getCurrentMonthId();
    const store = await readStore();
    const usersStore = await readUsersStore();
    response.json(buildMileageAdminOverview(store.mileageClaims || [], usersStore.users || [], monthId));
  });

  app.get("/api/mileage", async (request, response) => {
    if (!requireMileageAccess(request, response)) return;
    const store = await readStore();
    const userId = String(request.user?.id || "");
    const requestedMonth = String(request.query.month || "").trim();
    const monthId = parseMonthId(requestedMonth) ? requestedMonth : getCurrentMonthId();
    const userClaims = (store.mileageClaims || [])
      .map((claim) => sanitizeMileageClaim(claim))
      .filter((claim) => String(claim.userId || "") === userId);
    const currentClaim =
      userClaims.find((claim) => String(claim.monthId || "") === monthId) ||
      sanitizeMileageClaim({
        userId,
        userName: request.user?.displayName || "",
        monthId,
        lines: []
      });

    response.json({
      monthId,
      monthLabel: formatMileageMonthLabel(monthId),
      claim: currentClaim,
      history: buildMileageHistory(userClaims)
    });
  });

  app.post("/api/mileage/estimate", async (request, response) => {
    if (!requireMileageAccess(request, response)) return;
    const from = String(request.body?.from || "").trim();
    const to = String(request.body?.to || "").trim();
    if (!from || !to) {
      response.status(400).json({ error: "From and To destinations are required." });
      return;
    }

    try {
      response.json(await estimateDrivingMiles(from, to));
    } catch (error) {
      console.error("Mileage estimate failed.", error.message);
      response.json({ miles: 0, resolved: false, message: "Could not calculate a driving route yet." });
    }
  });

  app.post("/api/mileage", async (request, response) => {
    if (!requireMileageAccess(request, response)) return;
    const requestedMonth = String(request.body?.monthId || "").trim();
    const monthId = parseMonthId(requestedMonth) ? requestedMonth : getCurrentMonthId();
    const userId = String(request.user?.id || "");
    const nextClaim = sanitizeMileageClaim({
      ...request.body,
      userId,
      userName: request.user?.displayName || "",
      monthId,
      status: "submitted"
    });

    if (!nextClaim.lines.length) {
      response.status(400).json({ error: "Add at least one mileage line before submitting." });
      return;
    }
    const invalidLine = nextClaim.lines.find((line) => !isValidIsoDate(line.date) || !line.from || !line.to || !line.note || !Number(line.miles));
    if (invalidLine) {
      response.status(400).json({ error: "Every mileage line needs Date, From, To, Miles and a note explaining what it was for." });
      return;
    }

    const store = await readStore();
    store.mileageClaims = Array.isArray(store.mileageClaims) ? store.mileageClaims : [];
    const existingIndex = store.mileageClaims.findIndex(
      (claim) => String(claim.userId || "") === userId && String(claim.monthId || "") === monthId
    );
    let savedClaim;

    if (existingIndex >= 0) {
      const existingClaim = sanitizeMileageClaim(store.mileageClaims[existingIndex]);
      savedClaim = sanitizeMileageClaim({
        ...existingClaim,
        lines: [...existingClaim.lines, ...nextClaim.lines],
        submittedAt: new Date().toISOString()
      });
      store.mileageClaims[existingIndex] = savedClaim;
    } else {
      savedClaim = nextClaim;
      store.mileageClaims.unshift(savedClaim);
    }

    const usersStore = await readUsersStore();
    const matt = (usersStore.users || []).find(
      (user) => String(user.displayName || "").trim().toLowerCase() === "matt rutlidge"
    );
    if (matt?.id) {
      store.notifications = Array.isArray(store.notifications) ? store.notifications : [];
      store.notifications.unshift(
        createNotification({
          userId: matt.id,
          type: "mileage-submitted",
          title: "Mileage submitted",
          message: `${request.user?.displayName || "A user"} submitted ${nextClaim.totalMiles} miles for ${formatMileageMonthLabel(monthId)}.`,
          link: `/mileage?month=${encodeURIComponent(monthId)}`
        })
      );
    }

    const savedStore = await writeStore(store);
    const savedClaims = (savedStore.mileageClaims || [])
      .map((claim) => sanitizeMileageClaim(claim))
      .filter((claim) => String(claim.userId || "") === userId);
    response.json({
      monthId,
      monthLabel: formatMileageMonthLabel(monthId),
      claim: sanitizeMileageClaim({
        userId,
        userName: request.user?.displayName || "",
        monthId,
        lines: []
      }),
      history: buildMileageHistory(savedClaims)
    });
  });

  app.delete("/api/mileage/:monthId/lines/:lineId", async (request, response) => {
    if (!requireMileageAccess(request, response)) return;
    const monthId = String(request.params.monthId || "").trim();
    const lineId = String(request.params.lineId || "").trim();
    if (!parseMonthId(monthId) || !lineId) {
      response.status(400).json({ error: "A valid mileage month and journey are required." });
      return;
    }

    const userId = String(request.user?.id || "");
    const store = await readStore();
    const claimIndex = (store.mileageClaims || []).findIndex(
      (claim) => String(claim.userId || "") === userId && String(claim.monthId || "") === monthId
    );

    if (claimIndex === -1) {
      response.status(404).json({ error: "Mileage submission not found." });
      return;
    }

    const existingClaim = sanitizeMileageClaim(store.mileageClaims[claimIndex]);
    const removedLine = existingClaim.lines.find((line) => String(line.id || "") === lineId);
    if (!removedLine) {
      response.status(404).json({ error: "Mileage journey not found." });
      return;
    }

    const nextLines = existingClaim.lines.filter((line) => String(line.id || "") !== lineId);
    if (nextLines.length) {
      store.mileageClaims[claimIndex] = sanitizeMileageClaim({
        ...existingClaim,
        lines: nextLines
      });
    } else {
      store.mileageClaims = (store.mileageClaims || []).filter((_, index) => index !== claimIndex);
    }

    const usersStore = await readUsersStore();
    const matt = (usersStore.users || []).find(
      (user) => String(user.displayName || "").trim().toLowerCase() === "matt rutlidge"
    );
    if (matt?.id) {
      store.notifications = Array.isArray(store.notifications) ? store.notifications : [];
      store.notifications.unshift(
        createNotification({
          userId: matt.id,
          type: "mileage-journey-deleted",
          title: "Mileage journey deleted",
          message: `${request.user?.displayName || "A user"} deleted a ${removedLine.miles} mile journey from ${formatMileageMonthLabel(monthId)} mileage.`,
          link: `/mileage?month=${encodeURIComponent(monthId)}`
        })
      );
    }

    const savedStore = await writeStore(store);
    const savedClaims = (savedStore.mileageClaims || [])
      .map((claim) => sanitizeMileageClaim(claim))
      .filter((claim) => String(claim.userId || "") === userId);
    const savedClaim =
      savedClaims.find((claim) => String(claim.monthId || "") === monthId) ||
      sanitizeMileageClaim({
        userId,
        userName: request.user?.displayName || "",
        monthId,
        lines: []
      });
    response.json({
      monthId,
      monthLabel: formatMileageMonthLabel(monthId),
      claim: savedClaim,
      history: buildMileageHistory(savedClaims)
    });
  });

  app.delete("/api/mileage/:monthId", async (request, response) => {
    if (!requireMileageAccess(request, response)) return;
    const monthId = String(request.params.monthId || "").trim();
    if (!parseMonthId(monthId)) {
      response.status(400).json({ error: "A valid mileage month is required." });
      return;
    }

    const userId = String(request.user?.id || "");
    const store = await readStore();
    const existingClaim = (store.mileageClaims || []).find(
      (claim) => String(claim.userId || "") === userId && String(claim.monthId || "") === monthId
    );
    store.mileageClaims = (store.mileageClaims || []).filter(
      (claim) => !(String(claim.userId || "") === userId && String(claim.monthId || "") === monthId)
    );

    if (existingClaim) {
      const usersStore = await readUsersStore();
      const matt = (usersStore.users || []).find(
        (user) => String(user.displayName || "").trim().toLowerCase() === "matt rutlidge"
      );
      if (matt?.id) {
        store.notifications = Array.isArray(store.notifications) ? store.notifications : [];
        store.notifications.unshift(
          createNotification({
            userId: matt.id,
            type: "mileage-deleted",
            title: "Mileage deleted",
            message: `${request.user?.displayName || "A user"} deleted their ${formatMileageMonthLabel(monthId)} mileage submission.`,
            link: `/mileage?month=${encodeURIComponent(monthId)}`
          })
        );
      }
    }

    const savedStore = await writeStore(store);
    const savedClaims = (savedStore.mileageClaims || [])
      .map((claim) => sanitizeMileageClaim(claim))
      .filter((claim) => String(claim.userId || "") === userId);
    response.json({
      monthId,
      monthLabel: formatMileageMonthLabel(monthId),
      claim: sanitizeMileageClaim({
        userId,
        userName: request.user?.displayName || "",
        monthId,
        lines: []
      }),
      history: buildMileageHistory(savedClaims)
    });
  });

  app.get("/api/attendance", async (request, response) => {
    if (!requireAttendanceAccess(request, response)) return;
    const payload = await getAttendancePayload(request.user, String(request.query.month || "").trim());
    response.json(payload);
  });

  app.post("/api/attendance/entries", async (request, response) => {
    if (!requireAttendanceAdmin(request, response)) return;
    const nextEntry = sanitizeAttendanceEntry(request.body || {});
    if (!nextEntry.person || !isValidIsoDate(nextEntry.date)) {
      response.status(400).json({ error: "A valid employee and date are required." });
      return;
    }

    const store = await readStore();
    const usersStore = await readUsersStore();
    store.attendanceEntries = Array.isArray(store.attendanceEntries) ? store.attendanceEntries : [];
    const identity = getHolidayStaffIdentityKey(nextEntry.person);
    const existingIndex = store.attendanceEntries.findIndex(
      (entry) =>
        getHolidayStaffIdentityKey(entry.person) === identity &&
        String(entry.date || "") === nextEntry.date
    );

    if (existingIndex >= 0) {
      nextEntry.id = store.attendanceEntries[existingIndex].id;
      nextEntry.createdAt = store.attendanceEntries[existingIndex].createdAt || nextEntry.createdAt;
      nextEntry.employeeNote = store.attendanceEntries[existingIndex].employeeNote || nextEntry.employeeNote;
      store.attendanceEntries[existingIndex] = nextEntry;
    } else {
      store.attendanceEntries.unshift(nextEntry);
    }

    syncAttendanceMissingNotification(store, usersStore.users || [], nextEntry);
    const savedStore = await writeStore(store);
    broadcast("attendance-updated", { monthId: String(nextEntry.date || "").slice(0, 7), date: nextEntry.date });
    const payload = await getAttendancePayload(request.user, String(nextEntry.date || "").slice(0, 7));
    response.json(payload);
  });

  app.post("/api/attendance/explanations", async (request, response) => {
    if (!requireAttendanceAccess(request, response)) return;
    const targetDate = String(request.body?.date || "").trim();
    const note = String(request.body?.employeeNote || "").trim();
    if (!isValidIsoDate(targetDate)) {
      response.status(400).json({ error: "A valid attendance date is required." });
      return;
    }

    const person = getHolidayStaffPerson(request.user?.displayName) || request.user?.displayName || "";
    if (!person) {
      response.status(400).json({ error: "This user is not linked to an attendance record." });
      return;
    }

    const store = await readStore();
    const usersStore = await readUsersStore();
    store.attendanceEntries = Array.isArray(store.attendanceEntries) ? store.attendanceEntries : [];
    const identity = getHolidayStaffIdentityKey(person);
    const existingIndex = store.attendanceEntries.findIndex(
      (entry) =>
        getHolidayStaffIdentityKey(entry.person) === identity &&
        String(entry.date || "") === targetDate
    );

    let nextEntry = sanitizeAttendanceEntry({
      person,
      date: targetDate,
      employeeNote: note
    });

    if (existingIndex >= 0) {
      nextEntry = sanitizeAttendanceEntry({
        ...store.attendanceEntries[existingIndex],
        employeeNote: note
      });
      store.attendanceEntries[existingIndex] = nextEntry;
    } else {
      store.attendanceEntries.unshift(nextEntry);
    }

    const adminRecipients = (usersStore.users || []).filter((user) => canEditAttendance(sanitizeUser(user)));
    const explanationMessage = `${person} added an attendance note for ${formatBoardNotificationDate(targetDate)}${note ? ` (${note})` : "."}`;
    adminRecipients.forEach((user) => {
      store.notifications.unshift(
        createNotification({
          userId: user.id,
          type: "attendance-note",
          title: "Attendance note added",
          message: explanationMessage,
          link: getAttendanceLinkForUser(user, targetDate)
        })
      );
    });

    syncAttendanceMissingNotification(store, usersStore.users || [], nextEntry);
    const savedStore = await writeStore(store);
    broadcast("attendance-updated", { monthId: String(targetDate).slice(0, 7), date: targetDate });
    const payload = await getAttendancePayload(request.user, String(targetDate).slice(0, 7));
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

    const store = await readStore();
    const usersStore = await readUsersStore();
    const holidayStaffList = buildHolidayStaffList(usersStore.users || [], store.holidayAllowances || []);
    if (!holidayStaffList.some((entry) => getHolidayStaffIdentityKey(entry.person) === getHolidayStaffIdentityKey(nextRequest.person))) {
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
    if (!canEditHolidays(request.user) && nextRequest.action === "book") {
      const allowanceSummary = buildHolidayAllowanceSummaries(store, requestYearStart, holidayStaffList).find(
          (entry) => getHolidayStaffIdentityKey(entry.person) === getHolidayStaffIdentityKey(nextRequest.person)
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
    store.notifications = Array.isArray(store.notifications) ? store.notifications : [];
    store.holidayRequests.unshift(nextRequest);

    const recipientUsers = getHolidayNotificationRecipients(usersStore.users || [], [request.user?.id || ""]);
    const requestDateRange = formatHolidayRequestDateRange(nextRequest.startDate, nextRequest.endDate);
    recipientUsers.forEach((user) => {
      store.notifications.unshift(
        createNotification({
          userId: user.id,
          type: nextRequest.action === "cancel" ? "holiday-request" : "holiday-request",
          title: nextRequest.action === "cancel" ? "Holiday cancellation requested" : "Holiday request received",
          message: nextRequest.action === "cancel"
            ? `${nextRequest.requestedByName || nextRequest.person} requested cancellation for ${requestDateRange}${nextRequest.notes ? ` (${nextRequest.notes})` : ""}.`
            : `${nextRequest.requestedByName || nextRequest.person} requested ${requestDateRange}${nextRequest.notes ? ` (${nextRequest.notes})` : ""}.`,
          link: "/holidays"
        })
      );
    });

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
    store.notifications = Array.isArray(store.notifications) ? store.notifications : [];
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

    if (status === "approved" && holidayRequest.action !== "cancel") {
      const requestDates = enumerateIsoDates(holidayRequest.startDate, holidayRequest.endDate).filter(isWeekdayIsoDate);
      for (const date of requestDates) {
        const holidayEntry = sanitizeStaffHoliday({
          id: makeId(),
          date,
          person: holidayRequest.person,
          duration: holidayRequest.duration
        });

          const existingIndex = store.holidays.findIndex(
            (entry) =>
              String(entry.date || "") === date &&
              getHolidayType(entry) !== "birthday" &&
              getHolidayStaffIdentityKey(entry.person) === getHolidayStaffIdentityKey(holidayRequest.person)
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

    if (status === "approved" && holidayRequest.action === "cancel") {
      const requestDates = enumerateIsoDates(holidayRequest.startDate, holidayRequest.endDate).filter(isWeekdayIsoDate);
      store.holidays = (store.holidays || []).filter((entry) => {
        if (getHolidayType(entry) === "birthday") return true;
        const samePerson = getHolidayStaffIdentityKey(entry.person) === getHolidayStaffIdentityKey(holidayRequest.person);
        const sameDate = requestDates.includes(String(entry.date || ""));
        return !(samePerson && sameDate);
      });

      if (holidayRequest.targetRequestId) {
        store.holidayRequests = store.holidayRequests.filter(
          (entry, index) =>
            index === requestIndex ||
            String(entry.id || "") !== String(holidayRequest.targetRequestId || "")
        );
      }
    }

    if (holidayRequest.requestedByUserId) {
      const requestDateRange = formatHolidayRequestDateRange(holidayRequest.startDate, holidayRequest.endDate);
      store.notifications.unshift(
        createNotification({
          userId: holidayRequest.requestedByUserId,
          type: status === "approved" ? "holiday-approved" : "holiday-declined",
          title:
            holidayRequest.action === "cancel"
              ? status === "approved"
                ? "Holiday cancellation approved"
                : "Holiday cancellation declined"
              : status === "approved"
                ? "Holiday approved"
                : "Holiday declined",
          message:
            holidayRequest.action === "cancel"
              ? `${requestDateRange} cancellation for ${holidayRequest.person} was ${status} by ${request.user?.displayName || "admin"}.`
              : `${requestDateRange} for ${holidayRequest.person} was ${status} by ${request.user?.displayName || "admin"}.`,
          link: "/holidays"
        })
      );
    }

    const savedStore = await writeStore(store);
    const payloadYearStart = getHolidayYearStartForIsoDate(holidayRequest.startDate) || getCurrentHolidayYearStart();
    const visiblePayload = await getHolidayPayload(request.user, payloadYearStart);
    broadcast("board-updated", buildBoardRowsFromStore(savedStore));
    response.json(visiblePayload);
  });

  app.delete("/api/holiday-requests/:id", async (request, response) => {
    if (!requireHolidayAccess(request, response)) return;

    const store = await readStore();
    store.holidayRequests = Array.isArray(store.holidayRequests) ? store.holidayRequests : [];
    const requestIndex = store.holidayRequests.findIndex((item) => String(item.id || "") === String(request.params.id || ""));
    if (requestIndex === -1) {
      response.status(404).json({ error: "Holiday request not found." });
      return;
    }

    const holidayRequest = store.holidayRequests[requestIndex];
    const ownsRequest = String(holidayRequest.requestedByUserId || "") === String(request.user?.id || "");
    if (!canEditHolidays(request.user) && !ownsRequest) {
      response.status(403).json({ error: "You can only cancel your own holiday requests." });
      return;
    }

    if (String(holidayRequest.status || "").toLowerCase() === "approved") {
      const requestDates = enumerateIsoDates(holidayRequest.startDate, holidayRequest.endDate).filter(isWeekdayIsoDate);
      store.holidays = (store.holidays || []).filter((entry) => {
        if (getHolidayType(entry) === "birthday") return true;
        const samePerson = getHolidayStaffIdentityKey(entry.person) === getHolidayStaffIdentityKey(holidayRequest.person);
        const sameDate = requestDates.includes(String(entry.date || ""));
        return !(samePerson && sameDate);
      });
    }

    store.holidayRequests.splice(requestIndex, 1);
    const savedStore = await writeStore(store);
    const payloadYearStart = getHolidayYearStartForIsoDate(holidayRequest.startDate) || getCurrentHolidayYearStart();
    const visiblePayload = await getHolidayPayload(request.user, payloadYearStart);
    broadcast("board-updated", buildBoardRowsFromStore(savedStore));
    response.json(visiblePayload);
  });

  app.post("/api/holiday-allowances", async (request, response) => {
    if (!requireHolidayAdmin(request, response)) return;

    const nextAllowance = sanitizeHolidayAllowance(request.body || {});
    const usersStore = await readUsersStore();
    const holidayStaffList = buildHolidayStaffList(usersStore.users || [], []);
    if (!holidayStaffList.some((entry) => getHolidayStaffIdentityKey(entry.person) === getHolidayStaffIdentityKey(nextAllowance.person))) {
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

  app.post("/api/holiday-events", async (request, response) => {
    if (!requireHolidayAdmin(request, response)) return;

    const nextEvent = sanitizeHolidayEvent(request.body || {});
    if (!nextEvent.title || !isValidIsoDate(nextEvent.date)) {
      response.status(400).json({ error: "A valid date and event title are required." });
      return;
    }

    const store = await readStore();
    store.holidayEvents = Array.isArray(store.holidayEvents) ? store.holidayEvents : [];
    const existingIndex = store.holidayEvents.findIndex((entry) => String(entry.id || "") === String(nextEvent.id || ""));

    if (existingIndex >= 0) {
      nextEvent.createdAt = store.holidayEvents[existingIndex].createdAt || nextEvent.createdAt;
      store.holidayEvents[existingIndex] = nextEvent;
    } else {
      store.holidayEvents.unshift(nextEvent);
    }

    await writeStore(store);
    const payload = await getHolidayPayload(request.user, getHolidayYearStartForIsoDate(nextEvent.date) || getCurrentHolidayYearStart());
    response.json(payload);
  });

  app.delete("/api/holiday-events/:id", async (request, response) => {
    if (!requireHolidayAdmin(request, response)) return;
    const store = await readStore();
    store.holidayEvents = (store.holidayEvents || []).filter((entry) => String(entry.id || "") !== String(request.params.id || ""));
    await writeStore(store);
    const payload = await getHolidayPayload(request.user, Number(request.query.yearStart || getCurrentHolidayYearStart()));
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
      const nextIdentity = getHolidayStaffIdentityKey(nextHoliday.person);
      const matchingIndexes = store.holidays.reduce((indexes, item, index) => {
        if (item.id === nextHoliday.id) {
          indexes.push(index);
        return indexes;
      }

        if (
          getHolidayType(item) !== "birthday" &&
          getHolidayType(nextHoliday) !== "birthday" &&
          String(item.date || "") === String(nextHoliday.date || "") &&
          getHolidayStaffIdentityKey(item.person) === nextIdentity
        ) {
          indexes.push(index);
        }

      return indexes;
    }, []);

    if (matchingIndexes.length > 0) {
      const canonicalIndex = matchingIndexes[0];
      const canonicalEntry = store.holidays[canonicalIndex];
      nextHoliday.createdAt = canonicalEntry.createdAt || nextHoliday.createdAt;
      nextHoliday.id = canonicalEntry.id || nextHoliday.id;
      store.holidays = store.holidays.filter((_, index) => !matchingIndexes.includes(index));
      store.holidays.unshift(nextHoliday);
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
      const deleteDate = String(request.query?.date || "").trim();
      const deletePerson = String(request.query?.person || "").trim();
      const deleteIdentity = deletePerson ? getHolidayStaffIdentityKey(deletePerson) : "";
      store.holidays = store.holidays.filter((item) => {
        if (item.id === request.params.id) return false;
        if (
          deleteDate &&
          deleteIdentity &&
          getHolidayType(item) !== "birthday" &&
          String(item.date || "") === deleteDate &&
          getHolidayStaffIdentityKey(item.person) === deleteIdentity
        ) {
          return false;
        }
        return true;
      });
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
