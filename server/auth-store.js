const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");
const { snapshotFile } = require("./backup-store");

const DEFAULT_USERS_FILE = path.join(__dirname, "..", "data", "users.json");
const PERMISSION_VALUES = ["admin", "user", "none"];
const REMOVED_LEGACY_USERS = new Set(["tamas"]);
const SEEDED_USERS = [
  { displayName: "Matt Rutlidge", role: "host" },
  { displayName: "Tom Van-Boyd", role: "host" },
  { displayName: "Amber Hardman", role: "client" },
  { displayName: "Eddy D'Antonio", role: "client" },
  { displayName: "Kyle Wright", role: "client" },
  { displayName: "Keilan Curtis", role: "client" },
  { displayName: "Matt Carroll", role: "client" },
  { displayName: "Paul Morris", role: "client" },
  { displayName: "Dawn Dewhurst", role: "client" }
];

function makeId() {
  return crypto.randomUUID();
}

function getUsersFile() {
  if (process.env.AUTH_USERS_FILE || process.env.USERS_FILE) {
    return process.env.AUTH_USERS_FILE || process.env.USERS_FILE;
  }

  if (process.env.DATA_FILE) {
    return path.join(path.dirname(process.env.DATA_FILE), "users.json");
  }

  return DEFAULT_USERS_FILE;
}

function getDefaultPermissions(role) {
  if (String(role || "").toLowerCase() === "host") {
    return {
      board: "admin",
      installer: "admin",
      holidays: "admin",
      attendance: "admin",
      mileage: "admin",
      vanEstimator: "none",
      rams: "admin",
      socialPost: "admin",
      descriptionPull: "admin"
    };
  }

  return {
    board: "user",
    installer: "none",
    holidays: "user",
    attendance: "user",
    mileage: "user",
    vanEstimator: "none",
    rams: "none",
    socialPost: "none",
    descriptionPull: "none"
  };
}

function getDefaultAttendanceProfile() {
  return {
    mode: "required",
    contractedHours: {
      monday: { in: "", out: "", off: false },
      tuesday: { in: "", out: "", off: false },
      wednesday: { in: "", out: "", off: false },
      thursday: { in: "", out: "", off: false },
      friday: { in: "", out: "", off: false }
    }
  };
}

function normalizeAttendanceTime(value) {
  const raw = String(value || "").trim();
  const match = raw.match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return "";
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (!Number.isInteger(hours) || !Number.isInteger(minutes)) return "";
  if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) return "";
  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function normalizeAttendanceProfile(profile) {
  const defaults = getDefaultAttendanceProfile();
  const normalizedMode = String(profile?.mode || "").trim().toLowerCase();
  const mode = ["required", "fixed", "exempt"].includes(normalizedMode) ? normalizedMode : defaults.mode;
  const contractedHours = {};
  for (const day of Object.keys(defaults.contractedHours)) {
    contractedHours[day] = {
      in: normalizeAttendanceTime(profile?.contractedHours?.[day]?.in),
      out: normalizeAttendanceTime(profile?.contractedHours?.[day]?.out),
      off: Boolean(profile?.contractedHours?.[day]?.off)
    };
  }
  return { mode, contractedHours };
}

function normalizeProfileText(value, maxLength = 120) {
  return String(value || "").replace(/\s+/g, " ").trim().slice(0, maxLength);
}

function normalizeQualifications(value) {
  const entries = Array.isArray(value)
    ? value
    : String(value || "")
      .split(",")
      .map((entry) => entry.trim());
  return [...new Set(entries.map((entry) => normalizeProfileText(entry, 80)).filter(Boolean))].slice(0, 30);
}

function normalizePhotoDataUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (!/^data:image\/(png|jpe?g|webp);base64,/i.test(raw)) return "";
  return raw.length <= 750000 ? raw : "";
}

function normalizeUserProfile(user) {
  return {
    jobTitle: normalizeProfileText(user?.jobTitle || ""),
    phoneNumber: normalizeProfileText(user?.phoneNumber || "", 40),
    qualifications: normalizeQualifications(user?.qualifications),
    photoDataUrl: normalizePhotoDataUrl(user?.photoDataUrl)
  };
}

function normalizePermissionValue(value, fallback) {
  const normalized = String(value || "").trim().toLowerCase();
  return PERMISSION_VALUES.includes(normalized) ? normalized : fallback;
}

function normalizePermissions(permissions, role) {
  const defaults = getDefaultPermissions(role);
  return {
    board: normalizePermissionValue(permissions?.board, defaults.board),
    installer: normalizePermissionValue(permissions?.installer, defaults.installer),
    holidays: normalizePermissionValue(permissions?.holidays, defaults.holidays),
    attendance: normalizePermissionValue(permissions?.attendance, defaults.attendance),
    mileage: normalizePermissionValue(permissions?.mileage, defaults.mileage),
    vanEstimator: normalizePermissionValue(permissions?.vanEstimator, defaults.vanEstimator),
    rams: normalizePermissionValue(permissions?.rams, defaults.rams),
    socialPost: normalizePermissionValue(permissions?.socialPost, defaults.socialPost),
    descriptionPull: normalizePermissionValue(permissions?.descriptionPull, defaults.descriptionPull)
  };
}

function isOwnerUser(user) {
  return String(user?.displayName || "").trim().toLowerCase() === "matt rutlidge";
}

function applyOwnerPermissions(user) {
  if (!isOwnerUser(user)) return user;
  return {
    ...user,
    permissions: {
      board: "admin",
      installer: "admin",
      holidays: "admin",
      attendance: "admin",
      mileage: "admin",
      vanEstimator: "admin",
      rams: "admin",
      socialPost: "admin",
      descriptionPull: "admin"
    }
  };
}

function deriveRoleFromPermissions(permissions) {
  if (
    permissions?.board === "user" &&
    permissions?.installer === "none" &&
    permissions?.holidays === "user" &&
    permissions?.attendance === "user" &&
    permissions?.mileage === "user" &&
    permissions?.vanEstimator === "none" &&
    permissions?.rams === "none" &&
    permissions?.socialPost === "none"
  ) {
    return "client";
  }

  return "host";
}

function normalizeStore(parsed, options = {}) {
  const includeMissingSeedUsers = Boolean(options.includeMissingSeedUsers);
  const users = (Array.isArray(parsed?.users) ? parsed.users : []).filter(
    (user) => !REMOVED_LEGACY_USERS.has(String(user?.displayName || "").trim().toLowerCase())
  );
  const map = new Map(users.map((user) => [String(user.displayName || "").toLowerCase(), user]));

  for (const seeded of SEEDED_USERS) {
    const key = seeded.displayName.toLowerCase();
    const existing = map.get(key);
    if (existing) {
      existing.role = seeded.role;
      existing.displayName = seeded.displayName;
      existing.permissions = normalizePermissions(existing.permissions, existing.role);
      continue;
    }

    if (includeMissingSeedUsers) {
      users.push({
        id: makeId(),
        displayName: seeded.displayName,
        role: seeded.role,
        permissions: getDefaultPermissions(seeded.role),
        attendanceProfile: getDefaultAttendanceProfile(),
        passwordSalt: "",
        passwordHash: "",
        passwordUpdatedAt: ""
      });
    }
  }

  for (const user of users) {
    user.permissions = normalizePermissions(user.permissions, user.role);
    user.attendanceProfile = normalizeAttendanceProfile(user.attendanceProfile);
    const profile = normalizeUserProfile(user);
    user.jobTitle = profile.jobTitle;
    user.phoneNumber = profile.phoneNumber;
    user.qualifications = profile.qualifications;
    user.photoDataUrl = profile.photoDataUrl;
    if (isOwnerUser(user)) {
      user.permissions = {
        board: "admin",
        installer: "admin",
        holidays: "admin",
        attendance: "admin",
        mileage: "admin",
        vanEstimator: "admin",
        rams: "admin",
        socialPost: "admin",
        descriptionPull: "admin"
      };
    }
  }

  users.sort((left, right) => {
    if (left.role !== right.role) return left.role.localeCompare(right.role);
    return left.displayName.localeCompare(right.displayName);
  });

  return { users };
}

function ensureUsersFile() {
  const file = getUsersFile();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  if (!fs.existsSync(file)) {
    fs.writeFileSync(file, `${JSON.stringify(normalizeStore({ users: [] }, { includeMissingSeedUsers: true }), null, 2)}\n`, "utf8");
    return;
  }

  try {
    const raw = fs.readFileSync(file, "utf8");
    const parsed = JSON.parse(raw);
    const normalized = normalizeStore(parsed);
    fs.writeFileSync(file, `${JSON.stringify(normalized, null, 2)}\n`, "utf8");
  } catch (error) {
    const normalized = normalizeStore({ users: [] }, { includeMissingSeedUsers: true });
    fs.writeFileSync(file, `${JSON.stringify(normalized, null, 2)}\n`, "utf8");
  }
}

async function readUsersStore() {
  ensureUsersFile();
  const raw = await fsp.readFile(getUsersFile(), "utf8");
  try {
    return normalizeStore(JSON.parse(raw));
  } catch (error) {
    return normalizeStore({ users: [] });
  }
}

async function writeUsersStore(store) {
  ensureUsersFile();
  const normalized = normalizeStore(store);
  const usersFile = getUsersFile();
  await snapshotFile(usersFile, "users");
  await fsp.writeFile(usersFile, `${JSON.stringify(normalized, null, 2)}\n`, "utf8");
  return normalized;
}

function hashPassword(password, salt = crypto.randomBytes(16).toString("hex")) {
  const passwordHash = crypto.scryptSync(String(password), salt, 64).toString("hex");
  return { passwordSalt: salt, passwordHash };
}

function verifyPassword(password, user) {
  if (!user?.passwordSalt || !user?.passwordHash) return false;
  const candidate = crypto.scryptSync(String(password), user.passwordSalt, 64);
  const actual = Buffer.from(user.passwordHash, "hex");
  return actual.length === candidate.length && crypto.timingSafeEqual(actual, candidate);
}

function sanitizeUser(user) {
  const safeUser = applyOwnerPermissions(user);
  const permissions = normalizePermissions(safeUser.permissions, safeUser.role);
  return {
    id: safeUser.id,
    displayName: safeUser.displayName,
    role: deriveRoleFromPermissions(permissions),
    permissions,
    attendanceProfile: normalizeAttendanceProfile(safeUser.attendanceProfile),
    ...normalizeUserProfile(safeUser),
    hasPassword: Boolean(safeUser.passwordHash)
  };
}

async function createUser({ displayName, role = "client", password = "" }) {
  const nextDisplayName = String(displayName || "").trim();
  if (!nextDisplayName) {
    throw new Error("Display name is required.");
  }

  const normalizedRole = String(role || "client").trim().toLowerCase() === "host" ? "host" : "client";
  const store = await readUsersStore();
  const exists = store.users.some(
    (entry) => String(entry.displayName || "").trim().toLowerCase() === nextDisplayName.toLowerCase()
  );

  if (exists) {
    throw new Error(`A user named "${nextDisplayName}" already exists.`);
  }

  const nextUser = {
    id: makeId(),
    displayName: nextDisplayName,
    role: normalizedRole,
    permissions: getDefaultPermissions(normalizedRole),
    attendanceProfile: getDefaultAttendanceProfile(),
    jobTitle: "",
    phoneNumber: "",
    qualifications: [],
    photoDataUrl: "",
    passwordSalt: "",
    passwordHash: "",
    passwordUpdatedAt: ""
  };

  if (String(password || "").trim()) {
    const { passwordSalt, passwordHash } = hashPassword(password);
    nextUser.passwordSalt = passwordSalt;
    nextUser.passwordHash = passwordHash;
    nextUser.passwordUpdatedAt = new Date().toISOString();
  }

  store.users.push(nextUser);
  await writeUsersStore(store);
  return sanitizeUser(nextUser);
}

async function setUserPassword(displayName, password) {
  const store = await readUsersStore();
  const user = store.users.find(
    (entry) => String(entry.displayName || "").toLowerCase() === String(displayName || "").toLowerCase()
  );

  if (!user) {
    throw new Error(`No user found for "${displayName}".`);
  }

  const { passwordSalt, passwordHash } = hashPassword(password);
  user.passwordSalt = passwordSalt;
  user.passwordHash = passwordHash;
  user.passwordUpdatedAt = new Date().toISOString();
  await writeUsersStore(store);
  return sanitizeUser(user);
}

async function setUserPasswordById(userId, password) {
  const store = await readUsersStore();
  const user = store.users.find((entry) => String(entry.id || "") === String(userId || ""));

  if (!user) {
    throw new Error(`No user found for id "${userId}".`);
  }

  if (isOwnerUser(user) && !String(password || "").trim()) {
    throw new Error("Owner account password cannot be cleared.");
  }

  if (!String(password || "").trim()) {
    user.passwordSalt = "";
    user.passwordHash = "";
    user.passwordUpdatedAt = "";
    await writeUsersStore(store);
    return sanitizeUser(user);
  }

  const { passwordSalt, passwordHash } = hashPassword(password);
  user.passwordSalt = passwordSalt;
  user.passwordHash = passwordHash;
  user.passwordUpdatedAt = new Date().toISOString();
  await writeUsersStore(store);
  return sanitizeUser(user);
}

async function updateUserPermissions(userId, permissions) {
  const store = await readUsersStore();
  const user = store.users.find((entry) => String(entry.id || "") === String(userId || ""));

  if (!user) {
    throw new Error(`No user found for id "${userId}".`);
  }

  user.permissions = normalizePermissions(permissions, user.role);
    if (isOwnerUser(user)) {
      user.permissions = {
        board: "admin",
        installer: "admin",
        holidays: "admin",
        attendance: "admin",
        mileage: "admin",
        vanEstimator: "admin",
        rams: "admin",
        socialPost: "admin",
        descriptionPull: "admin"
      };
    }
  await writeUsersStore(store);
  return sanitizeUser(user);
}

async function updateUserAttendanceProfile(userId, attendanceProfile) {
  const store = await readUsersStore();
  const user = store.users.find((entry) => String(entry.id || "") === String(userId || ""));

  if (!user) {
    throw new Error(`No user found for id "${userId}".`);
  }

  user.attendanceProfile = normalizeAttendanceProfile(attendanceProfile);
  await writeUsersStore(store);
  return sanitizeUser(user);
}

async function updateUserProfile(userId, profile) {
  const store = await readUsersStore();
  const user = store.users.find((entry) => String(entry.id || "") === String(userId || ""));

  if (!user) {
    throw new Error(`No user found for id "${userId}".`);
  }

  const nextProfile = normalizeUserProfile(profile);
  user.jobTitle = nextProfile.jobTitle;
  user.phoneNumber = nextProfile.phoneNumber;
  user.qualifications = nextProfile.qualifications;
  user.photoDataUrl = nextProfile.photoDataUrl;
  await writeUsersStore(store);
  return sanitizeUser(user);
}

async function deleteUser(userId) {
  const store = await readUsersStore();
  const user = store.users.find((entry) => String(entry.id || "") === String(userId || ""));

  if (!user) {
    throw new Error(`No user found for id "${userId}".`);
  }

  if (isOwnerUser(user)) {
    throw new Error("The owner account cannot be deleted.");
  }

  store.users = store.users.filter((entry) => String(entry.id || "") !== String(userId || ""));
  await writeUsersStore(store);
  return sanitizeUser(user);
}

async function bootstrapPasswordsFromEnv() {
  const raw = String(process.env.AUTH_BOOTSTRAP_PASSWORDS || "").trim();
  if (!raw) return { updated: 0 };

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (error) {
    throw new Error("AUTH_BOOTSTRAP_PASSWORDS must be valid JSON.");
  }

  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error("AUTH_BOOTSTRAP_PASSWORDS must be a JSON object mapping display names to passwords.");
  }

  const store = await readUsersStore();
  let updated = 0;

  for (const [displayName, password] of Object.entries(parsed)) {
    const normalizedName = String(displayName || "").toLowerCase();
    const user = store.users.find((entry) => String(entry.displayName || "").toLowerCase() === normalizedName);
    if (!user || !password) continue;

    const { passwordSalt, passwordHash } = hashPassword(String(password));
    user.passwordSalt = passwordSalt;
    user.passwordHash = passwordHash;
    user.passwordUpdatedAt = new Date().toISOString();
    updated += 1;
  }

  if (updated > 0) {
    await writeUsersStore(store);
  }

  return { updated };
}

module.exports = {
  bootstrapPasswordsFromEnv,
  createUser,
  deleteUser,
  SEEDED_USERS,
  ensureUsersFile,
  getUsersFile,
  hashPassword,
  readUsersStore,
  sanitizeUser,
  setUserPassword,
  setUserPasswordById,
  updateUserAttendanceProfile,
  updateUserPermissions,
  updateUserProfile,
  verifyPassword,
  writeUsersStore
};
