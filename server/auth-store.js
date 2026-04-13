const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");

const DEFAULT_USERS_FILE = path.join(__dirname, "..", "data", "users.json");
const PERMISSION_VALUES = ["admin", "user", "none"];
const SEEDED_USERS = [
  { displayName: "Matt Rutlidge", role: "host" },
  { displayName: "Tom Van-Boyd", role: "host" },
  { displayName: "Tamas", role: "host" },
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
      holidays: "admin"
    };
  }

  return {
    board: "user",
    installer: "none",
    holidays: "user"
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
    holidays: normalizePermissionValue(permissions?.holidays, defaults.holidays)
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
      holidays: "admin"
    }
  };
}

function deriveRoleFromPermissions(permissions) {
  if (permissions?.board === "user" && permissions?.installer === "none" && permissions?.holidays === "user") {
    return "client";
  }

  return "host";
}

function normalizeStore(parsed) {
  const users = Array.isArray(parsed?.users) ? parsed.users : [];
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

    users.push({
      id: makeId(),
      displayName: seeded.displayName,
      role: seeded.role,
      permissions: getDefaultPermissions(seeded.role),
      passwordSalt: "",
      passwordHash: "",
      passwordUpdatedAt: ""
    });
  }

  for (const user of users) {
    user.permissions = normalizePermissions(user.permissions, user.role);
    if (isOwnerUser(user)) {
      user.permissions = {
        board: "admin",
        installer: "admin",
        holidays: "admin"
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
    fs.writeFileSync(file, `${JSON.stringify(normalizeStore({ users: [] }), null, 2)}\n`, "utf8");
    return;
  }

  try {
    const raw = fs.readFileSync(file, "utf8");
    const parsed = JSON.parse(raw);
    const normalized = normalizeStore(parsed);
    fs.writeFileSync(file, `${JSON.stringify(normalized, null, 2)}\n`, "utf8");
  } catch (error) {
    const normalized = normalizeStore({ users: [] });
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
  await fsp.writeFile(getUsersFile(), `${JSON.stringify(normalized, null, 2)}\n`, "utf8");
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
      holidays: "admin"
    };
  }
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
  updateUserPermissions,
  verifyPassword,
  writeUsersStore
};
