const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");

const DEFAULT_USERS_FILE = path.join(__dirname, "..", "data", "users.json");
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

function normalizeStore(parsed) {
  const users = Array.isArray(parsed?.users) ? parsed.users : [];
  const map = new Map(users.map((user) => [String(user.displayName || "").toLowerCase(), user]));

  for (const seeded of SEEDED_USERS) {
    const key = seeded.displayName.toLowerCase();
    const existing = map.get(key);
    if (existing) {
      existing.role = seeded.role;
      existing.displayName = seeded.displayName;
      continue;
    }

    users.push({
      id: makeId(),
      displayName: seeded.displayName,
      role: seeded.role,
      passwordSalt: "",
      passwordHash: "",
      passwordUpdatedAt: ""
    });
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
  return {
    id: user.id,
    displayName: user.displayName,
    role: user.role,
    hasPassword: Boolean(user.passwordHash)
  };
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

module.exports = {
  SEEDED_USERS,
  ensureUsersFile,
  getUsersFile,
  hashPassword,
  readUsersStore,
  sanitizeUser,
  setUserPassword,
  verifyPassword,
  writeUsersStore
};
