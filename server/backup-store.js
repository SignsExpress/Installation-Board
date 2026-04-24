const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");

function toBackupStamp(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");
  return `${year}${month}${day}-${hours}${minutes}${seconds}`;
}

function sanitizeLabel(label) {
  return String(label || "backup")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9-_]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "") || "backup";
}

async function pruneBackups(directory, prefix, keep = 20) {
  const entries = await fsp.readdir(directory, { withFileTypes: true });
  const matching = entries
    .filter((entry) => entry.isFile() && entry.name.startsWith(prefix))
    .map((entry) => entry.name)
    .sort()
    .reverse();
  const removals = matching.slice(keep);
  await Promise.all(removals.map((name) => fsp.unlink(path.join(directory, name))));
}

async function snapshotFile(sourceFile, label, keep = 20) {
  if (!sourceFile || !fs.existsSync(sourceFile)) return;
  const directory = path.join(path.dirname(sourceFile), "backups");
  await fsp.mkdir(directory, { recursive: true });
  const prefix = `${sanitizeLabel(label)}-`;
  const extension = path.extname(sourceFile) || ".json";
  const targetFile = path.join(directory, `${prefix}${toBackupStamp()}${extension}`);
  await fsp.copyFile(sourceFile, targetFile);
  await pruneBackups(directory, prefix, keep);
}

async function writeTextFileAtomically(targetFile, content) {
  const directory = path.dirname(targetFile);
  const tempFile = path.join(
    directory,
    `.${path.basename(targetFile)}.${process.pid}.${Date.now()}.tmp`
  );
  await fsp.mkdir(directory, { recursive: true });
  await fsp.writeFile(tempFile, content, "utf8");
  await fsp.rename(tempFile, targetFile);
}

module.exports = {
  snapshotFile,
  writeTextFileAtomically
};
