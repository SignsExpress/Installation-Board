const express = require("express");
const fs = require("node:fs/promises");
const path = require("node:path");
const { randomUUID } = require("node:crypto");
const { writeTextFileAtomically } = require("./backup-store");

const statuses = new Set(["current", "former", "unverified"]);
const reasons = new Set(["unknown", "resigned", "redundancy", "dismissed", "retired", "contract-ended", "other"]);
const MAX_PHOTO_BYTES = 256 * 1024;

function invalid(message) {
  throw Object.assign(new Error(message), { status: 400 });
}

function field(value, limit) {
  if (value == null) return "";
  if (typeof value !== "string" || value.length > limit) invalid("One of the fields is too long or has the wrong format.");
  return value.trim();
}

function date(value) {
  const iso = field(value, 10);
  if (!iso) return "";
  const parsed = new Date(iso);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso) || iso < "1900-01-01" || iso > "2100-12-31" ||
      !Number.isFinite(+parsed) || parsed.toISOString().slice(0, 10) !== iso) {
    invalid("Use valid calendar dates.");
  }
  return iso;
}

function personRecord(input) {
  const name = field(input?.name, 120);
  const status = input?.status;
  if (!name || !statuses.has(status)) invalid("A name and valid status are required.");
  if (!Array.isArray(input.spells) || !input.spells.length || input.spells.length > 30) {
    invalid("Enter between one and 30 employment spells.");
  }
  const spells = input.spells.map((spell) => {
    const start = date(spell?.start);
    const end = date(spell?.end);
    if (start && end && end < start) invalid("A finish date cannot be before its start date.");
    const reason = spell?.reason || "unknown";
    if (!reasons.has(reason)) invalid("Select a valid leaving reason.");
    return {
      start, end, reason,
      notice: date(spell?.notice),
      startLabel: field(spell?.startLabel, 100),
      endLabel: field(spell?.endLabel, 100),
      notes: field(spell?.notes, 2000),
      source: field(spell?.source, 500)
    };
  });
  if (status === "current" && spells.at(-1).end) invalid("Leave the finish date blank for the current spell.");
  return { name, status, role: field(input.role, 120), notes: field(input.notes, 3000), spells };
}

function publicRecord({ photo, ...record }) {
  return { ...record, photoVersion: photo?.version || "" };
}

function createStaffHistoryRouter({ dataDir, isOwner }) {
  const router = express.Router();
  const file = path.join(dataDir, "staff-history.local.json");
  let pending = Promise.resolve();

  async function read() {
    try {
      return JSON.parse(await fs.readFile(file, "utf8"));
    } catch (error) {
      if (error.code === "ENOENT") return [];
      throw error;
    }
  }

  function mutate(action) {
    const operation = pending.then(async () => {
      const records = await read();
      const result = action(records);
      await writeTextFileAtomically(file, JSON.stringify(records));
      return result;
    });
    pending = operation.catch(() => {});
    return operation;
  }

  function route(handler) {
    return (request, response, next) => Promise.resolve(handler(request, response)).catch(next);
  }

  function find(records, id) {
    const record = records.find((entry) => entry.id === id);
    if (!record) throw Object.assign(new Error("Staff record not found."), { status: 404 });
    return record;
  }

  router.use((request, response, next) => {
    response.set("Cache-Control", "private, no-store");
    Promise.resolve(isOwner(request.user)).then((allowed) => {
      if (!allowed) return response.status(403).json({ error: "Staff History is private to Matt Rutlidge." });
      if (!["GET", "HEAD", "OPTIONS"].includes(request.method) &&
          (request.get("X-Staff-History") !== "1" || request.get("Sec-Fetch-Site") === "cross-site")) {
        return response.status(403).json({ error: "Open Staff History in the portal to make changes." });
      }
      next();
    }).catch(next);
  });

  router.get("/", route(async (_request, response) => response.json({ people: (await read()).map(publicRecord) })));

  router.post("/import", route(async (request, response) => {
    const people = request.body?.people;
    if (!Array.isArray(people) || people.length > 500 || !people.length) invalid("Choose a valid Staff History import file.");
    const incoming = people.map(personRecord);
    const result = await mutate((records) => {
      const existing = new Set(records.map((person) => person.name.toLowerCase()));
      let added = 0;
      for (const person of incoming) {
        const key = person.name.toLowerCase();
        if (existing.has(key)) continue;
        if (records.length >= 500) invalid("The register is limited to 500 people.");
        records.push({ ...person, id: randomUUID(), revision: 1 });
        existing.add(key);
        added++;
      }
      return { added, skipped: incoming.length - added, people: records.map(publicRecord) };
    });
    response.json(result);
  }));

  router.post("/people", route(async (request, response) => {
    const record = personRecord(request.body);
    const saved = await mutate((records) => {
      if (records.length >= 500) invalid("The register is limited to 500 people.");
      if (records.some((person) => person.name.toLowerCase() === record.name.toLowerCase())) invalid("A record already exists for that name.");
      const person = { ...record, id: randomUUID(), revision: 1 };
      records.push(person);
      return publicRecord(person);
    });
    response.status(201).json(saved);
  }));

  router.put("/people/:id", route(async (request, response) => {
    const record = personRecord(request.body);
    const saved = await mutate((records) => {
      const person = find(records, request.params.id);
      if (person.revision !== request.body.revision) {
        throw Object.assign(new Error("This record changed. Reload Staff History before saving again."), { status: 409 });
      }
      if (records.some((other) => other !== person && other.name.toLowerCase() === record.name.toLowerCase())) {
        invalid("A record already exists for that name.");
      }
      Object.assign(person, record, { revision: person.revision + 1 });
      return publicRecord(person);
    });
    response.json(saved);
  }));

  router.get("/people/:id/photo", route(async (request, response) => {
    const person = find(await read(), request.params.id);
    if (!person.photo) return response.sendStatus(404);
    response.set("X-Content-Type-Options", "nosniff");
    response.type("image/jpeg").send(Buffer.from(person.photo.jpeg, "base64"));
  }));

  router.put("/people/:id/photo", route(async (request, response) => {
    const jpeg = request.body?.jpeg;
    if (typeof jpeg !== "string" || jpeg.length > Math.ceil(MAX_PHOTO_BYTES / 3) * 4 || !/^[A-Za-z0-9+/]+={0,2}$/.test(jpeg)) {
      invalid("Choose a JPEG photo under 256 KB.");
    }
    const bytes = Buffer.from(jpeg, "base64");
    if (bytes.length < 4 || bytes.length > MAX_PHOTO_BYTES || bytes[0] !== 255 || bytes[1] !== 216 || bytes.at(-2) !== 255 || bytes.at(-1) !== 217) {
      invalid("Choose a valid JPEG photo under 256 KB.");
    }
    const saved = await mutate((records) => {
      const person = find(records, request.params.id);
      person.photo = { jpeg, version: randomUUID() };
      return publicRecord(person);
    });
    response.json(saved);
  }));

  router.delete("/people/:id/photo", route(async (request, response) => {
    const saved = await mutate((records) => {
      const person = find(records, request.params.id);
      delete person.photo;
      return publicRecord(person);
    });
    response.json(saved);
  }));

  router.use((error, _request, response, _next) => {
    if (!error.status) console.error("Staff History storage error:", error.code || error.name);
    response.status(error.status || 500).json({ error: error.status ? error.message : "Could not access Staff History." });
  });
  return router;
}

module.exports = { createStaffHistoryRouter };
