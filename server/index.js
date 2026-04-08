const express = require("express");
const cors = require("cors");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");

const PORT = Number(process.env.PORT || 3030);
const HOST = process.env.HOST || "0.0.0.0";
const DEV_FRONTEND_PORT = Number(process.env.DEV_FRONTEND_PORT || 5173);
const DIST_DIR = path.join(__dirname, "..", "dist");
const DEFAULT_DATA_FILE = path.join(__dirname, "..", "data", "jobs.json");
const PUBLIC_DIR = path.join(__dirname, "..", "public");
const TIME_ZONE = "Europe/London";
const streamClients = new Set();
const DEFAULT_COREBRIDGE_BASE_URL = "https://corebridgev3.azure-api.net";
const DEFAULT_COREBRIDGE_ORDER_PATH = "/core/api/order";

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

function ensureStoreFile() {
  const file = getDataFile();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  if (!fs.existsSync(file)) {
    fs.writeFileSync(file, `${JSON.stringify({ jobs: [], holidays: [] }, null, 2)}\n`, "utf8");
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
        bucket[key.toLowerCase()] = joined;
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
        bucket[key.toLowerCase()] = stringValue;
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

function normalizeCoreBridgeOrder(record, index) {
  const flat = flattenRecord(record);

  const address = [
    pickFirst(flat, [
      "address1",
      "address.line1",
      "siteaddress1",
      "shiptoaddress1",
      "locationlocators.0.locator",
      "locator"
    ]),
    pickFirst(flat, ["address2", "address.line2", "siteaddress2", "shiptoaddress2"]),
    pickFirst(flat, ["address3", "address.line3", "siteaddress3", "shiptoaddress3"]),
    pickFirst(flat, ["city", "address.city", "town", "sitecity", "shiptocity"]),
    pickFirst(flat, ["county", "address.county", "state", "sitestate"]),
    pickFirst(flat, ["postcode", "postalcode", "zip", "address.postcode", "sitepostcode", "shiptopostcode"])
  ]
    .filter(Boolean)
    .join(", ");

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
    description: pickFirst(flat, [
      "description",
      "jobdescription",
      "title",
      "summary",
      "projectdescription",
      "items.0.name",
      "name"
    ]),
    contact: pickFirst(flat, [
      "contact",
      "contactname",
      "primarycontact",
      "contactperson",
      "customercontact",
      "contactroles.0.contactname"
    ]),
    number: pickFirst(flat, [
      "phone",
      "telephone",
      "mobilenumber",
      "contactphone",
      "contactnumber",
      "contactroles.0.ordercontactrolelocators.0.locator"
    ]),
    address: address || pickFirst(flat, ["address", "siteaddress", "shiptoaddress"]),
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

async function fetchCoreBridgeOrders(searchTerm = "") {
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
      const body = contentType.includes("application/json") ? await response.json() : JSON.parse(await response.text());
      const records = extractCoreBridgeRecords(body);
      const orders = filterCoreBridgeOrders(records.map(normalizeCoreBridgeOrder), normalizedSearch).filter(
        (order) => order.orderReference || order.customerName
      );

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
  const app = express();

  app.use(cors());
  app.use(express.json({ limit: "1mb" }));
  app.use(express.static(PUBLIC_DIR));

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

  app.get("/api/corebridge/orders", async (request, response) => {
    try {
      const searchTerm = String(request.query.q || "").trim();
      const payload = await fetchCoreBridgeOrders(searchTerm);
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

  app.post("/api/jobs", async (request, response) => {
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

  app.get("/api/holidays", async (request, response) => {
    const store = await readStore();
    response.json(store.holidays);
  });

  app.post("/api/holidays", async (request, response) => {
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
  startServer
};
