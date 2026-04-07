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

const weekdayFormatter = new Intl.DateTimeFormat("en-GB", {
  weekday: "short",
  timeZone: TIME_ZONE
});

const longDateFormatter = new Intl.DateTimeFormat("en-GB", {
  weekday: "long",
  day: "numeric",
  month: "long",
  year: "numeric",
  timeZone: TIME_ZONE
});

const monthFormatter = new Intl.DateTimeFormat("en-GB", {
  month: "long",
  year: "numeric",
  timeZone: TIME_ZONE
});

const dayFormatter = new Intl.DateTimeFormat("en-CA", {
  year: "numeric",
  month: "2-digit",
  day: "2-digit",
  timeZone: TIME_ZONE
});

const streamClients = new Set();

function getDataFile() {
  return process.env.DATA_FILE || DEFAULT_DATA_FILE;
}

function ensureStoreFile() {
  const file = getDataFile();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  if (!fs.existsSync(file)) {
    fs.writeFileSync(file, "[]\n", "utf8");
  }
}

async function readJobs() {
  ensureStoreFile();
  const raw = await fsp.readFile(getDataFile(), "utf8");

  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    console.error("Invalid job store JSON, returning empty array.", error);
    return [];
  }
}

async function writeJobs(jobs) {
  ensureStoreFile();
  const sorted = [...jobs].sort((left, right) => {
    if (left.date !== right.date) return left.date.localeCompare(right.date);
    return String(left.title || "").localeCompare(String(right.title || ""));
  });
  await fsp.writeFile(getDataFile(), `${JSON.stringify(sorted, null, 2)}\n`, "utf8");
  return sorted;
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
    if (day >= 1 && day <= 5) {
      dates.push(cursor);
    }
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

function buildBoardRows(jobs, today = getTodayInLondon()) {
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

  const rows = weekdayDates.map((date) => {
    const isoDate = toIsoDate(date);
    const sameDayJobs = (jobsByDate.get(isoDate) || []).sort((left, right) =>
      String(left.title || "").localeCompare(String(right.title || ""))
    );

    return {
      isoDate,
      dayLabel: weekdayFormatter.format(date).toUpperCase(),
      dayNumber: String(date.getUTCDate()).padStart(2, "0"),
      fullDateLabel: longDateFormatter.format(date),
      holiday: holidayMap.get(isoDate) || "",
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
  return {
    id: String(payload.id || makeId()),
    date: String(payload.date || "").trim(),
    title: String(payload.title || "").trim(),
    crew: String(payload.crew || "").trim(),
    notes: String(payload.notes || "").trim(),
    category: String(payload.category || "Install").trim(),
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
    const jobs = await readJobs();
    response.json(buildBoardRows(jobs));
  });

  app.get("/api/jobs", async (request, response) => {
    const jobs = await readJobs();
    response.json(jobs);
  });

  app.post("/api/jobs", async (request, response) => {
    const nextJob = sanitizeJob(request.body || {});

    if (!nextJob.title || !isValidIsoDate(nextJob.date)) {
      response.status(400).json({ error: "A valid date and job title are required." });
      return;
    }

    const jobs = await readJobs();
    const existingIndex = jobs.findIndex((job) => job.id === nextJob.id);

    if (existingIndex >= 0) {
      nextJob.createdAt = jobs[existingIndex].createdAt || nextJob.createdAt;
      jobs[existingIndex] = nextJob;
    } else {
      jobs.unshift(nextJob);
    }

    const saved = await writeJobs(jobs);
    const board = buildBoardRows(saved);
    broadcast("board-updated", board);
    response.json({ jobs: saved, board });
  });

  app.delete("/api/jobs/:id", async (request, response) => {
    const jobs = await readJobs();
    const saved = await writeJobs(jobs.filter((job) => job.id !== request.params.id));
    const board = buildBoardRows(saved);
    broadcast("board-updated", board);
    response.json({ jobs: saved, board });
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
