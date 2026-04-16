import { useEffect, useMemo, useRef, useState } from "react";
import InstallerDirectoryHost from "./installer/InstallerDirectoryHostV2";

const JOB_TYPES = [
  { value: "Install", colorClass: "job-type-install" },
  { value: "Vehicle", colorClass: "job-type-vehicle" },
  { value: "Delivery", colorClass: "job-type-delivery" },
  { value: "Subcontractor", colorClass: "job-type-subcontractor" },
  { value: "Signs Express", colorClass: "job-type-signs-express" },
  { value: "Survey", colorClass: "job-type-survey" },
  { value: "Other", colorClass: "job-type-other" }
];

const INSTALLER_OPTIONS = [
  { value: "MC", colorClass: "installer-mc" },
  { value: "KC", colorClass: "installer-kc" },
  { value: "ED", colorClass: "installer-ed" },
  { value: "KW", colorClass: "installer-kw" },
  { value: "PM", colorClass: "installer-pm" },
  { value: "MR", colorClass: "installer-mr" },
  { value: "Custom", colorClass: "installer-custom" }
];

const PERMISSION_OPTIONS = [
  { value: "admin", label: "Admin" },
  { value: "user", label: "User" },
  { value: "none", label: "No Access" }
];

const HOLIDAY_STAFF = [
  { code: "MR", person: "Matt R", fullName: "Matt Rutlidge", colorClass: "holiday-person-black", birthDate: "" },
  { code: "DD", person: "Dawn D", fullName: "Dawn Dewhurst", colorClass: "holiday-person-black", birthDate: "" },
  { code: "TVB", person: "Tom V-B", fullName: "Tom Van-Boyd", colorClass: "holiday-person-black", birthDate: "" },
  { code: "AH", person: "Amber H", fullName: "Amber Hardman", colorClass: "holiday-person-black", birthDate: "" },
  { code: "ED", person: "Eddy D'A", fullName: "Eddy D'Antonio", colorClass: "holiday-person-black", birthDate: "" },
  { code: "PM", person: "Paul M", fullName: "Paul Morris", colorClass: "holiday-person-green", birthDate: "" },
  { code: "KW", person: "Kyle W", fullName: "Kyle Wright", colorClass: "holiday-person-green", birthDate: "" },
  { code: "MC", person: "Matt C", fullName: "Matt Carroll", colorClass: "holiday-person-red", birthDate: "" },
  { code: "KC", person: "Keilan C", fullName: "Keilan Curtis", colorClass: "holiday-person-red", birthDate: "" }
];
const HOLIDAY_PERSON_COLORS = Object.fromEntries(HOLIDAY_STAFF.map((entry) => [entry.person, entry.colorClass]));
const UNSCHEDULED_DROP_ZONE = "__unscheduled__";

const EMPTY_FORM = {
  id: "",
  date: "",
  orderReference: "",
  customerName: "",
  description: "",
  contact: "",
  number: "",
  address: "",
  installers: [],
  customInstaller: "",
  jobType: "Install",
  customJobType: "",
  isPlaceholder: false,
  notes: ""
};

const EMPTY_HOLIDAY_REQUEST_FORM = {
  person: "",
  startDate: "",
  endDate: "",
  duration: "Full Day",
  notes: ""
};

const EMPTY_HOLIDAY_CANCEL_FORM = {
  requestId: "",
  notes: ""
};

const EMPTY_HOLIDAY_EVENT_FORM = {
  id: "",
  date: "",
  title: ""
};

const EMPTY_ATTENDANCE_NOTE_FORM = {
  date: "",
  note: ""
};

const ATTENDANCE_WEEKDAYS = [
  ["monday", "Mon"],
  ["tuesday", "Tue"],
  ["wednesday", "Wed"],
  ["thursday", "Thu"],
  ["friday", "Fri"]
];

const VAN_REFERENCE_TYRE_DIAMETER_MM = 686.5;
const VAN_REFERENCE_TYRE_DIAMETER_UNITS = 194.62;

const VAN_ESTIMATOR_TEMPLATE = {
  name: "Ford Transit Custom SWB",
  src: "/vans/ford-transit-custom-swb.svg",
  scaleFactor: VAN_REFERENCE_TYRE_DIAMETER_MM / VAN_REFERENCE_TYRE_DIAMETER_UNITS,
  scaleReference: "Calibrated from 686.5mm front tyre",
  viewBox: { x: 0, y: 0, width: 2280.56, height: 1298.24 }
};

const VEHICLE_GRAPHICS_PRICING = {
  standardVinylRate: 85,
  wrapFilmRate: 110,
  labourSellRate: 160,
  standardLargeHoursPerM2: 0.8,
  standardSmallHoursPerM2: 0.75,
  standardSmallMinHours: 1,
  standardSmallMaxHours: 2,
  partialWrapFlatHoursPerM2: 1.4,
  partialWrapCurvedHoursPerM2: 1.6,
  partialWrapComplexHoursPerM2: 1.8,
  wrapFlatHoursPerM2: 2,
  wrapCurvedHoursPerM2: 2.25,
  wrapComplexHoursPerM2: 2.5,
  minPrice: 250,
  minAnyWrapPrice: 600,
  minWrapLedPartialPrice: 900,
  minFullWrapPrice: 1800
};

const VEHICLE_ZONE_MATERIALS = {
  standard_vinyl: { label: "Standard vinyl", rate: VEHICLE_GRAPHICS_PRICING.standardVinylRate },
  wrap_film: { label: "Wrap film", rate: VEHICLE_GRAPHICS_PRICING.wrapFilmRate }
};

function getLocalTodayIso() {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/London",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).format(new Date());
}

function parseIsoDate(isoDate) {
  const [year, month, day] = String(isoDate || "").split("-").map(Number);
  if (!year || !month || !day) return null;
  return new Date(Date.UTC(year, month - 1, day));
}

function toIsoDate(date) {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "UTC",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).format(date);
}

function addDays(date, days) {
  const next = new Date(date);
  next.setUTCDate(next.getUTCDate() + days);
  return next;
}

function addMonths(date, months) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + months, 1));
}

function getStartOfMonth(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1));
}

function getEndOfMonth(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0));
}

function getHolidayStaffEntry(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return (
    HOLIDAY_STAFF.find((entry) => entry.fullName.toLowerCase() === normalized) ||
    HOLIDAY_STAFF.find((entry) => entry.person.toLowerCase() === normalized) ||
    HOLIDAY_STAFF.find((entry) => entry.code.toLowerCase() === normalized) ||
    null
  );
}

function getHolidayStaffPersonForUser(user) {
  return getHolidayStaffEntry(user?.displayName)?.person || "";
}

function getHolidayStaffIdentityKey(value) {
  const match = getHolidayStaffEntry(value);
  if (match?.code) return match.code.toLowerCase();
  return String(value || "").trim().toLowerCase();
}

function getHolidayDisplayToken(person) {
  const entry = getHolidayStaffEntry(person);
  if (entry) return entry.code;
  const value = String(person || "").trim();
  if (value && !value.includes(" ") && value.length <= 4) return value.toUpperCase();
  return toInitials(value);
}

function normalizeHolidayStaffEntries(entries) {
  if (!Array.isArray(entries) || !entries.length) return HOLIDAY_STAFF;
  return entries.map((entry) => ({
    ...entry,
    fullName: entry.fullName || entry.name || entry.person || ""
  }));
}

function toAllowanceNumber(value) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : 0;
}

function normalizeAttendanceDraft(profile) {
  const base = {
    mode: "required",
    contractedHours: Object.fromEntries(
      ATTENDANCE_WEEKDAYS.map(([dayKey]) => [dayKey, { in: "", out: "", off: false }])
    )
  };
  const nextMode = ["required", "fixed", "exempt"].includes(String(profile?.mode || "").trim().toLowerCase())
    ? String(profile.mode).trim().toLowerCase()
    : "required";

  for (const [dayKey] of ATTENDANCE_WEEKDAYS) {
    base.contractedHours[dayKey] = {
      in: String(profile?.contractedHours?.[dayKey]?.in || "").trim(),
      out: String(profile?.contractedHours?.[dayKey]?.out || "").trim(),
      off: Boolean(profile?.contractedHours?.[dayKey]?.off)
    };
  }

  return {
    mode: nextMode,
    contractedHours: base.contractedHours
  };
}

function getHolidayAllowanceSummary(entry) {
  const standardEntitlement = toAllowanceNumber(entry.standardEntitlement);
  const extraServiceDays = toAllowanceNumber(entry.extraServiceDays);
  const christmasDays = toAllowanceNumber(entry.christmasDays);
  const bankHolidayDays = toAllowanceNumber(entry.bankHolidayDays);
  const bookedDays = toAllowanceNumber(entry.bookedDays);
  const unpaidDaysBooked = toAllowanceNumber(entry.unpaidDaysBooked);
  const workDaysPerWeek = toAllowanceNumber(entry.workDaysPerWeek);
  const prorataAllowance = standardEntitlement + extraServiceDays;
  const daysLeft = prorataAllowance - christmasDays - bankHolidayDays - bookedDays;

  return {
    ...entry,
    workDaysPerWeek,
    standardEntitlement,
    extraServiceDays,
    christmasDays,
    bankHolidayDays,
    bookedDays,
    unpaidDaysBooked,
    prorataAllowance,
    daysLeft
  };
}

function formatHolidayBirthday(value) {
  const parsed = parseIsoDate(value);
  if (!parsed) return "-";
  return parsed.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    timeZone: "UTC"
  });
}

function isBirthdayHoliday(entry) {
  return String(entry?.type || "").trim().toLowerCase() === "birthday";
}

function getCurrentHolidayYearStart(anchorIsoDate = getLocalTodayIso()) {
  const anchor = parseIsoDate(anchorIsoDate) || parseIsoDate(getLocalTodayIso());
  if (!anchor) return new Date().getUTCFullYear();
  return anchor.getUTCMonth() >= 1 ? anchor.getUTCFullYear() : anchor.getUTCFullYear() - 1;
}

function getHolidayYearLabel(yearStart) {
  return `${yearStart}-${String(yearStart + 1).slice(-2)}`;
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

function buildHolidayYearRows(holidays, yearStart, holidayEvents = []) {
  const startMonth = new Date(Date.UTC(yearStart, 1, 1));
  const holidayMap = new Map();
  const eventMap = new Map();
  const years = new Set();
  const rows = [];

  for (let offset = 0; offset < 12; offset += 1) {
    const monthDate = addMonths(startMonth, offset);
    rows.push(monthDate);
    years.add(monthDate.getUTCFullYear());
  }

  const bankHolidayMap = new Map();
  years.forEach((year) => {
    getUkBankHolidays(year).forEach((holiday) => bankHolidayMap.set(holiday.date, holiday.label));
  });

  holidays.forEach((holiday) => {
    const bucket = holidayMap.get(holiday.date) || [];
    bucket.push(holiday);
    holidayMap.set(holiday.date, bucket);
  });

  holidayEvents.forEach((event) => {
    const bucket = eventMap.get(event.date) || [];
    bucket.push(event);
    eventMap.set(event.date, bucket);
  });

  return rows.map((monthDate) => {
    const year = monthDate.getUTCFullYear();
    const month = monthDate.getUTCMonth();
    const monthLabel = monthDate.toLocaleString("en-GB", { month: "short", year: "2-digit", timeZone: "UTC" });
    const days = Array.from({ length: 31 }, (_, index) => {
      const dayNumber = index + 1;
      const candidate = new Date(Date.UTC(year, month, dayNumber));
      const inMonth = candidate.getUTCMonth() === month;
      const isoDate = inMonth ? toIsoDate(candidate) : "";
      const weekday = candidate.getUTCDay();
      return {
        key: `${year}-${month + 1}-${dayNumber}`,
        dayNumber,
          isoDate,
          inMonth,
          weekend: inMonth && (weekday === 0 || weekday === 6),
          bankHoliday: inMonth ? bankHolidayMap.get(isoDate) || "" : "",
          holidays: inMonth ? (holidayMap.get(isoDate) || []) : [],
          events: inMonth ? (eventMap.get(isoDate) || []) : []
        };
      });

    return {
      id: `${year}-${String(month + 1).padStart(2, "0")}`,
      label: monthLabel,
      days
    };
  });
}

function createMessage(text, tone = "info") {
  return { text, tone, id: `${Date.now()}-${Math.random()}` };
}

function buildJobPhotoUrl(jobId, photoId) {
  return `/api/jobs/${encodeURIComponent(jobId)}/photos/${encodeURIComponent(photoId)}`;
}

function formatDateTime(value) {
  if (!value) return "";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return String(value);
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  }).format(parsed);
}

function formatJobDate(value) {
  const parsed = parseIsoDate(value);
  if (!parsed) return String(value || "");
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    timeZone: "Europe/London"
  }).format(parsed);
}

function formatHolidayRequestDateRange(startDate, endDate) {
  const start = formatJobDate(startDate);
  const end = formatJobDate(endDate || startDate);
  if (!start) return "";
  return end && end !== start ? `${start} to ${end}` : start;
}

function formatNotificationDate(value) {
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  }).format(parsed);
}

function toMonthIdFromIso(value) {
  const parsed = parseIsoDate(value);
  if (!parsed) return "";
  return `${parsed.getUTCFullYear()}-${String(parsed.getUTCMonth() + 1).padStart(2, "0")}`;
}

function shiftMonthId(value, offset) {
  const [year, month] = String(value || "").split("-").map(Number);
  if (!year || !month) return toMonthIdFromIso(getLocalTodayIso());
  const next = new Date(Date.UTC(year, month - 1 + offset, 1));
  return `${next.getUTCFullYear()}-${String(next.getUTCMonth() + 1).padStart(2, "0")}`;
}

function getAttendanceDisplayClass(cell) {
  const label = String(cell?.displayLabel || "").trim().toLowerCase();
  if (label.includes("unpaid") || label.includes("absence") || label.includes("absent")) return "is-unpaid";
  if (label.includes("birthday")) return "is-birthday";
  if (label.includes("bank holiday")) return "is-bank-holiday";
  if (label.includes("weekend")) return "is-weekend";
  if (label === "off") return "is-weekend";
  if (label.includes("holiday")) return "is-holiday";
  return "";
}

function formatNotificationMessage(message) {
  const raw = String(message || "").trim();
  if (!raw) return "";

  return raw
    .replace(/\b(\d{4})-(\d{2})-(\d{2})\b/g, (_, year, month, day) => `${day}/${month}/${String(year).slice(-2)}`)
    .replace(/\s+/g, " ")
    .trim();
}

function getNextWeekdayIsoDate(isoDate) {
  const parsed = parseIsoDate(isoDate);
  if (!parsed) return "";
  let cursor = addDays(parsed, 1);
  while (cursor.getUTCDay() === 0 || cursor.getUTCDay() === 6) {
    cursor = addDays(cursor, 1);
  }
  return toIsoDate(cursor);
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Could not read the selected photo."));
    reader.readAsDataURL(file);
  });
}

function loadImageFromDataUrl(dataUrl) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Could not load the selected photo."));
    image.src = dataUrl;
  });
}

async function compressPhotoForUpload(file) {
  const originalDataUrl = await readFileAsDataUrl(file);
  const image = await loadImageFromDataUrl(originalDataUrl);
  const maxDimension = 1600;
  const scale = Math.min(1, maxDimension / Math.max(image.width || 1, image.height || 1));
  const width = Math.max(1, Math.round((image.width || 1) * scale));
  const height = Math.max(1, Math.round((image.height || 1) * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) {
    throw new Error("Could not prepare the selected photo.");
  }
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.drawImage(image, 0, 0, width, height);

  const blob = await new Promise((resolve, reject) => {
    canvas.toBlob((nextBlob) => {
      if (nextBlob) resolve(nextBlob);
      else reject(new Error("Could not compress the selected photo."));
    }, "image/jpeg", 0.72);
  });

  const dataUrl = await readFileAsDataUrl(blob);
  const baseName = String(file.name || "job-photo").replace(/\.[^.]+$/, "") || "job-photo";
  return {
    fileName: `${baseName}.jpg`,
    dataUrl,
    width,
    height,
    size: blob.size
  };
}

function toInitials(name) {
  return String(name || "")
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part[0])
    .join("");
}

function renderJobCardContent({
  job,
  isCondensed = false,
  isClientMode,
  draggingJobId,
  getJobTypeMeta,
  getJobTypeLabel,
  getInstallerDisplayList,
  getInstallerMeta,
  editJob,
  handleDelete,
  setActiveClientJob,
  buildDragPreview,
  getTransparentDragImage,
  clearDragPreview,
  dragPreviewRef,
  dragPositionRef,
  setDraggingJobId,
  duplicatingJobId,
  setDuplicatingJobId,
  setDropDate
}) {
  const meta = getJobTypeMeta(job.jobType);
  const installerLabels = getInstallerDisplayList(job);

  return (
      <div
        key={job.id}
        className={`job-card ${meta.colorClass}-card ${job.isPlaceholder ? "is-placeholder" : ""} ${job.isCompleted ? "is-complete" : ""} ${job.isSnagging ? "is-snagging" : ""} ${isCondensed ? "is-condensed" : ""} ${draggingJobId === job.id ? "is-dragging" : ""}`}
        draggable={!isClientMode}
      onDragStart={(event) => {
        if (isClientMode) return;
        event.dataTransfer.setData("text/plain", job.id);
        event.dataTransfer.effectAllowed = "move";
        const preview = buildDragPreview(event.currentTarget);
        dragPreviewRef.current = preview;
        dragPositionRef.current = { x: event.clientX, y: event.clientY };
        preview.style.left = `${event.clientX + 18}px`;
        preview.style.top = `${event.clientY + 18}px`;
        preview.style.transform = "rotate(-2deg) translateY(0)";
        event.dataTransfer.setDragImage(getTransparentDragImage(), 0, 0);
        setDraggingJobId(job.id);
      }}
      onDragEnd={() => {
        if (isClientMode) return;
        setDraggingJobId("");
        setDropDate("");
        clearDragPreview();
      }}
      onClick={() => {
        if (isClientMode) {
          setActiveClientJob(job);
        } else {
          editJob(job);
        }
      }}
    >
        <div className="job-card-top">
          <div className="job-title-wrap">
            <strong className="job-title-line">
              {job.orderReference ? <span className="job-ref-inline">{job.orderReference}</span> : null}
              <span className="job-customer-inline">{job.customerName}</span>
            </strong>
            <p>{job.description || "No description"}</p>
          </div>
        <div className="job-title-meta">
          {job.isPlaceholder ? <span className="placeholder-status-pill">Placeholder</span> : null}
          {job.isSnagging ? <span className="job-snagging-pill">Snagging</span> : null}
          {job.isCompleted ? <span className="job-complete-pill">Complete</span> : null}
          {Array.isArray(job.photos) && job.photos.length ? <span className="job-photo-pill">{job.photos.length} photo{job.photos.length === 1 ? "" : "s"}</span> : null}
          {installerLabels.length ? (
            <div className="job-title-installers">
              {installerLabels.map((installer) => {
                const metaInstaller = getInstallerMeta(installer);
                return (
                  <span key={`title-${job.id}-${installer}`} className={`installer-badge title-inline ${metaInstaller.colorClass}`}>
                    {installer}
                  </span>
                );
              })}
            </div>
          ) : null}
            <span className={`job-tag ${meta.colorClass}`}>{getJobTypeLabel(job)}</span>
          </div>
        </div>
        {!isCondensed ? (
          <>
            <div className="job-meta-grid">
              <p><b>Address:</b> {job.address || "-"}</p>
              <p><b>Contact:</b> {job.contact || "-"}</p>
              <p><b>Number:</b> {job.number || "-"}</p>
            </div>
            <p className="job-notes compact"><b>Notes:</b> {job.notes || ""}</p>
            <div className="job-actions">
              {!isClientMode ? (
                <>
                  <button className="text-button" type="button" onClick={(event) => { event.stopPropagation(); editJob(job); }}>
                    Edit
                  </button>
                  <button className="text-button danger" type="button" onClick={(event) => { event.stopPropagation(); handleDelete(job.id); }}>
                    Delete
                  </button>
                  <button
                    type="button"
                    className="card-duplicate-handle"
                    draggable
                    onDragStart={(event) => {
                      event.stopPropagation();
                      event.dataTransfer.setData("job-copy", job.id);
                      event.dataTransfer.effectAllowed = "copy";
                      const preview = buildDragPreview(event.currentTarget.closest(".job-card"));
                      dragPreviewRef.current = preview;
                      dragPositionRef.current = { x: event.clientX, y: event.clientY };
                      preview.style.left = `${event.clientX + 18}px`;
                      preview.style.top = `${event.clientY + 18}px`;
                      event.dataTransfer.setDragImage(getTransparentDragImage(), 0, 0);
                      setDuplicatingJobId(job.id);
                    }}
                    onDragEnd={() => {
                      setDuplicatingJobId("");
                      setDropDate("");
                      clearDragPreview();
                    }}
                    title="Drag to copy"
                  >
                    Drag to Copy
                  </button>
                </>
              ) : null}
            </div>
          </>
        ) : null}
      </div>
    );
  }

function getPermissionForApp(user, key) {
  const fallback =
    key === "board"
      ? user?.role === "host"
        ? "admin"
        : "user"
      : key === "vanEstimator"
        ? "none"
      : user?.role === "host"
        ? "admin"
        : "none";
  const value = String(user?.permissions?.[key] || "").trim().toLowerCase();
  return ["admin", "user", "none"].includes(value) ? value : fallback;
}

function canAccessBoard(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "board") !== "none";
}

function canEditBoard(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "board") === "admin";
}

function canAccessInstaller(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "installer") !== "none";
}

function canEditInstaller(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "installer") === "admin";
}

function canAccessHolidays(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "holidays") !== "none";
}

function canEditHolidays(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "holidays") === "admin";
}

function canAccessAttendance(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "attendance") !== "none";
}

function canEditAttendance(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "attendance") === "admin";
}

function canAccessMileage(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "mileage") !== "none";
}

function canEditMileage(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "mileage") === "admin";
}

function canAccessVanEstimator(user) {
  if (user?.canManagePermissions) return true;
  return getPermissionForApp(user, "vanEstimator") !== "none";
}

function usesHostShell(user) {
  return Boolean(
    user &&
      (canAccessInstaller(user) || canEditBoard(user) || canAccessHolidays(user) || canEditAttendance(user) || canAccessMileage(user) || canAccessVanEstimator(user) || user.canManagePermissions)
  );
}

function getHomePathForUser(user) {
  return usesHostShell(user) ? "/" : "/client";
}

function getBoardPathForUser(user) {
  return canEditBoard(user) ? "/board" : "/client/board";
}

function HomeIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M4 11.5 12 5l8 6.5v7a1.5 1.5 0 0 1-1.5 1.5h-4.25V14h-4.5v6H5.5A1.5 1.5 0 0 1 4 18.5z" />
    </svg>
  );
}

function BoardIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M5 5h14a1 1 0 0 1 1 1v12.5A1.5 1.5 0 0 1 18.5 20h-13A1.5 1.5 0 0 1 4 18.5V6a1 1 0 0 1 1-1Zm1.5 3v9h11V8Zm0-1.5h11V6.5h-11Z" />
    </svg>
  );
}

function HolidayIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M7 3h2v2h6V3h2v2h1.5A1.5 1.5 0 0 1 20 6.5v12A1.5 1.5 0 0 1 18.5 20h-13A1.5 1.5 0 0 1 4 18.5v-12A1.5 1.5 0 0 1 5.5 5H7Zm11 6.5h-12v8.5h12Zm-9.5 2h3v3h-3Z" />
    </svg>
  );
}

function AttendanceIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M7 4h10a2 2 0 0 1 2 2v11a3 3 0 0 1-3 3H8a3 3 0 0 1-3-3V6a2 2 0 0 1 2-2Zm0 2v11a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V6Zm2 2h6v2H9Zm0 4h4v2H9Z" />
    </svg>
  );
}

function MileageIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M6.5 5a3.5 3.5 0 0 0-3.5 3.5c0 2.5 3.5 6.5 3.5 6.5S10 11 10 8.5A3.5 3.5 0 0 0 6.5 5Zm0 4.7a1.2 1.2 0 1 1 0-2.4 1.2 1.2 0 0 1 0 2.4ZM17.5 3A3.5 3.5 0 0 0 14 6.5c0 2.5 3.5 6.5 3.5 6.5S21 9 21 6.5A3.5 3.5 0 0 0 17.5 3Zm0 4.7a1.2 1.2 0 1 1 0-2.4 1.2 1.2 0 0 1 0 2.4ZM7 18h11a1 1 0 1 1 0 2H7a4 4 0 0 1-4-4h2a2 2 0 0 0 2 2Zm10-2a4 4 0 0 1-4-4h2a2 2 0 0 0 2 2h1v2Z" />
    </svg>
  );
}

function VanEstimatorIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M4.5 8.5 6.2 5h8.9a2 2 0 0 1 1.7.94l1.58 2.56H20a1 1 0 0 1 1 1v5.25a1.25 1.25 0 0 1-1.25 1.25h-.85a2.5 2.5 0 0 1-4.8 0H9.9a2.5 2.5 0 0 1-4.8 0h-.85A1.25 1.25 0 0 1 3 14.75V10a1.5 1.5 0 0 1 1.5-1.5Zm3-1.5-.75 1.5h4.75V7Zm6 0v1.5h2.55L15.11 7Zm-6 10a1 1 0 1 0 0-2 1 1 0 0 0 0 2Zm9 0a1 1 0 1 0 0-2 1 1 0 0 0 0 2Z" />
    </svg>
  );
}

function InstallerIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M6 4h12a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2h-4l-2 2-2-2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2Zm1.5 4v2h9V8Zm0 4v2h6v-2Z" />
    </svg>
  );
}

function NotificationIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M12 4a4 4 0 0 1 4 4v1.35c0 .83.27 1.63.78 2.29l1.28 1.69c.68.9.04 2.17-1.09 2.17H6.03c-1.13 0-1.77-1.27-1.09-2.17l1.28-1.69A3.75 3.75 0 0 0 7 9.35V8a5 5 0 0 1 5-5Zm0 16a2.75 2.75 0 0 0 2.58-1.8h-5.16A2.75 2.75 0 0 0 12 20Z" />
    </svg>
  );
}

function LogoutIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M10 4H6.5A1.5 1.5 0 0 0 5 5.5v13A1.5 1.5 0 0 0 6.5 20H10v-2H7V6h3Zm6.3 3.8-1.4 1.4 1.8 1.8H10v2h6.7l-1.8 1.8 1.4 1.4L20.5 12z" />
    </svg>
  );
}

function BrandLogoIcon() {
  return <img src="/branding/signs-express-logo.svg" alt="Signs Express" className="host-nav-brand-logo" />;
}

function getNotificationCategory(notification) {
  const title = String(notification?.title || "").toLowerCase();
  const message = String(notification?.message || "").toLowerCase();
  const link = String(notification?.link || "").toLowerCase();

  if (link.includes("/holidays") || title.includes("holiday") || message.includes("holiday")) {
    return { label: "Holiday", icon: HolidayIcon, className: "notification-type-holiday" };
  }
  if (link.includes("/attendance") || title.includes("clock") || message.includes("clock")) {
    return { label: "Attendance", icon: AttendanceIcon, className: "notification-type-update" };
  }
  if (link.includes("/mileage") || title.includes("mileage") || message.includes("miles")) {
    return { label: "Mileage", icon: MileageIcon, className: "notification-type-update" };
  }
  if (link.includes("/board") || title.includes("job") || message.includes("job")) {
    return { label: "Board", icon: BoardIcon, className: "notification-type-board" };
  }
  if (title.includes("message") || message.includes("message")) {
    return { label: "Message", icon: NotificationIcon, className: "notification-type-message" };
  }
  return { label: "Update", icon: NotificationIcon, className: "notification-type-update" };
}

function buildBoardUrl(startIso = "", endIso = "") {
  if (startIso && endIso) {
    return `/api/board?start=${encodeURIComponent(startIso)}&end=${encodeURIComponent(endIso)}`;
  }
  return "/api/board";
}

function MainNavBar({
  currentUser,
  active = "home",
  onLogout,
  notifications = []
}) {
  function goTo(path) {
    window.location.assign(path);
  }

  const boardAllowed = canAccessBoard(currentUser);
  const attendanceAllowed = canAccessAttendance(currentUser);
  const holidaysAllowed = canAccessHolidays(currentUser);
  const mileageAllowed = canAccessMileage(currentUser);
  const vanEstimatorAllowed = canAccessVanEstimator(currentUser);
  const installerAllowed = canAccessInstaller(currentUser);
  const homePath = getHomePathForUser(currentUser);
  const boardPath = getBoardPathForUser(currentUser);
  const attendancePath = "/attendance";
  const holidaysPath = "/holidays";
  const mileagePath = "/mileage";
  const vanEstimatorPath = "/van-estimator";
  const installerPath = "/installer";
  const notificationsPath = "/notifications";
  const unreadNotifications = notifications.filter((entry) => !entry.read);

  return (
    <header className="host-nav-shell">
      <nav className="host-nav">
        <div className="host-nav-inner">
          <button type="button" className="host-nav-brand" onClick={() => goTo(homePath)} aria-label="Go to home">
            <BrandLogoIcon />
          </button>
          <div className="host-nav-links">
            <button
              type="button"
              className={`host-nav-link ${active === "home" ? "active" : ""}`}
              onClick={() => goTo(homePath)}
            >
              <span className="host-nav-link-label">Home</span>
            </button>
            {boardAllowed ? (
              <button
                type="button"
                className={`host-nav-link ${active === "board" ? "active" : ""}`}
                onClick={() => goTo(boardPath)}
              >
                <span className="host-nav-link-label">Installation Board</span>
              </button>
            ) : null}
            {attendanceAllowed ? (
              <button
                type="button"
                className={`host-nav-link ${active === "attendance" ? "active" : ""}`}
                onClick={() => goTo(attendancePath)}
              >
                <span className="host-nav-link-label">Attendance</span>
              </button>
            ) : null}
            {holidaysAllowed ? (
              <button
                type="button"
                className={`host-nav-link ${active === "holidays" ? "active" : ""}`}
                onClick={() => goTo(holidaysPath)}
              >
                <span className="host-nav-link-label">Holidays</span>
              </button>
            ) : null}
            {mileageAllowed ? (
              <button
                type="button"
                className={`host-nav-link ${active === "mileage" ? "active" : ""}`}
                onClick={() => goTo(mileagePath)}
              >
                <span className="host-nav-link-label">Mileage</span>
              </button>
            ) : null}
            <button
              type="button"
              className={`host-nav-link ${active === "van-estimator" ? "active" : ""} ${vanEstimatorAllowed ? "" : "disabled"}`}
              disabled={!vanEstimatorAllowed}
              onClick={() => goTo(vanEstimatorPath)}
              title={vanEstimatorAllowed ? "Open Vinyl Estimator" : "Vinyl Estimator is inactive"}
            >
              <span className="host-nav-link-label">Vinyl Estimator</span>
            </button>
            {installerAllowed ? (
              <button
                type="button"
                className={`host-nav-link ${active === "installer" ? "active" : ""}`}
                onClick={() => goTo(installerPath)}
              >
                <span className="host-nav-link-label">Subcontractor Directory</span>
              </button>
            ) : null}
            <button
              type="button"
              className={`host-nav-link ${active === "notifications" ? "active" : ""}`}
              onClick={() => goTo(notificationsPath)}
            >
              <span className="host-nav-link-label">
                Notifications
                {unreadNotifications.length ? <span className="host-nav-badge inline">{unreadNotifications.length}</span> : null}
              </span>
            </button>
          </div>
          <div className="host-nav-meta">
            <span className="host-nav-user">Logged in as <strong>{currentUser.displayName}</strong></span>
            <button className="host-nav-logout" type="button" onClick={onLogout}>
              <span className="host-nav-link-label">Log out</span>
            </button>
          </div>
        </div>
      </nav>
    </header>
  );
}

function PermissionsPanel({
  currentUser,
  users,
  savingKey,
  onChangePermission,
  onUpdateAttendanceProfile,
  onCreateUser,
  onResetPassword,
  onDeleteUser
}) {
  const [createForm, setCreateForm] = useState({ displayName: "", role: "client", password: "" });
  const [passwordDrafts, setPasswordDrafts] = useState({});
  const [attendanceDrafts, setAttendanceDrafts] = useState({});
  const visibleUsers = [...users].sort((left, right) => left.displayName.localeCompare(right.displayName));

  useEffect(() => {
    setAttendanceDrafts(
      Object.fromEntries(
        users.map((user) => [user.id, normalizeAttendanceDraft(user.attendanceProfile)])
      )
    );
  }, [users]);

  function updateAttendanceDraft(userId, updater) {
    setAttendanceDrafts((current) => {
      const existing = current[userId] || normalizeAttendanceDraft(null);
      const nextValue = typeof updater === "function" ? updater(existing) : updater;
      return {
        ...current,
        [userId]: normalizeAttendanceDraft(nextValue)
      };
    });
  }

  return (
      <section className="panel permissions-panel">
        <div className="permissions-head">
          <h3>User portal</h3>
          <p>Manage access, passwords, contracted hours and attendance rules in one place.</p>
        </div>
        <div className="permissions-admin-tools">
        <input
          type="text"
          placeholder="Full name"
          value={createForm.displayName}
          onChange={(event) => setCreateForm((current) => ({ ...current, displayName: event.target.value }))}
        />
        <select
          value={createForm.role}
          onChange={(event) => setCreateForm((current) => ({ ...current, role: event.target.value }))}
        >
          <option value="client">Client</option>
          <option value="host">Host</option>
        </select>
        <input
          type="password"
          placeholder="Temporary password"
          value={createForm.password}
          onChange={(event) => setCreateForm((current) => ({ ...current, password: event.target.value }))}
        />
        <button
          className="primary-button"
          type="button"
          onClick={async () => {
            await onCreateUser(createForm);
            setCreateForm({ displayName: "", role: "client", password: "" });
          }}
          disabled={!createForm.displayName.trim() || !createForm.password}
        >
          Add user
        </button>
      </div>
      <div className="permissions-grid">
          {visibleUsers.map((user) => {
            const isSelf = user.id === currentUser.id;
            const boardPermission = getPermissionForApp(user, "board");
            const holidaysPermission = getPermissionForApp(user, "holidays");
            const installerPermission = getPermissionForApp(user, "installer");
            const attendancePermission = getPermissionForApp(user, "attendance");
            const mileagePermission = getPermissionForApp(user, "mileage");
            const vanEstimatorPermission = getPermissionForApp(user, "vanEstimator");
            const attendanceProfile = normalizeAttendanceDraft(user.attendanceProfile);
            const attendanceDraft = attendanceDrafts[user.id] || attendanceProfile;
            const attendanceMode = String(attendanceDraft.mode || "required");
            const contractedHours = attendanceDraft.contractedHours || {};
            const exemptFromClocking = attendanceMode === "exempt";
            const fixedHoursMode = attendanceMode === "fixed";
            const attendanceChanged = JSON.stringify(attendanceDraft) !== JSON.stringify(attendanceProfile);

            return (
              <article key={user.id} className="permissions-user-card">
              <div className="permissions-user-head">
                <div className="permissions-user-identity">
                  <strong>{user.displayName}</strong>
                  <span className="permissions-user-meta">{user.role === "host" ? "Host" : "Client"}</span>
                  </div>
                  {isSelf ? <span className="permissions-owner-pill">Owner</span> : null}
                </div>

                <div className="permissions-user-body">
                  <div className="permissions-main-grid">
                    <div className="permissions-app-row">
                      <span className="permissions-app-label">Installation Board</span>
                      <div className="permission-segment">
                        {PERMISSION_OPTIONS.map((option) => (
                          <button
                            key={`${user.id}-board-${option.value}`}
                            type="button"
                            className={`permission-chip ${boardPermission === option.value ? "active" : ""}`}
                            disabled={isSelf || savingKey === `${user.id}:board`}
                            onClick={() => onChangePermission(user.id, "board", option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div className="permissions-app-row">
                      <span className="permissions-app-label">Subcontractor Directory</span>
                      <div className="permission-segment">
                        {PERMISSION_OPTIONS.map((option) => (
                          <button
                            key={`${user.id}-installer-${option.value}`}
                            type="button"
                            className={`permission-chip ${installerPermission === option.value ? "active" : ""}`}
                            disabled={isSelf || savingKey === `${user.id}:installer`}
                            onClick={() => onChangePermission(user.id, "installer", option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div className="permissions-app-row">
                      <span className="permissions-app-label">Holidays</span>
                      <div className="permission-segment">
                        {PERMISSION_OPTIONS.map((option) => (
                          <button
                            key={`${user.id}-holidays-${option.value}`}
                            type="button"
                            className={`permission-chip ${holidaysPermission === option.value ? "active" : ""}`}
                            disabled={isSelf || savingKey === `${user.id}:holidays`}
                            onClick={() => onChangePermission(user.id, "holidays", option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div className="permissions-app-row">
                      <span className="permissions-app-label">Attendance</span>
                      <div className="permission-segment">
                        {PERMISSION_OPTIONS.map((option) => (
                          <button
                            key={`${user.id}-attendance-${option.value}`}
                            type="button"
                            className={`permission-chip ${attendancePermission === option.value ? "active" : ""}`}
                            disabled={isSelf || savingKey === `${user.id}:attendance`}
                            onClick={() => onChangePermission(user.id, "attendance", option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div className="permissions-app-row">
                      <span className="permissions-app-label">Mileage</span>
                      <div className="permission-segment">
                        {PERMISSION_OPTIONS.map((option) => (
                          <button
                            key={`${user.id}-mileage-${option.value}`}
                            type="button"
                            className={`permission-chip ${mileagePermission === option.value ? "active" : ""}`}
                            disabled={isSelf || savingKey === `${user.id}:mileage`}
                            onClick={() => onChangePermission(user.id, "mileage", option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div className="permissions-app-row">
                      <span className="permissions-app-label">Vinyl Estimator</span>
                      <div className="permission-segment">
                        {PERMISSION_OPTIONS.map((option) => (
                          <button
                            key={`${user.id}-vanEstimator-${option.value}`}
                            type="button"
                            className={`permission-chip ${vanEstimatorPermission === option.value ? "active" : ""}`}
                            disabled={isSelf || savingKey === `${user.id}:vanEstimator`}
                            onClick={() => onChangePermission(user.id, "vanEstimator", option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                    </div>
                  </div>

                  <div className="permissions-attendance-settings">
                    <div className="permissions-attendance-head">
                      <div>
                        <strong>Attendance profile</strong>
                        <p>
                          {exemptFromClocking
                            ? "Removed from the attendance board."
                            : fixedHoursMode
                            ? "Uses contracted hours instead of clockings."
                            : "Uses live clock in / out times."}
                        </p>
                      </div>
                      {savingKey === `${user.id}:attendance-profile` ? (
                        <span className="permissions-saving-pill">Saving...</span>
                      ) : null}
                    </div>
                    <div className="permissions-attendance-toggles">
                      <label className="permissions-toggle">
                        <input
                          type="checkbox"
                          checked={exemptFromClocking}
                          disabled={savingKey === `${user.id}:attendance-profile`}
                          onChange={(event) => {
                            const checked = event.target.checked;
                            const nextMode = checked ? "exempt" : fixedHoursMode ? "fixed" : "required";
                            updateAttendanceDraft(user.id, (current) => ({
                              ...current,
                              mode: nextMode
                            }));
                          }}
                        />
                        <span>Exempt from clocking in / out</span>
                      </label>
                      <label className="permissions-toggle">
                        <input
                          type="checkbox"
                          checked={fixedHoursMode}
                          disabled={exemptFromClocking || savingKey === `${user.id}:attendance-profile`}
                          onChange={(event) => {
                            const checked = event.target.checked;
                            updateAttendanceDraft(user.id, (current) => ({
                              ...current,
                              mode: checked ? "fixed" : "required"
                            }));
                          }}
                        />
                        <span>No clocking in / out required</span>
                      </label>
                    </div>

                    <div className="permissions-hours-grid">
                      {ATTENDANCE_WEEKDAYS.map(([dayKey, dayLabel]) => (
                        <div key={`${user.id}-${dayKey}`} className="permissions-hours-row">
                          <span className="permissions-hours-day">{dayLabel}</span>
                          <input
                            type="text"
                            inputMode="numeric"
                            placeholder="09:00"
                            className="permissions-hours-input"
                            value={contractedHours?.[dayKey]?.in || ""}
                            disabled={Boolean(contractedHours?.[dayKey]?.off) || savingKey === `${user.id}:attendance-profile`}
                            onChange={(event) =>
                              updateAttendanceDraft(user.id, (current) => ({
                                ...current,
                                contractedHours: {
                                  ...current.contractedHours,
                                  [dayKey]: {
                                    ...(current.contractedHours?.[dayKey] || {}),
                                    in: event.target.value
                                  }
                                }
                              }))
                            }
                          />
                          <input
                            type="text"
                            inputMode="numeric"
                            placeholder="17:00"
                            className="permissions-hours-input"
                            value={contractedHours?.[dayKey]?.out || ""}
                            disabled={Boolean(contractedHours?.[dayKey]?.off) || savingKey === `${user.id}:attendance-profile`}
                            onChange={(event) =>
                              updateAttendanceDraft(user.id, (current) => ({
                                ...current,
                                contractedHours: {
                                  ...current.contractedHours,
                                  [dayKey]: {
                                    ...(current.contractedHours?.[dayKey] || {}),
                                    out: event.target.value
                                  }
                                }
                              }))
                            }
                          />
                          <label className="permissions-hours-off">
                            <input
                              type="checkbox"
                              checked={Boolean(contractedHours?.[dayKey]?.off)}
                              disabled={savingKey === `${user.id}:attendance-profile`}
                              onChange={(event) =>
                                updateAttendanceDraft(user.id, (current) => ({
                                  ...current,
                                  contractedHours: {
                                    ...current.contractedHours,
                                    [dayKey]: {
                                      ...(current.contractedHours?.[dayKey] || {}),
                                      off: event.target.checked
                                    }
                                  }
                                }))
                              }
                            />
                            <span>Off</span>
                          </label>
                        </div>
                      ))}
                    </div>

                    <div className="permissions-attendance-actions">
                      <button
                        className="ghost-button"
                        type="button"
                        disabled={!attendanceChanged || savingKey === `${user.id}:attendance-profile`}
                        onClick={() => updateAttendanceDraft(user.id, attendanceProfile)}
                      >
                        Reset
                      </button>
                      <button
                        className="primary-button"
                        type="button"
                        disabled={!attendanceChanged || savingKey === `${user.id}:attendance-profile`}
                        onClick={() => onUpdateAttendanceProfile(user.id, attendanceDraft)}
                      >
                        Save attendance settings
                      </button>
                    </div>
                  </div>
                </div>

              <div className="permissions-user-actions">
                <input
                  type="password"
                  className="permissions-password-input"
                  placeholder={user.hasPassword ? "New password" : "Set password"}
                  value={passwordDrafts[user.id] || ""}
                  onChange={(event) =>
                    setPasswordDrafts((current) => ({
                      ...current,
                      [user.id]: event.target.value
                    }))
                  }
                />
                <button
                  className="ghost-button"
                  type="button"
                  onClick={async () => {
                    await onResetPassword(user.id, passwordDrafts[user.id] || "");
                    setPasswordDrafts((current) => ({ ...current, [user.id]: "" }));
                  }}
                  disabled={!passwordDrafts[user.id]}
                >
                  Update password
                </button>
                <button
                  className="text-button danger"
                  type="button"
                  onClick={() => onDeleteUser(user)}
                  disabled={isSelf}
                >
                  Delete user
                </button>
              </div>
            </article>
          );
        })}
      </div>
    </section>
  );
}

function NotificationsPage({
  currentUser,
  onLogout,
  notifications,
  onOpenNotification,
  onMarkNotificationRead,
  onMarkAllNotificationsRead
}) {
  const [activeFilter, setActiveFilter] = useState("all");
  const filterOptions = [
    { value: "all", label: "All" },
    { value: "unread", label: "Unread" },
    { value: "holiday", label: "Holidays" },
    { value: "board", label: "Jobs" },
    { value: "mileage", label: "Mileage" },
    { value: "message", label: "Messages" }
  ];
  const unreadCount = notifications.filter((entry) => !entry.read).length;
  const filteredNotifications = notifications.filter((notification) => {
    const category = getNotificationCategory(notification).label.toLowerCase();
    if (activeFilter === "all") return true;
    if (activeFilter === "unread") return !notification.read;
    return category === activeFilter;
  });

  return (
    <div className="app-shell notifications-shell">
      <div className="page notifications-page">
        <MainNavBar currentUser={currentUser} active="notifications" onLogout={onLogout} notifications={notifications} />

        <section className="panel notifications-panel">
          <div className="notifications-panel-head">
            <div>
              <h2>Notifications</h2>
              <p>{unreadCount ? `${unreadCount} unread notification${unreadCount === 1 ? "" : "s"}` : "You're all caught up."}</p>
            </div>
            {notifications.length ? (
              <button className="ghost-button" type="button" onClick={onMarkAllNotificationsRead}>
                Mark all read
              </button>
            ) : null}
          </div>

          <div className="notifications-filter-row">
            {filterOptions.map((option) => (
              <button
                key={option.value}
                type="button"
                className={`notification-filter-chip ${activeFilter === option.value ? "active" : ""}`}
                onClick={() => setActiveFilter(option.value)}
              >
                {option.label}
              </button>
            ))}
          </div>

          {filteredNotifications.length ? (
                <div className="notifications-feed">
                  {filteredNotifications.map((notification) => {
                    const category = getNotificationCategory(notification);
                    const CategoryIcon = category.icon;
                    const formattedMessage = formatNotificationMessage(notification.message);
                    const formattedTimestamp = formatNotificationDate(notification.createdAt);
                    return (
                      <article
                        key={notification.id}
                        className={`notification-feed-card ${notification.read ? "read" : "unread"}`}
                      >
                        <button
                          type="button"
                          className="notification-feed-main"
                          onClick={() => onOpenNotification(notification)}
                        >
                          <span className={`notification-feed-icon ${category.className}`}>
                            <CategoryIcon />
                          </span>
                          <div className="notification-feed-copy">
                            <div className="notification-feed-top">
                              <div className="notification-feed-title-row">
                                <strong>{notification.title}</strong>
                                {formattedTimestamp ? (
                                  <time className="notification-feed-time" dateTime={notification.createdAt}>
                                    {formattedTimestamp}
                                  </time>
                                ) : null}
                              </div>
                              <div className="notification-feed-meta-row">
                                <span className={`notification-feed-tag ${category.className}`}>{category.label}</span>
                                {!notification.read ? <span className="notification-feed-status">Unread</span> : null}
                              </div>
                            </div>
                            <p className="notification-feed-message">{formattedMessage || notification.message}</p>
                          </div>
                        </button>
                        <div className="notification-feed-actions">
                          {!notification.read ? (
                            <button
                              type="button"
                              className="text-button"
                              onClick={() => onMarkNotificationRead(notification.id)}
                            >
                              Mark read
                            </button>
                          ) : null}
                        </div>
                      </article>
                    );
                  })}
              </div>
            ) : (
            <div className="notifications-empty">No notifications in this view.</div>
          )}
        </section>
      </div>
    </div>
  );
}

function HostLandingPage({
  currentUser,
  onLogout,
  users,
  savingKey,
  onChangePermission,
  onUpdateAttendanceProfile,
  onCreateUser,
  onResetPassword,
  onDeleteUser,
  notifications
}) {
  const [permissionsOpen, setPermissionsOpen] = useState(false);

  function goTo(path) {
    window.location.assign(path);
  }

  return (
    <div className="app-shell host-landing-shell">
      <div className="page host-landing-page">
        <MainNavBar
          currentUser={currentUser}
          active="home"
          onLogout={onLogout}
          notifications={notifications}
        />

        <section className="panel host-landing-panel">
          <div className="host-landing-actions">
            {canAccessAttendance(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/attendance")}>
                <strong>Attendance</strong>
              </button>
            ) : null}
            {canAccessHolidays(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/holidays")}>
                <strong>Holidays</strong>
              </button>
            ) : null}
            {canAccessMileage(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/mileage")}>
                <strong>Mileage</strong>
              </button>
            ) : null}
            <button
              className={`host-launch-card ${canAccessVanEstimator(currentUser) ? "" : "disabled"}`}
              type="button"
              disabled={!canAccessVanEstimator(currentUser)}
              onClick={() => goTo("/van-estimator")}
              title={canAccessVanEstimator(currentUser) ? "Open Vinyl Estimator" : "Vinyl Estimator is inactive"}
            >
              <strong>Vinyl Estimator</strong>
              {!canAccessVanEstimator(currentUser) ? <span className="host-launch-status">Inactive</span> : null}
            </button>
            {canAccessInstaller(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/installer")}>
                <strong>Subcontractor Database</strong>
              </button>
            ) : null}
            {canAccessBoard(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo(getBoardPathForUser(currentUser))}>
                <strong>Installation Board</strong>
              </button>
            ) : null}
            {currentUser?.canManagePermissions ? (
              <button className="host-launch-card" type="button" onClick={() => setPermissionsOpen(true)}>
                <strong>Manage Permissions</strong>
              </button>
            ) : null}
          </div>
        </section>

      </div>

      {currentUser?.canManagePermissions && permissionsOpen ? (
        <div className="modal-backdrop" onClick={() => setPermissionsOpen(false)}>
          <div className="modal permissions-modal" onClick={(event) => event.stopPropagation()}>
            <div className="modal-head">
              <div>
                <h3>Permissions</h3>
                <p>Set who can use each part of the system.</p>
              </div>
              <button className="icon-button" type="button" onClick={() => setPermissionsOpen(false)}>
                x
              </button>
            </div>
              <PermissionsPanel
                currentUser={currentUser}
                users={users}
                savingKey={savingKey}
                onChangePermission={onChangePermission}
                onUpdateAttendanceProfile={onUpdateAttendanceProfile}
                onCreateUser={onCreateUser}
                onResetPassword={onResetPassword}
                onDeleteUser={onDeleteUser}
              />
          </div>
        </div>
      ) : null}
    </div>
  );
}

function ClientLandingPage({
  currentUser,
  onLogout,
  notifications
}) {
  function goTo(path) {
    window.location.assign(path);
  }

  return (
    <div className="app-shell host-landing-shell client-landing-shell">
      <div className="page host-landing-page">
        <MainNavBar
          currentUser={currentUser}
          active="home"
          onLogout={onLogout}
          notifications={notifications}
        />

        <section className="panel host-landing-panel">
          <div className="host-landing-actions">
            {canAccessAttendance(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/attendance")}>
                <strong>Attendance</strong>
              </button>
            ) : null}
            {canAccessHolidays(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/holidays")}>
                <strong>Holidays</strong>
              </button>
            ) : null}
            {canAccessMileage(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/mileage")}>
                <strong>Mileage</strong>
              </button>
            ) : null}
            <button
              className={`host-launch-card ${canAccessVanEstimator(currentUser) ? "" : "disabled"}`}
              type="button"
              disabled={!canAccessVanEstimator(currentUser)}
              onClick={() => goTo("/van-estimator")}
              title={canAccessVanEstimator(currentUser) ? "Open Vinyl Estimator" : "Vinyl Estimator is inactive"}
            >
              <strong>Vinyl Estimator</strong>
              {!canAccessVanEstimator(currentUser) ? <span className="host-launch-status">Inactive</span> : null}
            </button>
            {canAccessInstaller(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo("/installer")}>
                <strong>Subcontractor Directory</strong>
              </button>
            ) : null}
            {canAccessBoard(currentUser) ? (
              <button className="host-launch-card" type="button" onClick={() => goTo(getBoardPathForUser(currentUser))}>
                <strong>Installation Board</strong>
              </button>
            ) : null}
          </div>
        </section>

      </div>
    </div>
  );
}

function HolidaysPage({
  currentUser,
  onLogout,
  notifications,
  holidays,
  holidayRequests,
  approvedHolidayRequests,
  holidayStaff,
  holidayAllowances,
  holidayEvents,
  holidayRows,
  holidayYearStart,
  holidayYearLabel,
  currentHolidayYearStart,
  setHolidayYearStart,
  holidayRequestOpen,
  setHolidayRequestOpen,
  holidayRequestForm,
  setHolidayRequestForm,
  holidayRequestSaving,
  holidayCancelOpen,
  setHolidayCancelOpen,
  holidayCancelForm,
  setHolidayCancelForm,
  holidayEventOpen,
  setHolidayEventOpen,
  holidayEventForm,
  setHolidayEventForm,
  holidayEventSaving,
  holidayAllowanceSavingKey,
  onChangeHolidayAllowanceDraft,
  onSaveHolidayAllowance,
  onToggleHolidayDate,
  onSubmitHolidayEvent,
  onDeleteHolidayEvent,
  onSubmitHolidayRequest,
  onReviewHolidayRequest,
  onCancelHolidayRequest
}) {
  const canReview = canEditHolidays(currentUser);
  const currentPerson = getHolidayStaffPersonForUser(currentUser);
  const showingFutureYear = holidayYearStart > currentHolidayYearStart;
  const [selectedHolidayPerson, setSelectedHolidayPerson] = useState("");
  const cancellableHolidayRequests = useMemo(() => {
    if (canReview || !currentPerson) return [];
    const personKey = getHolidayStaffIdentityKey(currentPerson);
    const currentUserId = String(currentUser?.id || "");
    return approvedHolidayRequests.filter((request) => {
      const requestStatus = String(request.status || "").trim().toLowerCase();
      const requestAction = String(request.action || "book").trim().toLowerCase();
      const sameUser =
        (currentUserId && String(request.requestedByUserId || "") === currentUserId) ||
        getHolidayStaffIdentityKey(request.person) === personKey;
      return sameUser && requestStatus === "approved" && requestAction === "book";
    });
  }, [canReview, currentPerson, currentUser, approvedHolidayRequests]);
  const visibleHolidayAllowances = useMemo(
    () => {
      const allowedPeople = new Set(
        holidayStaff.map((entry) => getHolidayStaffIdentityKey(entry.person || entry.fullName || entry.name))
      );
      return holidayAllowances
        .map((rawEntry) => getHolidayAllowanceSummary(rawEntry))
        .filter((entry) => allowedPeople.has(getHolidayStaffIdentityKey(entry.person)));
    },
    [holidayAllowances, holidayStaff]
  );
  const activeHolidayFilter = selectedHolidayPerson;
  const filteredHolidayRows = useMemo(
    () =>
      holidayRows.map((month) => ({
        ...month,
        days: month.days.map((day) => ({
          ...day,
          holidays: activeHolidayFilter
            ? day.holidays.filter(
                (holiday) =>
                  getHolidayStaffIdentityKey(holiday.person) === getHolidayStaffIdentityKey(activeHolidayFilter)
              )
            : day.holidays
        }))
      })),
    [holidayRows, activeHolidayFilter]
  );

  return (
    <div className="app-shell holidays-shell">
      <div className="page holidays-page">
        <MainNavBar
          currentUser={currentUser}
          active="holidays"
          onLogout={onLogout}
          notifications={notifications}
        />

        <section className="panel holidays-panel">
          <div className="holidays-toolbar">
            <div>
              <h2>Holiday Calendar {holidayYearLabel}</h2>
            </div>
            <div className="holidays-toolbar-actions">
              {showingFutureYear ? (
                <button className="ghost-button" type="button" onClick={() => setHolidayYearStart(currentHolidayYearStart)}>
                  Current year
                </button>
                ) : null}
                <button className="ghost-button" type="button" onClick={() => setHolidayRequestOpen(true)}>
                  Request holiday
                </button>
                {!canReview ? (
                  <button
                    className="ghost-button"
                    type="button"
                    onClick={() => setHolidayCancelOpen(true)}
                    disabled={!cancellableHolidayRequests.length}
                  >
                    Cancel holiday
                  </button>
                ) : null}
              </div>
            </div>

          {holidayRequests.length ? (
            <section className="holiday-requests-panel">
              <div className="holiday-requests-head">
                <h3>{canReview ? "Pending Requests" : "Your Requests"}</h3>
              </div>
              <div className="holiday-request-list">
                {holidayRequests.map((request) => (
                  <article key={request.id} className={`holiday-request-card status-${request.status || "pending"}`}>
                    <div className="holiday-request-main">
                      <strong>{request.person}</strong>
                      <span>{formatHolidayRequestDateRange(request.startDate, request.endDate)}</span>
                      <span>{request.duration || "Full Day"}</span>
                      <span>{String(request.action || "book").toLowerCase() === "cancel" ? "Cancellation request" : "Holiday request"}</span>
                      {request.notes ? <p>{request.notes}</p> : null}
                    </div>
                    <div className="holiday-request-side">
                      <span className={`holiday-request-status status-${request.status || "pending"}`}>
                        {request.status || "pending"}
                      </span>
                      <small>Requested by {request.requestedByName || request.person}</small>
                      {canReview && request.status === "pending" ? (
                        <div className="holiday-request-actions">
                          <button className="ghost-button" type="button" onClick={() => onReviewHolidayRequest(request.id, "approved")}>
                            Approve
                          </button>
                          <button className="text-button danger" type="button" onClick={() => onReviewHolidayRequest(request.id, "rejected")}>
                            Reject
                          </button>
                        </div>
                      ) : null}
                    </div>
                  </article>
                ))}
              </div>
            </section>
          ) : null}

          <div className="holiday-calendar-wrap">
            <div className="holiday-calendar-tools">
              <div className="holiday-calendar-filter-summary">
                {activeHolidayFilter ? (
                  <>
                    <span className="holiday-filter-label">Showing</span>
                    <button
                      type="button"
                      className="holiday-filter-pill active"
                      onClick={() => setSelectedHolidayPerson("")}
                    >
                      {activeHolidayFilter}
                    </button>
                    <button
                      type="button"
                      className="text-button"
                      onClick={() => setSelectedHolidayPerson("")}
                    >
                      Show everyone
                    </button>
                  </>
                ) : (
                  <span className="holiday-filter-label">
                    {canReview ? "Showing all employees" : "Showing all employees on the calendar"}
                  </span>
                )}
              </div>
            </div>
            <div className="holiday-calendar-grid">
              <div className="holiday-calendar-header month-label-cell">Month</div>
              {Array.from({ length: 31 }, (_, index) => (
                <div key={`header-${index + 1}`} className="holiday-calendar-header day-header-cell">
                  {String(index + 1).padStart(2, "0")}
                </div>
              ))}

              {filteredHolidayRows.map((month) => (
                <div key={month.id} className="holiday-calendar-row">
                  <div className="month-label-cell month-row-label">{month.label}</div>
                  {month.days.map((day) => (
                      (() => {
                        const matchingHoliday = activeHolidayFilter
                          ? day.holidays.find(
                              (holiday) =>
                                getHolidayStaffIdentityKey(holiday.person) ===
                                getHolidayStaffIdentityKey(activeHolidayFilter)
                            )
                          : null;

                      return (
                        <div
                          key={day.key}
                          className={[
                            "holiday-day-cell",
                            !day.inMonth ? "is-empty" : "",
                            day.weekend ? "is-weekend" : "",
                            day.bankHoliday ? "is-bank-holiday" : "",
                            canReview && activeHolidayFilter && day.inMonth && !day.weekend && !day.bankHoliday ? "is-editable" : ""
                          ].join(" ").trim()}
                          title={day.bankHoliday || day.isoDate || ""}
                            onClick={() => {
                              if (!canReview || !day.inMonth) return;
                              if (activeHolidayFilter) {
                                if (day.weekend || day.bankHoliday) return;
                                if (matchingHoliday && isBirthdayHoliday(matchingHoliday)) return;
                                onToggleHolidayDate(day.isoDate, activeHolidayFilter);
                                return;
                              }
                              setHolidayEventForm({
                                id: day.events?.[0]?.id || "",
                                date: day.isoDate,
                                title: day.events?.[0]?.title || ""
                              });
                              setHolidayEventOpen(true);
                            }}
                          >
                            {day.events.map((event) => (
                              <button
                                key={`${day.key}-event-${event.id}`}
                                type="button"
                                className="holiday-day-event"
                                onClick={(clickEvent) => {
                                  clickEvent.stopPropagation();
                                  if (!canReview) return;
                                  setHolidayEventForm({
                                    id: event.id,
                                    date: day.isoDate,
                                    title: event.title || ""
                                  });
                                  setHolidayEventOpen(true);
                                }}
                              >
                                {event.title}
                              </button>
                            ))}
                            {day.holidays.map((holiday) => (
                              <span
                                key={`${day.key}-${holiday.id}`}
                              className={`holiday-day-token ${HOLIDAY_PERSON_COLORS[holiday.person] || "holiday-person-black"} ${isBirthdayHoliday(holiday) ? "holiday-birthday-token" : ""}`}
                            >
                              {getHolidayDisplayToken(holiday.person)}
                              {holiday.duration === "Morning" ? " AM" : holiday.duration === "Afternoon" ? " PM" : ""}
                            </span>
                          ))}
                        </div>
                      );
                    })()
                  ))}
                </div>
              ))}
            </div>
          </div>

          <section className="holiday-breakdown-panel">
            <div className="holiday-requests-head">
                <h3>{canReview ? `Holiday Breakdown ${holidayYearLabel}` : `Your Holiday Breakdown ${holidayYearLabel}`}</h3>
              </div>
            <div className="holiday-breakdown-wrap">
              <table className="holiday-breakdown-table">
                <colgroup>
                  <col className="holiday-col-employee" />
                  <col span="11" className="holiday-col-metric" />
                </colgroup>
                <thead>
                  <tr>
                    <th>Employee</th>
                    <th>Birthday</th>
                    <th>Number Days work pw</th>
                    <th>Standard Entitlement (21 + 8BH)</th>
                    <th>Extra Days (Service)</th>
                    <th>Pro-rata Allowance</th>
                    <th>Allocated for Xmas</th>
                    <th>Allocated Bank Holiday</th>
                    <th>Total Days Booked</th>
                    <th>Days Left to Book</th>
                    <th>Unpaid Days Booked</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleHolidayAllowances.map((entry) => {
                    const isSelected =
                      String(activeHolidayFilter || "").trim().toLowerCase() ===
                      String(entry.person || "").trim().toLowerCase();
                    return (
                    <tr key={`${holidayYearStart}-${entry.person}`}>
                      <td>
                        <button
                          type="button"
                          className={`holiday-person-filter ${isSelected ? "active" : ""}`}
                          onClick={() =>
                            setSelectedHolidayPerson((current) =>
                              current.toLowerCase() === String(entry.person || "").trim().toLowerCase() ? "" : entry.person
                            )
                          }
                        >
                          {entry.fullName}
                        </button>
                      </td>
                      <td>
                        {canReview ? (
                          <input
                            className="holiday-allowance-input holiday-birthday-input"
                            type="date"
                            value={entry.birthDate || ""}
                            disabled={holidayAllowanceSavingKey === `${entry.person}:birthDate`}
                            onChange={(event) => onChangeHolidayAllowanceDraft(entry.person, { birthDate: event.target.value })}
                            onBlur={(event) => onSaveHolidayAllowance(entry.person, { birthDate: event.target.value })}
                          />
                        ) : (
                          <span className="holiday-birthday-label">{formatHolidayBirthday(entry.birthDate)}</span>
                        )}
                      </td>
                      {[
                        ["workDaysPerWeek", entry.workDaysPerWeek],
                        ["standardEntitlement", entry.standardEntitlement],
                        ["extraServiceDays", entry.extraServiceDays]
                      ].map(([field, value]) => (
                        <td key={`${entry.person}-${field}`}>
                          {canReview ? (
                            <input
                              className="holiday-allowance-input"
                              type="number"
                              min="0"
                              step="0.5"
                              value={value}
                              disabled={holidayAllowanceSavingKey === `${entry.person}:${field}`}
                              onChange={(event) => onChangeHolidayAllowanceDraft(entry.person, { [field]: event.target.value })}
                              onBlur={(event) => onSaveHolidayAllowance(entry.person, { [field]: event.target.value })}
                              onKeyDown={(event) => {
                                if (event.key === "Enter") {
                                  event.preventDefault();
                                  event.currentTarget.blur();
                                }
                              }}
                            />
                          ) : (
                            <span>{value}</span>
                          )}
                        </td>
                      ))}
                      <td><strong>{entry.prorataAllowance}</strong></td>
                      {[
                        ["christmasDays", entry.christmasDays],
                        ["bankHolidayDays", entry.bankHolidayDays]
                      ].map(([field, value]) => (
                        <td key={`${entry.person}-${field}`}>
                          {canReview ? (
                            <input
                              className="holiday-allowance-input"
                              type="number"
                              min="0"
                              step="0.5"
                              value={value}
                              disabled={holidayAllowanceSavingKey === `${entry.person}:${field}`}
                              onChange={(event) => onChangeHolidayAllowanceDraft(entry.person, { [field]: event.target.value })}
                              onBlur={(event) => onSaveHolidayAllowance(entry.person, { [field]: event.target.value })}
                              onKeyDown={(event) => {
                                if (event.key === "Enter") {
                                  event.preventDefault();
                                  event.currentTarget.blur();
                                }
                              }}
                            />
                          ) : (
                            <span>{value}</span>
                          )}
                        </td>
                      ))}
                      <td>{entry.bookedDays}</td>
                      <td className={entry.daysLeft < 0 ? "holiday-days-negative" : "holiday-days-positive"}>
                        <strong>{entry.daysLeft}</strong>
                      </td>
                      <td>
                        {canReview ? (
                          <input
                            className="holiday-allowance-input"
                            type="number"
                            min="0"
                            step="0.5"
                            value={entry.unpaidDaysBooked || 0}
                            disabled={holidayAllowanceSavingKey === `${entry.person}:unpaidDaysBooked`}
                            onChange={(event) => onChangeHolidayAllowanceDraft(entry.person, { unpaidDaysBooked: event.target.value })}
                            onBlur={(event) => onSaveHolidayAllowance(entry.person, { unpaidDaysBooked: event.target.value })}
                            onKeyDown={(event) => {
                              if (event.key === "Enter") {
                                event.preventDefault();
                                event.currentTarget.blur();
                              }
                            }}
                          />
                        ) : (
                          <span>{entry.unpaidDaysBooked || 0}</span>
                        )}
                      </td>
                    </tr>
                  )})}
                </tbody>
              </table>
            </div>
          </section>

          <div className="holiday-year-footer">
            <button className="ghost-button" type="button" onClick={() => setHolidayYearStart((current) => current + 1)}>
              {holidayYearStart + 1} Holidays
            </button>
          </div>
        </section>
      </div>

        {holidayRequestOpen ? (
          <div className="modal-backdrop" onClick={() => setHolidayRequestOpen(false)}>
          <div className="modal holiday-request-modal" onClick={(event) => event.stopPropagation()}>
            <div className="modal-head">
              <div>
                <h3>Request holiday</h3>
                <p>Send a request for approval. Approved holidays will show on the board automatically.</p>
              </div>
              <button className="icon-button" type="button" onClick={() => setHolidayRequestOpen(false)}>
                x
              </button>
            </div>

            <form
              className="job-form"
              onSubmit={(event) => {
                event.preventDefault();
                onSubmitHolidayRequest();
              }}
            >
              {canReview ? (
                <label>
                  Employee
                  <select
                    value={holidayRequestForm.person}
                    onChange={(event) => setHolidayRequestForm((current) => ({ ...current, person: event.target.value }))}
                  >
                    <option value="">Select employee</option>
                    {holidayStaff.map((entry) => (
                      <option key={entry.person} value={entry.person}>
                        {entry.code} - {entry.fullName}
                      </option>
                    ))}
                  </select>
                </label>
              ) : (
                <label>
                  Employee
                  <input type="text" value={currentPerson} readOnly />
                </label>
              )}

              <div className="split-fields">
                <label>
                  Start date
                  <input
                    type="date"
                    value={holidayRequestForm.startDate}
                    onChange={(event) =>
                      setHolidayRequestForm((current) => ({
                        ...current,
                        startDate: event.target.value,
                        endDate: current.endDate || event.target.value
                      }))
                    }
                  />
                </label>

                <label>
                  End date
                  <input
                    type="date"
                    value={holidayRequestForm.endDate}
                    onChange={(event) => setHolidayRequestForm((current) => ({ ...current, endDate: event.target.value }))}
                  />
                </label>
              </div>

              <label>
                Duration
                <select
                  value={holidayRequestForm.duration}
                  onChange={(event) => setHolidayRequestForm((current) => ({ ...current, duration: event.target.value }))}
                >
                  <option value="Full Day">Full Day</option>
                  <option value="Morning">Morning</option>
                  <option value="Afternoon">Afternoon</option>
                </select>
              </label>

              <label>
                Notes
                <textarea
                  rows="4"
                  value={holidayRequestForm.notes}
                  onChange={(event) => setHolidayRequestForm((current) => ({ ...current, notes: event.target.value }))}
                />
              </label>

              <div className="form-actions">
                <button className="primary-button" type="submit" disabled={holidayRequestSaving}>
                  {holidayRequestSaving ? "Sending..." : "Send request"}
                </button>
                <button className="ghost-button" type="button" onClick={() => setHolidayRequestOpen(false)}>
                  Cancel
                </button>
              </div>
            </form>
          </div>
          </div>
          ) : null}

        {holidayCancelOpen ? (
          <div className="modal-backdrop" onClick={() => setHolidayCancelOpen(false)}>
            <div className="modal holiday-request-modal" onClick={(event) => event.stopPropagation()}>
              <div className="modal-head">
                <div>
                  <h3>Cancel holiday</h3>
                  <p>Send a cancellation request for approval. Nothing will be removed until an admin approves it.</p>
                </div>
                <button className="icon-button" type="button" onClick={() => setHolidayCancelOpen(false)}>
                  x
                </button>
              </div>

              <form
                className="job-form"
                onSubmit={(event) => {
                  event.preventDefault();
                  onCancelHolidayRequest();
                }}
              >
                <label>
                  Approved holiday
                  <select
                    value={holidayCancelForm.requestId}
                    onChange={(event) => setHolidayCancelForm((current) => ({ ...current, requestId: event.target.value }))}
                  >
                    <option value="">Select approved holiday request</option>
                    {cancellableHolidayRequests.map((holidayRequest) => (
                      <option key={holidayRequest.id} value={holidayRequest.id}>
                        {formatHolidayRequestDateRange(holidayRequest.startDate, holidayRequest.endDate)} - {holidayRequest.duration || "Full Day"}
                      </option>
                    ))}
                  </select>
                </label>

                <label>
                  Notes
                  <textarea
                    rows="4"
                    value={holidayCancelForm.notes}
                    onChange={(event) => setHolidayCancelForm((current) => ({ ...current, notes: event.target.value }))}
                  />
                </label>

                <div className="form-actions">
                  <button className="primary-button" type="submit" disabled={holidayRequestSaving || !cancellableHolidayRequests.length}>
                    {holidayRequestSaving ? "Sending..." : "Send cancellation request"}
                  </button>
                  <button className="ghost-button" type="button" onClick={() => setHolidayCancelOpen(false)}>
                    Close
                  </button>
                </div>
              </form>
            </div>
          </div>
        ) : null}

          {holidayEventOpen ? (
          <div className="modal-backdrop" onClick={() => setHolidayEventOpen(false)}>
            <div className="modal holiday-request-modal" onClick={(event) => event.stopPropagation()}>
              <div className="modal-head">
                <div>
                  <h3>Calendar event</h3>
                  <p>Add something like Christmas shutdown or Summer party. This will not affect holiday allowances.</p>
                </div>
                <button className="icon-button" type="button" onClick={() => setHolidayEventOpen(false)}>
                  x
                </button>
              </div>

              <form
                className="job-form"
                onSubmit={(event) => {
                  event.preventDefault();
                  onSubmitHolidayEvent();
                }}
              >
                <label>
                  Date
                  <input
                    type="date"
                    value={holidayEventForm.date}
                    onChange={(event) => setHolidayEventForm((current) => ({ ...current, date: event.target.value }))}
                  />
                </label>

                <label>
                  Event title
                  <input
                    type="text"
                    value={holidayEventForm.title}
                    placeholder="Christmas shutdown"
                    onChange={(event) => setHolidayEventForm((current) => ({ ...current, title: event.target.value }))}
                  />
                </label>

                <div className="modal-actions">
                  <button className="primary-button" type="submit" disabled={holidayEventSaving}>
                    {holidayEventSaving ? "Saving..." : holidayEventForm.id ? "Update event" : "Add event"}
                  </button>
                  {holidayEventForm.id ? (
                    <button
                      className="text-button danger"
                      type="button"
                      onClick={() => onDeleteHolidayEvent(holidayEventForm.id)}
                    >
                      Delete
                    </button>
                  ) : null}
                  <button className="ghost-button" type="button" onClick={() => setHolidayEventOpen(false)}>
                    Cancel
                  </button>
                </div>
              </form>
            </div>
          </div>
        ) : null}
      </div>
    );
  }

function createMileageLine(overrides = {}) {
  return {
    id: overrides.id || (typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : String(Date.now())),
    date: overrides.date || "",
    from: overrides.from || "",
    to: overrides.to || "",
    note: overrides.note || "",
    miles: overrides.miles ?? ""
  };
}

function MileageUserPage({ currentUser, onLogout, notifications, onRefreshNotifications }) {
  const initialMonth = useMemo(() => {
    const params = new URLSearchParams(typeof window !== "undefined" ? window.location.search : "");
    return params.get("month") || toMonthIdFromIso(getLocalTodayIso());
  }, []);
  const [monthId, setMonthId] = useState(initialMonth);
  const [lines, setLines] = useState([createMileageLine()]);
  const [history, setHistory] = useState([]);
  const [monthLabel, setMonthLabel] = useState("");
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [estimatingLineId, setEstimatingLineId] = useState("");
  const [statusMessage, setStatusMessage] = useState(null);

  const totalMiles = useMemo(
    () => Math.round(lines.reduce((sum, line) => sum + (Number(line.miles) || 0), 0) * 10) / 10,
    [lines]
  );

  function updateLine(lineId, key, value) {
    setLines((current) =>
      current.map((line) => (line.id === lineId ? { ...line, [key]: value } : line))
    );
  }

  function removeLine(lineId) {
    setLines((current) => {
      const next = current.filter((line) => line.id !== lineId);
      return next.length ? next : [createMileageLine()];
    });
  }

  function applyMileagePayload(payload) {
    setMonthLabel(payload.monthLabel || "");
    setHistory(Array.isArray(payload.history) ? payload.history : []);
    setLines([createMileageLine()]);
  }

  async function loadMileage(nextMonthId = monthId) {
    try {
      setLoading(true);
      const response = await fetch(`/api/mileage?month=${encodeURIComponent(nextMonthId)}`);
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not load mileage.");
      applyMileagePayload(payload);
    } catch (error) {
      console.error(error);
      setStatusMessage(createMessage(error.message || "Could not load mileage.", "error"));
    } finally {
      setLoading(false);
    }
  }

  async function estimateLine(line) {
    if (!line.from.trim() || !line.to.trim()) {
      setStatusMessage(createMessage("Enter both From and To before estimating miles.", "error"));
      return;
    }

    try {
      setEstimatingLineId(line.id);
      const response = await fetch("/api/mileage/estimate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ from: line.from, to: line.to })
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not estimate mileage.");
      if (payload.resolved && Number(payload.miles) > 0) {
        updateLine(line.id, "miles", String(payload.miles));
        setStatusMessage(createMessage("Mileage estimate added. You can still adjust it if needed.", "success"));
      } else {
        setStatusMessage(createMessage(payload.message || "Could not estimate that route. Please enter the miles manually.", "error"));
      }
    } catch (error) {
      console.error(error);
      setStatusMessage(createMessage(error.message || "Could not estimate mileage.", "error"));
    } finally {
      setEstimatingLineId("");
    }
  }

  async function submitMileage(event) {
    event.preventDefault();
    const cleanLines = lines
      .map((line) => ({
        id: line.id,
        date: line.date,
        from: line.from.trim(),
        to: line.to.trim(),
        note: line.note.trim(),
        miles: Number(line.miles) || 0
      }))
      .filter((line) => line.date || line.from || line.to || line.note || line.miles);

    if (!cleanLines.length) {
      setStatusMessage(createMessage("Add at least one mileage line before submitting.", "error"));
      return;
    }
    if (cleanLines.some((line) => !line.date || !line.from || !line.to || !line.note || !line.miles)) {
      setStatusMessage(createMessage("Every journey needs Date, From, To, Miles and a note explaining what it was for.", "error"));
      return;
    }

    try {
      setSaving(true);
      const response = await fetch("/api/mileage", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ monthId, lines: cleanLines })
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not submit mileage.");
      applyMileagePayload(payload);
      await onRefreshNotifications?.();
      setStatusMessage(createMessage("Mileage submitted to Matt.", "success"));
    } catch (error) {
      console.error(error);
      setStatusMessage(createMessage(error.message || "Could not submit mileage.", "error"));
    } finally {
      setSaving(false);
    }
  }

  async function deleteMileageJourney(targetMonthId, lineId, targetMonthLabel = "") {
    if (!targetMonthId || !lineId) return;
    if (!window.confirm(`Delete this journey from ${targetMonthLabel || targetMonthId}?`)) return;

    try {
      setDeleting(true);
      const response = await fetch(
        `/api/mileage/${encodeURIComponent(targetMonthId)}/lines/${encodeURIComponent(lineId)}`,
        { method: "DELETE" }
      );
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not delete mileage journey.");
      if (targetMonthId === monthId) {
        applyMileagePayload(payload);
      } else {
        setHistory(Array.isArray(payload.history) ? payload.history : []);
      }
      await onRefreshNotifications?.();
      setStatusMessage(createMessage("Mileage journey deleted.", "success"));
    } catch (error) {
      console.error(error);
      setStatusMessage(createMessage(error.message || "Could not delete mileage journey.", "error"));
    } finally {
      setDeleting(false);
    }
  }

  useEffect(() => {
    loadMileage(monthId);
  }, [monthId]);

  useEffect(() => {
    if (!statusMessage) return undefined;
    const timer = window.setTimeout(() => setStatusMessage(null), 3500);
    return () => window.clearTimeout(timer);
  }, [statusMessage]);

  return (
    <div className="app-shell mileage-shell">
      <div className="page mileage-page">
        <MainNavBar currentUser={currentUser} active="mileage" onLogout={onLogout} notifications={notifications} />

        <section className="panel mileage-panel">
          <div className="mileage-head">
            <div>
              <p className="eyebrow">Mileage</p>
              <h2>{monthLabel || "Mileage claim"}</h2>
              <p>Add journeys, estimate the driving miles, and submit the monthly total to Matt.</p>
            </div>
            <div className="mileage-month-tools">
              <button className="ghost-button" type="button" onClick={() => setMonthId(shiftMonthId(monthId, -1))}>
                Previous month
              </button>
              <input
                type="month"
                value={monthId}
                onChange={(event) => setMonthId(event.target.value || toMonthIdFromIso(getLocalTodayIso()))}
              />
              <button className="ghost-button" type="button" onClick={() => setMonthId(shiftMonthId(monthId, 1))}>
                Next month
              </button>
            </div>
          </div>

          {statusMessage ? <div className={`flash ${statusMessage.tone}`}>{statusMessage.text}</div> : null}

          <form className="mileage-form" onSubmit={submitMileage}>
            <div className="mileage-lines">
              {lines.map((line, index) => (
                <div key={line.id} className="mileage-line">
                  <span className="mileage-line-number">{index + 1}</span>
                  <label className="mileage-date-field">
                    Date
                    <input
                      type="date"
                      required
                      value={line.date}
                      onChange={(event) => updateLine(line.id, "date", event.target.value)}
                    />
                  </label>
                  <label>
                    From
                    <input
                      type="text"
                      required
                      value={line.from}
                      placeholder="Start destination"
                      onChange={(event) => updateLine(line.id, "from", event.target.value)}
                    />
                  </label>
                  <label>
                    To
                    <input
                      type="text"
                      required
                      value={line.to}
                      placeholder="End destination"
                      onChange={(event) => updateLine(line.id, "to", event.target.value)}
                    />
                  </label>
                  <label className="mileage-note-field">
                    Journey note
                    <input
                      type="text"
                      required
                      value={line.note}
                      placeholder="What was this journey for?"
                      onChange={(event) => updateLine(line.id, "note", event.target.value)}
                    />
                  </label>
                  <label className="mileage-miles-field">
                    Miles
                    <input
                      type="number"
                      required
                      min="0"
                      step="0.1"
                      value={line.miles}
                      placeholder="0"
                      onChange={(event) => updateLine(line.id, "miles", event.target.value)}
                    />
                  </label>
                  <div className="mileage-line-actions">
                    <button
                      className="ghost-button"
                      type="button"
                      onClick={() => estimateLine(line)}
                      disabled={estimatingLineId === line.id}
                    >
                      {estimatingLineId === line.id ? "Checking..." : "Suggest"}
                    </button>
                    <button className="text-button danger" type="button" onClick={() => removeLine(line.id)}>
                      Remove
                    </button>
                  </div>
                </div>
              ))}
            </div>

            <div className="mileage-footer">
              <button
                className="ghost-button mileage-add-line"
                type="button"
                onClick={() => setLines((current) => [...current, createMileageLine()])}
              >
                + Add journey
              </button>
              <div className="mileage-total">
                <span>Total miles to submit</span>
                <strong>{totalMiles.toFixed(1)} miles</strong>
              </div>
              <button className="primary-button" type="submit" disabled={saving || loading}>
                {saving ? "Submitting..." : "Submit mileage"}
              </button>
            </div>
          </form>
        </section>

        <section className="panel mileage-history-panel">
          <div className="mileage-history-head">
            <h3>History</h3>
            <p>Your previously submitted mileage totals.</p>
          </div>
          {history.length ? (
            <div className="mileage-history-list">
              {history.map((entry) => (
                <article
                  key={entry.id || entry.monthId}
                  className={`mileage-history-card ${entry.monthId === monthId ? "active" : ""}`}
                >
                  <button type="button" className="mileage-history-open" onClick={() => setMonthId(entry.monthId)}>
                    <span>{entry.monthLabel}</span>
                    <strong>{Number(entry.totalMiles || 0).toFixed(1)} miles</strong>
                    <small>{entry.lineCount} journey{entry.lineCount === 1 ? "" : "s"}</small>
                  </button>
                  <div className="mileage-history-journeys">
                    {(Array.isArray(entry.lines) ? entry.lines : []).map((line) => (
                      <div key={`${entry.monthId}-${line.id}`} className="mileage-history-journey">
                        <div>
                          <strong>{line.note || "No note"}</strong>
                          <small>{formatJobDate(line.date) || "-"}</small>
                          <span>{line.from || "-"} to {line.to || "-"}</span>
                        </div>
                        <span className="mileage-history-miles">{Number(line.miles || 0).toFixed(1)} miles</span>
                        <button
                          type="button"
                          className="text-button danger mileage-history-delete"
                          onClick={() => deleteMileageJourney(line.claimMonthId || entry.monthId, line.id, entry.monthLabel)}
                          disabled={deleting}
                        >
                          Delete
                        </button>
                      </div>
                    ))}
                  </div>
                </article>
              ))}
            </div>
          ) : (
            <div className="notifications-empty">No mileage submitted yet.</div>
          )}
        </section>
      </div>
    </div>
  );
}

function MileageAdminPage({ currentUser, onLogout, notifications }) {
  const initialMonth = useMemo(() => {
    const params = new URLSearchParams(typeof window !== "undefined" ? window.location.search : "");
    return params.get("month") || toMonthIdFromIso(getLocalTodayIso());
  }, []);
  const [monthId, setMonthId] = useState(initialMonth);
  const [overview, setOverview] = useState(null);
  const [loading, setLoading] = useState(false);
  const [statusMessage, setStatusMessage] = useState(null);

  async function loadMileageOverview(nextMonthId = monthId) {
    try {
      setLoading(true);
      const response = await fetch(`/api/mileage/admin?month=${encodeURIComponent(nextMonthId)}`);
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not load mileage overview.");
      setOverview(payload);
    } catch (error) {
      console.error(error);
      setStatusMessage(createMessage(error.message || "Could not load mileage overview.", "error"));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadMileageOverview(monthId);
  }, [monthId]);

  useEffect(() => {
    if (!statusMessage) return undefined;
    const timer = window.setTimeout(() => setStatusMessage(null), 3500);
    return () => window.clearTimeout(timer);
  }, [statusMessage]);

  const users = Array.isArray(overview?.users) ? overview.users : [];
  const monthLabel = overview?.monthLabel || "Mileage overview";

  return (
    <div className="app-shell mileage-shell">
      <div className="page mileage-page mileage-admin-page">
        <MainNavBar currentUser={currentUser} active="mileage" onLogout={onLogout} notifications={notifications} />

        <section className="panel mileage-panel mileage-admin-panel">
          <div className="mileage-head">
            <div>
              <p className="eyebrow">Mileage</p>
              <h2>{monthLabel}</h2>
              <p>Admin overview of submitted journeys by each mileage user.</p>
            </div>
            <div className="mileage-month-tools">
              <button className="ghost-button" type="button" onClick={() => setMonthId(shiftMonthId(monthId, -1))}>
                Previous month
              </button>
              <input
                type="month"
                value={monthId}
                onChange={(event) => setMonthId(event.target.value || toMonthIdFromIso(getLocalTodayIso()))}
              />
              <button className="ghost-button" type="button" onClick={() => setMonthId(shiftMonthId(monthId, 1))}>
                Next month
              </button>
            </div>
          </div>

          {statusMessage ? <div className={`flash ${statusMessage.tone}`}>{statusMessage.text}</div> : null}

          <div className="mileage-admin-summary">
            <div>
              <span>Total miles</span>
              <strong>{Number(overview?.totalMiles || 0).toFixed(1)}</strong>
            </div>
            <div>
              <span>Journeys</span>
              <strong>{Number(overview?.lineCount || 0)}</strong>
            </div>
            <div>
              <span>Submitted</span>
              <strong>{Number(overview?.submittedUserCount || 0)} / {Number(overview?.userCount || 0)}</strong>
            </div>
          </div>

          {loading && !overview ? (
            <div className="notifications-empty">Loading mileage overview...</div>
          ) : (
            <div className="mileage-admin-users">
              {users.length ? (
                users.map((user) => (
                  <article key={user.userId || user.userName} className={`mileage-admin-user-card ${user.lineCount ? "" : "empty"}`}>
                    <div className="mileage-admin-user-head">
                      <div>
                        <h3>{user.userName}</h3>
                        <p>{user.lineCount} journey{user.lineCount === 1 ? "" : "s"} submitted</p>
                      </div>
                      <strong>{Number(user.totalMiles || 0).toFixed(1)} miles</strong>
                    </div>

                    {Array.isArray(user.journeys) && user.journeys.length ? (
                      <div className="mileage-admin-journeys">
                        {user.journeys.map((journey) => (
                          <div key={`${user.userId}-${journey.claimMonthId}-${journey.id}`} className="mileage-admin-journey">
                            <div className="mileage-admin-journey-main">
                              <strong>{journey.note || "No note"}</strong>
                              <span>{journey.from || "-"} to {journey.to || "-"}</span>
                            </div>
                            <time>{formatJobDate(journey.date) || "-"}</time>
                            <strong>{Number(journey.miles || 0).toFixed(1)} miles</strong>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="notifications-empty compact">No mileage submitted for this month.</div>
                    )}
                  </article>
                ))
              ) : (
                <div className="notifications-empty">No mileage users found.</div>
              )}
            </div>
          )}
        </section>
      </div>
    </div>
  );
}

function MileagePage(props) {
  return canEditMileage(props.currentUser) ? <MileageAdminPage {...props} /> : <MileageUserPage {...props} />;
}

function AttendancePage({
  currentUser,
  onLogout,
  notifications,
  attendanceData,
  loading,
  attendanceMonthId,
  setAttendanceMonthId,
  attendanceSavingKey,
  attendanceNoteSavingKey,
  attendanceFocusDate,
  onSaveAttendanceEntry,
  onSubmitAttendanceExplanation
}) {
  const adminMode = canEditAttendance(currentUser);
  const [drafts, setDrafts] = useState({});
  const [noteForm, setNoteForm] = useState(EMPTY_ATTENDANCE_NOTE_FORM);

  useEffect(() => {
    setDrafts({});
  }, [attendanceData?.monthId]);

  useEffect(() => {
    if (!attendanceFocusDate) return;
    setNoteForm((current) =>
      current.date === attendanceFocusDate
        ? current
        : {
            date: attendanceFocusDate,
            note: ""
          }
    );
  }, [attendanceFocusDate]);

  const rows = Array.isArray(attendanceData?.rows) ? attendanceData.rows : [];
  const staff = Array.isArray(attendanceData?.staff) ? attendanceData.staff : [];
  const missingEntries = Array.isArray(attendanceData?.missingEntries) ? attendanceData.missingEntries : [];
  const attendanceMonthLabel = attendanceData?.monthLabel || "Attendance";
  const focusedMissingEntry =
    missingEntries.find((entry) => entry.isoDate === noteForm.date) || missingEntries[0] || null;

  useEffect(() => {
    if (adminMode) return;
    if (noteForm.date) return;
    if (!missingEntries.length) return;
    setNoteForm({
      date: missingEntries[0].isoDate,
      note: missingEntries[0].employeeNote || ""
    });
  }, [adminMode, missingEntries, noteForm.date]);

  function getDraftValue(person, date, field, fallback = "") {
    const key = `${person}:${date}`;
    return drafts[key]?.[field] ?? fallback;
  }

  function setDraftValue(person, date, updates) {
    const key = `${person}:${date}`;
    setDrafts((current) => ({
      ...current,
      [key]: {
        ...(current[key] || {}),
        ...updates
      }
    }));
  }

  async function handleAttendanceBlur(cell) {
    const person = cell.person;
    const date = cell.isoDate;
    const key = `${person}:${date}`;
    const draft = drafts[key] || {};
    const clockIn = draft.clockIn ?? cell.clockIn ?? "";
    const clockOut = draft.clockOut ?? cell.clockOut ?? "";
    const adminNote = draft.adminNote ?? cell.adminNote ?? "";
    await onSaveAttendanceEntry({ person, date, clockIn, clockOut, adminNote });
  }

  async function submitEmployeeNote(event) {
    event.preventDefault();
    if (!noteForm.date || !noteForm.note.trim()) return;
    await onSubmitAttendanceExplanation({
      date: noteForm.date,
      note: noteForm.note.trim()
    });
    setNoteForm(EMPTY_ATTENDANCE_NOTE_FORM);
  }

  return (
    <div className="app-shell holidays-shell">
      <div className="page holidays-page attendance-page">
        <MainNavBar
          currentUser={currentUser}
          active="attendance"
          onLogout={onLogout}
          notifications={notifications}
        />

        <section className="panel holidays-panel attendance-panel">
          <div className="holidays-toolbar attendance-toolbar">
            <div>
              <h2>Attendance</h2>
              <p>{attendanceMonthLabel}</p>
            </div>
            <div className="holidays-toolbar-actions">
              <button className="ghost-button" type="button" onClick={() => setAttendanceMonthId(shiftMonthId(attendanceMonthId, -1))}>
                Previous month
              </button>
              <button className="ghost-button" type="button" onClick={() => setAttendanceMonthId(toMonthIdFromIso(getLocalTodayIso()))}>
                Current month
              </button>
              <button className="ghost-button" type="button" onClick={() => setAttendanceMonthId(shiftMonthId(attendanceMonthId, 1))}>
                Next month
              </button>
            </div>
          </div>

          {loading ? <div className="board-loading">Loading attendance...</div> : null}

          {!loading && adminMode ? (
            <div className="attendance-grid-wrap">
              <table className="attendance-grid-table">
                <thead>
                  <tr>
                    <th className="attendance-date-head" rowSpan={2}>Date</th>
                    {staff.map((person) => (
                      <th
                        key={`staff-${person.person}`}
                        className="attendance-staff-head"
                        colSpan={2}
                        title={person.fullName || person.person}
                      >
                        <span>{person.code || person.fullName || person.person}</span>
                        <small>{person.fullName || person.person}</small>
                      </th>
                    ))}
                  </tr>
                  <tr>
                    {staff.map((person) => (
                      <>
                        <th key={`${person.person}-in`} className="attendance-sub-head">In</th>
                        <th key={`${person.person}-out`} className="attendance-sub-head">Out</th>
                      </>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row) => (
                    <tr key={row.isoDate} className={row.isToday ? "attendance-row-today" : ""}>
                      <th className="attendance-date-cell">
                        <span>{row.weekdayLabel}</span>
                        <strong>{row.dateLabel}</strong>
                        {row.isToday ? <em>Today</em> : null}
                      </th>
                      {row.cells.map((cell) => {
                        const missingClass = cell.hasMissingClock ? "attendance-cell-missing" : "";
                        const displayClass = getAttendanceDisplayClass(cell);
                        if (cell.displayLabel) {
                          return (
                            <td
                              key={`${row.isoDate}-${cell.person}`}
                              className={`attendance-merged-cell ${displayClass} ${missingClass}`.trim()}
                              colSpan={2}
                            >
                              <span className={cell.isHoliday ? "attendance-merged-holiday" : ""}>{cell.displayLabel}</span>
                            </td>
                          );
                        }
                          return (
                            <>
                              <td key={`${row.isoDate}-${cell.person}-in`} className={`attendance-value-cell ${missingClass}`}>
                                <input
                                  className="attendance-time-input"
                                  value={getDraftValue(cell.person, row.isoDate, "clockIn", cell.clockIn)}
                                  placeholder="--:--"
                                  onChange={(event) => setDraftValue(cell.person, row.isoDate, { clockIn: event.target.value })}
                                  onBlur={() => handleAttendanceBlur(cell)}
                                />
                                {cell.halfDayHolidayLabel ? (
                                  <span className="attendance-half-day-chip">{cell.halfDayHolidayLabel}</span>
                                ) : null}
                              </td>
                              <td key={`${row.isoDate}-${cell.person}-out`} className={`attendance-value-cell ${missingClass}`}>
                                <input
                                  className="attendance-time-input"
                                  value={getDraftValue(cell.person, row.isoDate, "clockOut", cell.clockOut)}
                                placeholder="--:--"
                                onChange={(event) => setDraftValue(cell.person, row.isoDate, { clockOut: event.target.value })}
                                onBlur={() => handleAttendanceBlur(cell)}
                              />
                            </td>
                          </>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : null}

          {!loading && !adminMode ? (
            <div className="attendance-self-service">
              <section className="attendance-self-panel">
                <h3>Your times</h3>
                <div className="attendance-self-list">
                  {rows.map((row) => {
                    const cell = row.cells[0];
                    if (!cell) return null;
                    return (
                      <article
                        key={row.isoDate}
                        className={`attendance-self-card ${row.isToday ? "is-today" : ""} ${cell.hasMissingClock ? "is-missing" : ""} ${noteForm.date === row.isoDate ? "is-focused" : ""}`}
                      >
                        <div className="attendance-self-card-head">
                          <div>
                            <strong>{row.dateLabel}</strong>
                            <span>{row.weekdayLabel}</span>
                          </div>
                          {cell.displayLabel ? <span className="attendance-self-status">{cell.displayLabel}</span> : null}
                           {!cell.displayLabel && cell.hasMissingClock ? <span className="attendance-self-status missing">Missing clocking</span> : null}
                           {!cell.displayLabel && cell.halfDayHolidayLabel ? (
                             <span className="attendance-self-status holiday">{cell.halfDayHolidayLabel}</span>
                           ) : null}
                          </div>
                        {cell.displayLabel ? null : (
                          <div className="attendance-self-times">
                            <span>In: <strong>{cell.clockIn || "--:--"}</strong></span>
                            <span>Out: <strong>{cell.clockOut || "--:--"}</strong></span>
                          </div>
                        )}
                        {!cell.displayLabel && cell.canExplain ? (
                          <button
                            type="button"
                            className="ghost-button attendance-note-button"
                            onClick={() => setNoteForm({ date: row.isoDate, note: cell.employeeNote || "" })}
                          >
                            {cell.employeeNote ? "Update note" : "Add note"}
                          </button>
                        ) : null}
                        {cell.employeeNote ? <p className="attendance-self-note">{cell.employeeNote}</p> : null}
                      </article>
                    );
                  })}
                </div>
              </section>

              <section className="attendance-self-panel">
                <h3>Missing clockings</h3>
                {missingEntries.length ? (
                  <>
                    <div className="attendance-missing-list">
                      {missingEntries.map((entry) => (
                        <button
                          key={`${entry.person}-${entry.isoDate}`}
                          type="button"
                          className={`attendance-missing-chip ${noteForm.date === entry.isoDate ? "active" : ""}`}
                          onClick={() => setNoteForm({ date: entry.isoDate, note: entry.employeeNote || "" })}
                        >
                          {entry.dateLabel}: {entry.clockIn ? "Missing out" : entry.clockOut ? "Missing in" : "Missing in/out"}
                        </button>
                      ))}
                    </div>
                    <form className="attendance-note-form" onSubmit={submitEmployeeNote}>
                      <label>
                        <span>Date</span>
                        <input type="text" value={focusedMissingEntry?.dateLabel || ""} readOnly />
                      </label>
                      <label>
                        <span>Explanation</span>
                        <textarea
                          rows={4}
                          value={noteForm.note}
                          onChange={(event) => setNoteForm((current) => ({ ...current, note: event.target.value }))}
                          placeholder="Explain the missing clocking so admin can check it."
                        />
                      </label>
                      <div className="attendance-note-actions">
                        <button
                          className="primary-button"
                          type="submit"
                          disabled={!noteForm.date || !noteForm.note.trim() || attendanceNoteSavingKey === noteForm.date}
                        >
                          {attendanceNoteSavingKey === noteForm.date ? "Saving..." : "Send note"}
                        </button>
                      </div>
                    </form>
                  </>
                ) : (
                  <div className="notifications-empty">No missing clockings for this month.</div>
                )}
              </section>
            </div>
          ) : null}
        </section>
      </div>
    </div>
  );
}

function VinylEstimatorPage({ currentUser, onLogout, notifications }) {
  const [svgMarkup, setSvgMarkup] = useState("");
  const [svgError, setSvgError] = useState("");
  const [shapes, setShapes] = useState([]);
  const [drawMode, setDrawMode] = useState("rectangle");
  const [drawingRect, setDrawingRect] = useState(null);
  const [drawStart, setDrawStart] = useState(null);
  const [polygonPoints, setPolygonPoints] = useState([]);
  const [polygonPreviewPoint, setPolygonPreviewPoint] = useState(null);
  const [editDrag, setEditDrag] = useState(null);
  const inlineSvgRef = useRef(null);
  const overlaySvgRef = useRef(null);
  const wrapLinesRef = useRef([]);

  useEffect(() => {
    let active = true;
    fetch(VAN_ESTIMATOR_TEMPLATE.src)
      .then((response) => {
        if (!response.ok) throw new Error("Could not load the van SVG.");
        return response.text();
      })
      .then((text) => {
        if (!active) return;
        setSvgMarkup(text);
        setSvgError("");
      })
      .catch((error) => {
        console.error(error);
        if (active) setSvgError(error.message || "Could not load the van SVG.");
      });
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (!svgMarkup || !inlineSvgRef.current) return;
    const svg = inlineSvgRef.current.querySelector("svg");
    if (svg) {
      svg.setAttribute("width", "100%");
      svg.removeAttribute("height");
      svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
      svg.classList.add("van-template-svg");
    }

    const wrapLayer = inlineSvgRef.current.querySelector("#Wrap_Film_Lines");
    if (!wrapLayer) {
      wrapLinesRef.current = [];
      return;
    }

    wrapLinesRef.current = Array.from(wrapLayer.querySelectorAll("path,line,polyline,polygon"))
      .filter((element) => {
        const styles = window.getComputedStyle(element);
        const fill = styles.fill || element.getAttribute("fill") || "";
        const stroke = styles.stroke || element.getAttribute("stroke") || "";
        return fill === "none" && stroke !== "none";
      })
      .map((element) => {
        try {
          const box = element.getBBox();
          const length = typeof element.getTotalLength === "function" ? element.getTotalLength() : 0;
          const sampleCount = Math.max(2, Math.ceil(length / 8));
          const points =
            length > 0
              ? Array.from({ length: sampleCount + 1 }, (_, index) => {
                  const point = element.getPointAtLength((length * index) / sampleCount);
                  return { x: point.x, y: point.y };
                })
              : [];
          return {
            box: {
              x: box.x,
              y: box.y,
              width: box.width,
              height: box.height
            },
            points
          };
        } catch (error) {
          return null;
        }
      })
      .filter((line) => line?.points?.length);
  }, [svgMarkup]);

  function getPointerPoint(event) {
    const svg = overlaySvgRef.current;
    if (!svg) return null;
    const point = svg.createSVGPoint();
    point.x = event.clientX;
    point.y = event.clientY;
    const matrix = svg.getScreenCTM();
    if (!matrix) return null;
    return point.matrixTransform(matrix.inverse());
  }

  function normalizeRect(start, end) {
    const x = Math.min(start.x, end.x);
    const y = Math.min(start.y, end.y);
    const width = Math.abs(end.x - start.x);
    const height = Math.abs(end.y - start.y);
    return { x, y, width, height };
  }

  function getPolygonBounds(points) {
    const xValues = points.map((point) => point.x);
    const yValues = points.map((point) => point.y);
    const x = Math.min(...xValues);
    const y = Math.min(...yValues);
    return {
      x,
      y,
      width: Math.max(...xValues) - x,
      height: Math.max(...yValues) - y
    };
  }

  function pointsToSvg(points) {
    return points.map((point) => `${point.x},${point.y}`).join(" ");
  }

  function getRectanglePoints(rect) {
    return [
      { corner: "top-left", x: rect.x, y: rect.y },
      { corner: "top-right", x: rect.x + rect.width, y: rect.y },
      { corner: "bottom-right", x: rect.x + rect.width, y: rect.y + rect.height },
      { corner: "bottom-left", x: rect.x, y: rect.y + rect.height }
    ];
  }

  function rectsIntersect(left, right) {
    return (
      left.x < right.x + right.width &&
      left.x + left.width > right.x &&
      left.y < right.y + right.height &&
      left.y + left.height > right.y
    );
  }

  function isPointInPolygon(point, polygon) {
    let inside = false;
    for (let index = 0, previousIndex = polygon.length - 1; index < polygon.length; previousIndex = index++) {
      const current = polygon[index];
      const previous = polygon[previousIndex];
      const crossesY = current.y > point.y !== previous.y > point.y;
      if (!crossesY) continue;
      const crossingX = ((previous.x - current.x) * (point.y - current.y)) / (previous.y - current.y) + current.x;
      if (point.x < crossingX) inside = !inside;
    }
    return inside;
  }

  function isWrapFilmArea(area) {
    const detectionPadding = 2;
    const rect = area.bounds || area;
    const expandedRect = {
      x: rect.x - detectionPadding,
      y: rect.y - detectionPadding,
      width: rect.width + detectionPadding * 2,
      height: rect.height + detectionPadding * 2
    };

    return wrapLinesRef.current.some((line) => {
      if (!rectsIntersect(expandedRect, line.box)) return false;
      return line.points.some((point) => {
        if (
          point.x < expandedRect.x ||
          point.x > expandedRect.x + expandedRect.width ||
          point.y < expandedRect.y ||
          point.y > expandedRect.y + expandedRect.height
        ) {
          return false;
        }
        return area.points?.length ? isPointInPolygon(point, area.points) : true;
      });
    });
  }

  function getVehicleZoneMetadata(area) {
    const wrapRequired = isWrapFilmArea(area);
    if (wrapRequired) {
      return {
        material_type: "wrap_film",
        surface_type: "curved",
        complexity_factor: 1.25,
        install_group: "wrap_required"
      };
    }

    return {
      material_type: "standard_vinyl",
      surface_type: "flat",
      complexity_factor: 1,
      install_group: "standard_panel"
    };
  }

  function getWrapHoursPerM2(zoneMetadata, coverage) {
    const isPartialWrap = coverage < 0.7;
    if (zoneMetadata?.complexity_factor >= 1.4) {
      return isPartialWrap
        ? VEHICLE_GRAPHICS_PRICING.partialWrapComplexHoursPerM2
        : VEHICLE_GRAPHICS_PRICING.wrapComplexHoursPerM2;
    }
    if (zoneMetadata?.surface_type === "curved") {
      return isPartialWrap
        ? VEHICLE_GRAPHICS_PRICING.partialWrapCurvedHoursPerM2
        : VEHICLE_GRAPHICS_PRICING.wrapCurvedHoursPerM2;
    }
    return isPartialWrap
      ? VEHICLE_GRAPHICS_PRICING.partialWrapFlatHoursPerM2
      : VEHICLE_GRAPHICS_PRICING.wrapFlatHoursPerM2;
  }

  function clampNumber(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }

  function clearTextSelection() {
    if (typeof window !== "undefined") window.getSelection?.()?.removeAllRanges();
  }

  function getRectAreaM2(rect) {
    const widthMm = rect.width * VAN_ESTIMATOR_TEMPLATE.scaleFactor;
    const heightMm = rect.height * VAN_ESTIMATOR_TEMPLATE.scaleFactor;
    return (widthMm / 1000) * (heightMm / 1000);
  }

  function getPolygonAreaM2(points) {
    const areaUnits =
      Math.abs(
        points.reduce((sum, point, index) => {
          const nextPoint = points[(index + 1) % points.length];
          return sum + point.x * nextPoint.y - nextPoint.x * point.y;
        }, 0)
      ) / 2;
    const scaleFactor = VAN_ESTIMATOR_TEMPLATE.scaleFactor;
    return (areaUnits * scaleFactor * scaleFactor) / 1000000;
  }

  function refreshShapeMetrics(shape) {
    if (shape.type === "polygon") {
      const bounds = getPolygonBounds(shape.points);
      const zoneMetadata = getVehicleZoneMetadata({ points: shape.points, bounds });
      return {
        ...shape,
        bounds,
        width: bounds.width,
        height: bounds.height,
        areaM2: getPolygonAreaM2(shape.points),
        zoneMetadata,
        isWrapFilm: zoneMetadata.material_type === "wrap_film"
      };
    }

    const rect = {
      x: shape.x,
      y: shape.y,
      width: shape.width,
      height: shape.height
    };
    const zoneMetadata = getVehicleZoneMetadata(rect);

    return {
      ...shape,
      bounds: rect,
      areaM2: getRectAreaM2(rect),
      zoneMetadata,
      isWrapFilm: zoneMetadata.material_type === "wrap_film"
    };
  }

  function updateShapeCorner(shape, dragState, point) {
    if (shape.type === "polygon") {
      const points = shape.points.map((entry, index) => (index === dragState.pointIndex ? point : entry));
      return refreshShapeMetrics({ ...shape, points });
    }

    const anchor = dragState.anchor;
    const rect = normalizeRect(anchor, point);
    return refreshShapeMetrics({
      ...shape,
      ...rect
    });
  }

  function finishPolygon() {
    if (polygonPoints.length < 3) return;
    const bounds = getPolygonBounds(polygonPoints);
    const areaM2 = getPolygonAreaM2(polygonPoints);
    if (areaM2 <= 0.001) return;
    const zoneMetadata = getVehicleZoneMetadata({ points: polygonPoints, bounds });

    setShapes((current) => [
      ...current,
      {
        id: `vinyl-poly-${Date.now()}-${Math.round(bounds.x)}-${Math.round(bounds.y)}`,
        type: "polygon",
        points: polygonPoints,
        bounds,
        width: bounds.width,
        height: bounds.height,
        zoneMetadata,
        isWrapFilm: zoneMetadata.material_type === "wrap_film",
        areaM2
      }
    ]);
    setPolygonPoints([]);
    setPolygonPreviewPoint(null);
  }

  function startDrawing(event) {
    if (event.button !== 0) return;
    if (editDrag) return;
    clearTextSelection();
    const point = getPointerPoint(event);
    if (!point) return;
    if (drawMode === "polygon") {
      setPolygonPoints((current) => [...current, point]);
      setPolygonPreviewPoint(null);
      return;
    }
    event.currentTarget.setPointerCapture(event.pointerId);
    setDrawStart(point);
    setDrawingRect({ x: point.x, y: point.y, width: 0, height: 0 });
  }

  function updateDrawing(event) {
    const point = getPointerPoint(event);
    if (!point) return;
    if (editDrag) {
      setShapes((current) =>
        current.map((shape) => (shape.id === editDrag.shapeId ? updateShapeCorner(shape, editDrag, point) : shape))
      );
      return;
    }
    if (drawMode === "polygon") {
      if (polygonPoints.length) setPolygonPreviewPoint(point);
      return;
    }
    if (!drawStart) return;
    setDrawingRect(normalizeRect(drawStart, point));
  }

  function finishDrawing(event) {
    if (editDrag) {
      try {
        event.currentTarget.releasePointerCapture(event.pointerId);
      } catch (error) {
        // Pointer capture may already be released if the pointer leaves the browser chrome.
      }
      setEditDrag(null);
      return;
    }
    if (drawMode === "polygon") return;
    if (!drawStart || !drawingRect) return;
    try {
      event.currentTarget.releasePointerCapture(event.pointerId);
    } catch (error) {
      // Pointer capture may already be released if the pointer leaves the browser chrome.
    }

    const rect = {
      ...drawingRect,
      id: `vinyl-${Date.now()}-${Math.round(drawingRect.x)}-${Math.round(drawingRect.y)}`
    };
    setDrawStart(null);
    setDrawingRect(null);
    if (rect.width < 4 || rect.height < 4) return;

    const zoneMetadata = getVehicleZoneMetadata(rect);
    const areaM2 = getRectAreaM2(rect);
    setShapes((current) => [
      ...current,
      {
        ...rect,
        type: "rectangle",
        bounds: rect,
        zoneMetadata,
        isWrapFilm: zoneMetadata.material_type === "wrap_film",
        areaM2
      }
    ]);
  }

  function startShapeCornerDrag(event, shape, cornerOrPointIndex) {
    event.stopPropagation();
    event.preventDefault();
    if (event.button !== 0) return;
    clearTextSelection();
    try {
      event.currentTarget.setPointerCapture(event.pointerId);
    } catch (error) {
      // Some browsers may not allow capture on SVG children; the overlay still receives moves.
    }

    if (shape.type === "polygon") {
      setEditDrag({ shapeId: shape.id, pointIndex: cornerOrPointIndex });
      return;
    }

    const rect = shape.bounds || shape;
    const anchors = {
      "top-left": { x: rect.x + rect.width, y: rect.y + rect.height },
      "top-right": { x: rect.x, y: rect.y + rect.height },
      "bottom-right": { x: rect.x, y: rect.y },
      "bottom-left": { x: rect.x + rect.width, y: rect.y }
    };
    setEditDrag({
      shapeId: shape.id,
      corner: cornerOrPointIndex,
      anchor: anchors[cornerOrPointIndex]
    });
  }

  function deleteShape(event, shapeId) {
    event.stopPropagation();
    event.preventDefault();
    clearTextSelection();
    setShapes((current) => current.filter((shape) => shape.id !== shapeId));
  }

  const totals = useMemo(() => {
    const classifiedShapes = shapes.map((shape) => {
      const zoneMetadata =
        shape.zoneMetadata ||
        (shape.isWrapFilm
          ? {
              material_type: "wrap_film",
              surface_type: "curved",
              complexity_factor: 1.25,
              install_group: "wrap_required"
            }
          : {
              material_type: "standard_vinyl",
              surface_type: "flat",
              complexity_factor: 1,
              install_group: "standard_panel"
            });
      return { ...shape, zoneMetadata };
    });
    const standardShapes = classifiedShapes.filter((shape) => shape.zoneMetadata.material_type !== "wrap_film");
    const wrapShapes = classifiedShapes.filter((shape) => shape.zoneMetadata.material_type === "wrap_film");
    const standardArea = standardShapes.reduce((sum, shape) => sum + shape.areaM2, 0);
    const wrapArea = wrapShapes.reduce((sum, shape) => sum + shape.areaM2, 0);
    const flatArea = classifiedShapes
      .filter((shape) => shape.zoneMetadata.surface_type === "flat")
      .reduce((sum, shape) => sum + shape.areaM2, 0);
    const curvedArea = classifiedShapes
      .filter((shape) => shape.zoneMetadata.surface_type === "curved")
      .reduce((sum, shape) => sum + shape.areaM2, 0);
    const totalArea = standardArea + wrapArea;
    const vehicleArea =
      (VAN_ESTIMATOR_TEMPLATE.viewBox.width *
        VAN_ESTIMATOR_TEMPLATE.viewBox.height *
        VAN_ESTIMATOR_TEMPLATE.scaleFactor *
        VAN_ESTIMATOR_TEMPLATE.scaleFactor) /
      1000000;
    const coverage = vehicleArea > 0 ? totalArea / vehicleArea : 0;
    if (totalArea <= 0) {
      return {
        standardArea,
        wrapArea,
        flatArea,
        curvedArea,
        totalArea,
        materialSell: 0,
        labourHours: 0,
        labourSell: 0,
        basePrice: 0,
        vehicleArea,
        coverage: 0,
        anchor: 0,
        estimate: 0
      };
    }
    const materialSell = classifiedShapes.reduce((sum, shape) => {
      const material = VEHICLE_ZONE_MATERIALS[shape.zoneMetadata.material_type] || VEHICLE_ZONE_MATERIALS.standard_vinyl;
      return sum + shape.areaM2 * material.rate;
    }, 0);
    const standardComplexity = Math.max(1, ...standardShapes.map((shape) => shape.zoneMetadata.complexity_factor || 1));
    const standardLabourHours =
      standardArea <= 0
        ? 0
        : coverage < 0.15
          ? clampNumber(
              standardArea * VEHICLE_GRAPHICS_PRICING.standardSmallHoursPerM2 * standardComplexity,
              VEHICLE_GRAPHICS_PRICING.standardSmallMinHours,
              VEHICLE_GRAPHICS_PRICING.standardSmallMaxHours
            )
          : standardArea * VEHICLE_GRAPHICS_PRICING.standardLargeHoursPerM2 * standardComplexity;
    const wrapLabourHours = wrapShapes.reduce(
      (sum, shape) => sum + shape.areaM2 * getWrapHoursPerM2(shape.zoneMetadata, coverage),
      0
    );
    const labourHours = standardLabourHours + wrapLabourHours;
    const labourSell = labourHours * VEHICLE_GRAPHICS_PRICING.labourSellRate;
    const basePrice = materialSell + labourSell;
    const anchor = coverage < 0.15 ? 400 : coverage < 0.4 ? 900 : coverage < 0.7 ? 1500 : 2500;
    let estimate = 0.65 * basePrice + 0.35 * anchor;
    estimate = Math.max(estimate, VEHICLE_GRAPHICS_PRICING.minPrice);
    if (wrapArea > 0) {
      estimate = Math.max(estimate, VEHICLE_GRAPHICS_PRICING.minAnyWrapPrice);
    }
    if (wrapArea > 0 && coverage < 0.7 && wrapArea >= standardArea) {
      estimate = Math.max(estimate, VEHICLE_GRAPHICS_PRICING.minWrapLedPartialPrice);
    }
    if (wrapArea > 0 && coverage >= 0.7) {
      estimate = Math.max(estimate, VEHICLE_GRAPHICS_PRICING.minFullWrapPrice);
    }
    estimate = Math.round(estimate / 50) * 50;

    return {
      standardArea,
      wrapArea,
      flatArea,
      curvedArea,
      totalArea,
      materialSell,
      labourHours,
      labourSell,
      basePrice,
      vehicleArea,
      coverage,
      anchor,
      estimate
    };
  }, [shapes]);

  const currencyFormatter = useMemo(
    () =>
      new Intl.NumberFormat("en-GB", {
        style: "currency",
        currency: "GBP",
        maximumFractionDigits: 0
      }),
    []
  );

  function formatM2(value) {
    return `${(Number(value) || 0).toFixed(2)}m²`;
  }

  return (
    <div className="app-shell">
      <div className="page vinyl-estimator-page">
        <MainNavBar
          currentUser={currentUser}
          active="van-estimator"
          onLogout={onLogout}
          notifications={notifications}
        />

        <section className="panel vinyl-estimator-panel">
          <div className="vinyl-estimator-head">
            <div>
              <span className="eyebrow">Vehicle vinyl</span>
              <h2>Vinyl Estimator</h2>
              <p>
                Draw rectangles or point-click shapes on the van. Anything crossing a wrap-film line is counted as wrap film.
              </p>
            </div>
            <div className="vinyl-estimator-template">
              <span>{VAN_ESTIMATOR_TEMPLATE.name}</span>
              <strong>{VAN_ESTIMATOR_TEMPLATE.scaleReference}</strong>
            </div>
          </div>

          <div className="vinyl-estimator-grid">
            <div className="vinyl-canvas-card">
              <div className="vinyl-tool-row">
                <div className="vinyl-tool-segment" aria-label="Drawing mode">
                  <button
                    type="button"
                    className={drawMode === "rectangle" ? "active" : ""}
                    onClick={() => {
                      setDrawMode("rectangle");
                      setPolygonPoints([]);
                      setPolygonPreviewPoint(null);
                    }}
                  >
                    Rectangle
                  </button>
                  <button
                    type="button"
                    className={drawMode === "polygon" ? "active" : ""}
                    onClick={() => {
                      setDrawMode("polygon");
                      setDrawStart(null);
                      setDrawingRect(null);
                    }}
                  >
                    Point shape
                  </button>
                </div>
                {drawMode === "polygon" ? (
                  <div className="vinyl-point-actions">
                    <button className="ghost-button" type="button" disabled={polygonPoints.length < 3} onClick={finishPolygon}>
                      Finish shape
                    </button>
                    <button
                      className="ghost-button"
                      type="button"
                      disabled={!polygonPoints.length}
                      onClick={() => {
                        setPolygonPoints((current) => current.slice(0, -1));
                        setPolygonPreviewPoint(null);
                      }}
                    >
                      Undo point
                    </button>
                    <button
                      className="text-button danger"
                      type="button"
                      disabled={!polygonPoints.length}
                      onClick={() => {
                        setPolygonPoints([]);
                        setPolygonPreviewPoint(null);
                      }}
                    >
                      Cancel
                    </button>
                  </div>
                ) : null}
              </div>
              {svgError ? <div className="flash error">{svgError}</div> : null}
              <div className="vinyl-canvas">
                <div
                  ref={inlineSvgRef}
                  className="vinyl-template"
                  dangerouslySetInnerHTML={{ __html: svgMarkup }}
                />
                <svg
                  ref={overlaySvgRef}
                  className="vinyl-drawing-layer"
                  viewBox={`${VAN_ESTIMATOR_TEMPLATE.viewBox.x} ${VAN_ESTIMATOR_TEMPLATE.viewBox.y} ${VAN_ESTIMATOR_TEMPLATE.viewBox.width} ${VAN_ESTIMATOR_TEMPLATE.viewBox.height}`}
                  preserveAspectRatio="xMidYMid meet"
                  onPointerDown={startDrawing}
                  onPointerMove={updateDrawing}
                  onPointerUp={finishDrawing}
                  onPointerCancel={() => {
                    setDrawStart(null);
                    setDrawingRect(null);
                    setEditDrag(null);
                  }}
                >
                  {shapes.map((shape) => (
                    <g key={shape.id} className="vinyl-shape-group">
                      {shape.type === "polygon" ? (
                        <polygon
                          points={pointsToSvg(shape.points)}
                          className={`vinyl-shape ${shape.isWrapFilm ? "wrap" : "standard"}`}
                        />
                      ) : (
                        <rect
                          x={shape.x}
                          y={shape.y}
                          width={shape.width}
                          height={shape.height}
                          className={`vinyl-shape ${shape.isWrapFilm ? "wrap" : "standard"}`}
                        />
                      )}
                      <g
                        className="vinyl-shape-delete"
                        role="button"
                        tabIndex="0"
                        aria-label="Delete drawn area"
                        onPointerDown={(event) => {
                          event.stopPropagation();
                          event.preventDefault();
                        }}
                        onClick={(event) => deleteShape(event, shape.id)}
                      >
                        <circle cx={(shape.bounds || shape).x + (shape.bounds || shape).width - 16} cy={(shape.bounds || shape).y + 16} r="13" />
                        <line
                          x1={(shape.bounds || shape).x + (shape.bounds || shape).width - 21}
                          y1={(shape.bounds || shape).y + 11}
                          x2={(shape.bounds || shape).x + (shape.bounds || shape).width - 11}
                          y2={(shape.bounds || shape).y + 21}
                        />
                        <line
                          x1={(shape.bounds || shape).x + (shape.bounds || shape).width - 11}
                          y1={(shape.bounds || shape).y + 11}
                          x2={(shape.bounds || shape).x + (shape.bounds || shape).width - 21}
                          y2={(shape.bounds || shape).y + 21}
                        />
                      </g>
                      {(shape.type === "polygon" ? shape.points : getRectanglePoints(shape.bounds || shape)).map((point, pointIndex) => (
                        <circle
                          key={`${shape.id}-handle-${pointIndex}`}
                          cx={point.x}
                          cy={point.y}
                          r="8"
                          className="vinyl-corner-handle"
                          onPointerDown={(event) =>
                            startShapeCornerDrag(event, shape, shape.type === "polygon" ? pointIndex : point.corner)
                          }
                        />
                      ))}
                    </g>
                  ))}
                  {drawingRect ? (
                    <rect
                      x={drawingRect.x}
                      y={drawingRect.y}
                      width={drawingRect.width}
                      height={drawingRect.height}
                      className="vinyl-shape drawing"
                    />
                  ) : null}
                  {polygonPoints.length ? (
                    <g>
                      <polyline
                        points={pointsToSvg(polygonPreviewPoint ? [...polygonPoints, polygonPreviewPoint] : polygonPoints)}
                        className="vinyl-shape drawing vinyl-polygon-preview"
                      />
                      {polygonPoints.map((point, index) => (
                        <circle key={`polygon-point-${index}`} cx={point.x} cy={point.y} r="7" className="vinyl-point-handle" />
                      ))}
                    </g>
                  ) : null}
                </svg>
              </div>
              <div className="vinyl-canvas-actions">
                <span>
                  {drawMode === "polygon"
                    ? "Point shape: click each corner, then finish the shape."
                    : "Rectangle: drag across the panel you want to cover."}
                </span>
                <div>
                  <button
                    className="ghost-button"
                    type="button"
                    disabled={!shapes.length}
                    onClick={() => setShapes((current) => current.slice(0, -1))}
                  >
                    Undo last
                  </button>
                  <button
                    className="danger-button"
                    type="button"
                    disabled={!shapes.length}
                    onClick={() => setShapes([])}
                  >
                    Clear all
                  </button>
                </div>
              </div>
            </div>

            <aside className="vinyl-estimate-card">
              <div className="vinyl-total-hero">
                <span>Estimated supply & install</span>
                <strong>{currencyFormatter.format(totals.estimate)}</strong>
              </div>

              <div className="vinyl-stats-grid">
                <div>
                  <span>Standard vinyl</span>
                  <strong>{formatM2(totals.standardArea)}</strong>
                </div>
                <div className="wrap">
                  <span>Wrap film</span>
                  <strong>{formatM2(totals.wrapArea)}</strong>
                </div>
                <div>
                  <span>Flat area</span>
                  <strong>{formatM2(totals.flatArea)}</strong>
                </div>
                <div>
                  <span>Curved area</span>
                  <strong>{formatM2(totals.curvedArea)}</strong>
                </div>
                <div>
                  <span>Labour hours</span>
                  <strong>{totals.labourHours.toFixed(1)}</strong>
                </div>
                <div>
                  <span>Market anchor</span>
                  <strong>{currencyFormatter.format(totals.anchor)}</strong>
                </div>
              </div>

              <div className="vinyl-classification-note">
                <strong>Auto classified from drawn zones</strong>
                <span>Wrap-required zones use wrap film, curved surface rules, and wrap labour automatically.</span>
              </div>
              <div className="vinyl-breakdown">
                <div>
                  <span>Material</span>
                  <strong>{currencyFormatter.format(totals.materialSell)}</strong>
                </div>
                <div>
                  <span>Labour</span>
                  <strong>{currencyFormatter.format(totals.labourSell)}</strong>
                </div>
                <div>
                  <span>Total coverage</span>
                  <strong>{formatM2(totals.totalArea)}</strong>
                </div>
                <div>
                  <span>Vehicle coverage</span>
                  <strong>{Math.round(totals.coverage * 100)}%</strong>
                </div>
              </div>

              <div className="vinyl-shape-list">
                <h3>Drawn areas</h3>
                {shapes.length ? (
                  shapes.map((shape) => (
                    <div key={`shape-row-${shape.id}`} className="vinyl-shape-row">
                      <div>
                        <strong>
                          {VEHICLE_ZONE_MATERIALS[shape.zoneMetadata?.material_type]?.label ||
                            (shape.isWrapFilm ? "Wrap film" : "Standard vinyl")}
                        </strong>
                        <small>
                          {formatM2(shape.areaM2)} · {(shape.width * VAN_ESTIMATOR_TEMPLATE.scaleFactor).toFixed(0)}mm x{" "}
                          {(shape.height * VAN_ESTIMATOR_TEMPLATE.scaleFactor).toFixed(0)}mm
                        </small>
                      </div>
                      <button
                        className="text-button danger"
                        type="button"
                        onClick={() => setShapes((current) => current.filter((entry) => entry.id !== shape.id))}
                      >
                        Delete
                      </button>
                    </div>
                  ))
                ) : (
                  <p className="muted">No vinyl areas drawn yet.</p>
                )}
              </div>
            </aside>
          </div>
        </section>
      </div>
    </div>
  );
}

export default function App() {
  const pathname = typeof window !== "undefined" ? window.location.pathname.toLowerCase() : "/";
  const search = typeof window !== "undefined" ? window.location.search : "";
  const isClientRoute = pathname.startsWith("/client");
  const isClientBoardRoute = pathname.startsWith("/client/board");
  const isInstallerRoute = pathname.startsWith("/installer");
  const isAttendanceRoute = pathname.startsWith("/attendance");
  const isHolidaysRoute = pathname.startsWith("/holidays");
  const isMileageRoute = pathname.startsWith("/mileage");
  const isVanEstimatorRoute = pathname.startsWith("/van-estimator");
  const isNotificationsRoute = pathname.startsWith("/notifications");
  const isBoardRoute = pathname.startsWith("/board");
  const [board, setBoard] = useState(null);
  const [jobs, setJobs] = useState([]);
  const [holidays, setHolidays] = useState([]);
  const [holidayRequests, setHolidayRequests] = useState([]);
  const [approvedHolidayRequests, setApprovedHolidayRequests] = useState([]);
  const [holidayStaff, setHolidayStaff] = useState(HOLIDAY_STAFF);
  const [holidayAllowances, setHolidayAllowances] = useState([]);
  const [holidayEvents, setHolidayEvents] = useState([]);
  const [holidayAllowanceSavingKey, setHolidayAllowanceSavingKey] = useState("");
  const [holidayRequestOpen, setHolidayRequestOpen] = useState(false);
  const [holidayRequestForm, setHolidayRequestForm] = useState(EMPTY_HOLIDAY_REQUEST_FORM);
  const [holidayRequestSaving, setHolidayRequestSaving] = useState(false);
  const [holidayCancelOpen, setHolidayCancelOpen] = useState(false);
  const [holidayCancelForm, setHolidayCancelForm] = useState(EMPTY_HOLIDAY_CANCEL_FORM);
  const [holidayEventOpen, setHolidayEventOpen] = useState(false);
  const [holidayEventForm, setHolidayEventForm] = useState(EMPTY_HOLIDAY_EVENT_FORM);
  const [holidayEventSaving, setHolidayEventSaving] = useState(false);
  const [holidayYearStart, setHolidayYearStart] = useState(getCurrentHolidayYearStart());
  const [currentHolidayYearStart, setCurrentHolidayYearStart] = useState(getCurrentHolidayYearStart());
  const [attendanceData, setAttendanceData] = useState(null);
  const [attendanceMonthId, setAttendanceMonthId] = useState(toMonthIdFromIso(getLocalTodayIso()));
  const [attendanceLoading, setAttendanceLoading] = useState(false);
  const [attendanceSavingKey, setAttendanceSavingKey] = useState("");
  const [attendanceNoteSavingKey, setAttendanceNoteSavingKey] = useState("");
  const [form, setForm] = useState({ ...EMPTY_FORM, date: getLocalTodayIso() });
  const [editingId, setEditingId] = useState("");
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState(null);
  const [draggingJobId, setDraggingJobId] = useState("");
  const [duplicatingJobId, setDuplicatingJobId] = useState("");
  const [draggingHolidayId, setDraggingHolidayId] = useState("");
  const [dropDate, setDropDate] = useState("");
  const [activeHolidayDate, setActiveHolidayDate] = useState("");
  const [activeHolidayId, setActiveHolidayId] = useState("");
  const [jobModalDate, setJobModalDate] = useState("");
  const [activeClientJob, setActiveClientJob] = useState(null);
  const [clientCompletePrompt, setClientCompletePrompt] = useState(false);
  const [clientPhotoUploading, setClientPhotoUploading] = useState(false);
  const [clientExporting, setClientExporting] = useState(false);
  const [adminCompletePrompt, setAdminCompletePrompt] = useState(false);
  const [adminPhotoUploading, setAdminPhotoUploading] = useState(false);
  const [adminExporting, setAdminExporting] = useState(false);
  const [orderLookupOpen, setOrderLookupOpen] = useState(false);
  const [orderLookupQuery, setOrderLookupQuery] = useState("");
  const [orderLookupLoading, setOrderLookupLoading] = useState(false);
  const [orderLookupResults, setOrderLookupResults] = useState([]);
  const [orderLookupError, setOrderLookupError] = useState("");
  const [currentUser, setCurrentUser] = useState(null);
  const [authChecked, setAuthChecked] = useState(false);
  const [loginUsers, setLoginUsers] = useState([]);
  const [loginDisplayName, setLoginDisplayName] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginError, setLoginError] = useState("");
  const [permissionSavingKey, setPermissionSavingKey] = useState("");
  const [notifications, setNotifications] = useState([]);
  const [previousMonthDepth, setPreviousMonthDepth] = useState(0);
  const [futureMonthDepth, setFutureMonthDepth] = useState(0);
  const boardNotificationJobId = useMemo(() => {
    const params = new URLSearchParams(search);
    return params.get("job") || "";
  }, [search]);
  const attendanceNotificationDate = useMemo(() => {
    const params = new URLSearchParams(search);
    return params.get("date") || "";
  }, [search]);
  const dragPreviewRef = useRef(null);
  const transparentDragImageRef = useRef(null);
  const dragPositionRef = useRef({ x: 0, y: 0 });
  const clientPhotoInputRef = useRef(null);
  const adminPhotoInputRef = useRef(null);
  const openedNotificationJobIdRef = useRef("");
  const boardEditable = canEditBoard(currentUser);
  const installerEditable = canEditInstaller(currentUser);
  const attendanceEditable = canEditAttendance(currentUser);
  const hostShellMode = usesHostShell(currentUser);
  const isClientMode = currentUser ? !boardEditable : false;
  const showInstallerDirectory = Boolean(currentUser && canAccessInstaller(currentUser) && isInstallerRoute);
  const showAttendance = Boolean(currentUser && canAccessAttendance(currentUser) && isAttendanceRoute);
  const showHolidays = Boolean(currentUser && canAccessHolidays(currentUser) && isHolidaysRoute);
  const showMileage = Boolean(currentUser && canAccessMileage(currentUser) && isMileageRoute);
  const showVanEstimator = Boolean(currentUser && canAccessVanEstimator(currentUser) && isVanEstimatorRoute);
  const showNotifications = Boolean(currentUser && isNotificationsRoute);
  const showBoard = Boolean(
    currentUser &&
      canAccessBoard(currentUser) &&
      ((boardEditable && isBoardRoute) || (!boardEditable && isClientBoardRoute))
  );
  const showHostLanding = Boolean(currentUser && hostShellMode && !isInstallerRoute && !isBoardRoute && !isClientBoardRoute && !isAttendanceRoute && !isHolidaysRoute && !isMileageRoute && !isVanEstimatorRoute && !isNotificationsRoute);
  const showClientLanding = Boolean(currentUser && !hostShellMode && (canAccessBoard(currentUser) || canAccessAttendance(currentUser) || canAccessHolidays(currentUser) || canAccessMileage(currentUser) || canAccessVanEstimator(currentUser)) && !isClientBoardRoute && !isAttendanceRoute && !isHolidaysRoute && !isMileageRoute && !isVanEstimatorRoute && !isNotificationsRoute);
  const activeAdminJob = useMemo(() => {
    if (!editingId) return null;
    return jobs.find((job) => String(job.id || "") === String(editingId)) || null;
  }, [editingId, jobs]);

  const todayIso = board?.today || getLocalTodayIso();
  const rollingStartIso = useMemo(() => {
    const today = parseIsoDate(todayIso);
    return today ? toIsoDate(addDays(today, -7)) : "";
  }, [todayIso]);
  const rollingEndIso = useMemo(() => {
    const today = parseIsoDate(todayIso);
    return today ? toIsoDate(addDays(today, 21)) : "";
  }, [todayIso]);
  const boardRange = useMemo(() => {
    const today = parseIsoDate(todayIso);
    if (!today) {
      return { startIso: "", endIso: "" };
    }

    const currentMonthStart = getStartOfMonth(today);
    const start =
      previousMonthDepth > 0
        ? getStartOfMonth(addMonths(currentMonthStart, -previousMonthDepth))
        : parseIsoDate(rollingStartIso);
    const end =
      futureMonthDepth > 0
        ? getEndOfMonth(addMonths(currentMonthStart, futureMonthDepth))
        : parseIsoDate(rollingEndIso);

    return {
      startIso: start ? toIsoDate(start) : "",
      endIso: end ? toIsoDate(end) : ""
    };
  }, [futureMonthDepth, previousMonthDepth, rollingEndIso, rollingStartIso, todayIso]);

  function resetBoardWindow() {
    setPreviousMonthDepth(0);
    setFutureMonthDepth(0);
  }

  useEffect(() => {
    let active = true;

    async function loadAuth() {
      try {
        const meResponse = await fetch("/api/auth/me");
        const mePayload = meResponse.ok ? await meResponse.json() : null;
        let usersPayload = [];

        if (mePayload?.user) {
          const usersResponse = await fetch("/api/auth/users");
          usersPayload = usersResponse.ok ? await usersResponse.json() : [];
        }

        if (!active) return;

        setLoginUsers(Array.isArray(usersPayload) ? usersPayload : []);
        if (mePayload?.user) {
          setCurrentUser(mePayload.user);
        }
      } catch (error) {
        console.error(error);
      } finally {
        if (active) setAuthChecked(true);
      }
    }

    loadAuth();
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (!currentUser) {
      setNotifications([]);
      return undefined;
    }

    let active = true;

    async function loadNotifications() {
      try {
        const response = await fetch("/api/notifications");
        if (!response.ok) {
          throw new Error("Could not load notifications.");
        }
        const payload = await response.json();
        if (!active) return;
        setNotifications(Array.isArray(payload) ? payload : []);
      } catch (error) {
        console.error(error);
      }
    }

    loadNotifications();
    return () => {
      active = false;
    };
  }, [currentUser]);

  useEffect(() => {
    if (!currentUser || !showBoard) return undefined;
    let active = true;

    async function loadBoard() {
      try {
        setLoading(true);
        const [boardResponse, jobsResponse, holidaysResponse] = await Promise.all([
          fetch(buildBoardUrl(boardRange.startIso, boardRange.endIso)),
          fetch("/api/jobs"),
          fetch("/api/holidays")
        ]);
        if (!boardResponse.ok || !jobsResponse.ok || !holidaysResponse.ok) {
          throw new Error("Could not load the installation board.");
        }

        const [boardData, jobsData, holidaysData] = await Promise.all([
          boardResponse.json(),
          jobsResponse.json(),
          holidaysResponse.json()
        ]);
        if (!active) return;
        setBoard(boardData);
        setJobs(Array.isArray(jobsData) ? jobsData : []);
        setHolidays(Array.isArray(holidaysData) ? holidaysData : []);
      } catch (error) {
        console.error(error);
        if (active) setMessage(createMessage("Could not load the shared board.", "error"));
      } finally {
        if (active) setLoading(false);
      }
    }

    loadBoard();
    return () => {
      active = false;
    };
  }, [currentUser, showBoard, boardRange.endIso, boardRange.startIso]);

  useEffect(() => {
    if (!showBoard || !boardNotificationJobId || !Array.isArray(jobs) || !jobs.length) return;
    if (openedNotificationJobIdRef.current === String(boardNotificationJobId)) return;

    const matchedJob = jobs.find((job) => String(job.id || "") === String(boardNotificationJobId));
    if (!matchedJob) return;

    openedNotificationJobIdRef.current = String(boardNotificationJobId);
    if (isClientMode) {
      setActiveClientJob(matchedJob);
    } else {
      editJob(matchedJob);
    }

    const params = new URLSearchParams(window.location.search);
    params.delete("job");
    const nextSearch = params.toString();
    const nextUrl = `${window.location.pathname}${nextSearch ? `?${nextSearch}` : ""}${window.location.hash || ""}`;
    window.history.replaceState({}, "", nextUrl);
  }, [boardNotificationJobId, isClientMode, jobs, showBoard]);

  useEffect(() => {
    setClientCompletePrompt(false);
    setClientPhotoUploading(false);
    setClientExporting(false);
    if (clientPhotoInputRef.current) {
      clientPhotoInputRef.current.value = "";
    }
  }, [activeClientJob?.id]);

  useEffect(() => {
    setAdminCompletePrompt(false);
    setAdminPhotoUploading(false);
    setAdminExporting(false);
    if (adminPhotoInputRef.current) {
      adminPhotoInputRef.current.value = "";
    }
  }, [activeAdminJob?.id, jobModalDate]);

  useEffect(() => {
    if (!currentUser || !showHolidays) return undefined;
    const stream = new EventSource("/api/events");

    async function handleUpdate() {
      try {
        await refreshHolidayData();
      } catch (error) {
        console.error(error);
      }
    }

    stream.addEventListener("board-updated", handleUpdate);
    stream.onerror = () => {
      stream.close();
    };

    return () => {
      stream.removeEventListener("board-updated", handleUpdate);
      stream.close();
    };
  }, [currentUser, showHolidays]);

  useEffect(() => {
    if (!currentUser || !showHolidays) return undefined;
    let active = true;

    async function loadHolidayData() {
      try {
        setLoading(true);
        const response = await fetch(`/api/holiday-requests?yearStart=${encodeURIComponent(holidayYearStart)}`);
        if (!response.ok) {
          throw new Error("Could not load holiday calendar.");
        }

          const payload = await response.json();
          if (!active) return;
          setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
          setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
          setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
          setHolidayStaff(normalizeHolidayStaffEntries(payload.holidayStaff));
          setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
          setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
          setCurrentHolidayYearStart(Number(payload.currentHolidayYearStart || getCurrentHolidayYearStart()));
        } catch (error) {
        console.error(error);
        if (active) setMessage(createMessage("Could not load holiday calendar.", "error"));
      } finally {
        if (active) setLoading(false);
      }
    }

    loadHolidayData();
    return () => {
      active = false;
    };
  }, [currentUser, showHolidays, holidayYearStart]);

  useEffect(() => {
    if (!showAttendance) return;
    if (!attendanceNotificationDate) return;
    const focusMonthId = toMonthIdFromIso(attendanceNotificationDate);
    if (focusMonthId) {
      setAttendanceMonthId(focusMonthId);
    }
  }, [attendanceNotificationDate, showAttendance]);

  useEffect(() => {
    if (!currentUser || !showAttendance) return undefined;
    let active = true;

    async function loadAttendance() {
      try {
        setAttendanceLoading(true);
        const response = await fetch(`/api/attendance?month=${encodeURIComponent(attendanceMonthId)}`);
        if (!response.ok) {
          throw new Error("Could not load attendance.");
        }
        const payload = await response.json();
        if (!active) return;
        setAttendanceData(payload);
      } catch (error) {
        console.error(error);
        if (active) setMessage(createMessage(error.message || "Could not load attendance.", "error"));
      } finally {
        if (active) setAttendanceLoading(false);
      }
    }

    loadAttendance();
    return () => {
      active = false;
    };
  }, [attendanceMonthId, currentUser, showAttendance]);

  useEffect(() => {
    if (!currentUser) return;
    const nextHomePath = getHomePathForUser(currentUser);
    const nextBoardPath = getBoardPathForUser(currentUser);

    if (isHolidaysRoute && !canAccessHolidays(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if (isAttendanceRoute && !canAccessAttendance(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if (isMileageRoute && !canAccessMileage(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if (isVanEstimatorRoute && !canAccessVanEstimator(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if (isNotificationsRoute && !currentUser) {
      window.location.replace("/");
      return;
    }

    if (isInstallerRoute && !canAccessInstaller(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if ((isBoardRoute || isClientBoardRoute) && !canAccessBoard(currentUser)) {
      window.location.replace(nextHomePath);
      return;
    }

    if (hostShellMode && isClientRoute && !isClientBoardRoute) {
      window.location.replace(nextHomePath);
      return;
    }

    if (!hostShellMode && !isClientRoute && !isHolidaysRoute && !isAttendanceRoute && !isMileageRoute && !isVanEstimatorRoute && !isNotificationsRoute) {
      window.location.replace(nextHomePath);
      return;
    }

    if ((isBoardRoute || isClientBoardRoute) && nextBoardPath !== window.location.pathname) {
      window.location.replace(nextBoardPath);
    }
  }, [currentUser, isClientRoute, isClientBoardRoute, isInstallerRoute, isBoardRoute, isAttendanceRoute, isHolidaysRoute, isMileageRoute, isVanEstimatorRoute, isNotificationsRoute, hostShellMode]);

  useEffect(() => {
    if (!currentUser || !showBoard) return undefined;
    const stream = new EventSource("/api/events");

    async function handleUpdate() {
      try {
        const response = await fetch(buildBoardUrl(boardRange.startIso, boardRange.endIso));
        if (!response.ok) return;
        const nextBoard = await response.json();
        setBoard(nextBoard);
      } catch (error) {
        console.error(error);
      }
    }

    stream.addEventListener("board-updated", handleUpdate);
    stream.onerror = () => {
      stream.close();
      window.setTimeout(() => window.location.reload(), 3000);
    };

    return () => {
      stream.removeEventListener("board-updated", handleUpdate);
      stream.close();
    };
  }, [currentUser, showBoard, boardRange.endIso, boardRange.startIso]);

  useEffect(() => {
    if (!currentUser || !showAttendance) return undefined;
    const stream = new EventSource("/api/events");

    async function handleUpdate() {
      try {
        const response = await fetch(`/api/attendance?month=${encodeURIComponent(attendanceMonthId)}`);
        if (!response.ok) return;
        const payload = await response.json();
        setAttendanceData(payload);
      } catch (error) {
        console.error(error);
      }
    }

    stream.addEventListener("attendance-updated", handleUpdate);
    stream.onerror = () => {
      stream.close();
    };

    return () => {
      stream.removeEventListener("attendance-updated", handleUpdate);
      stream.close();
    };
  }, [attendanceMonthId, currentUser, showAttendance]);

  useEffect(() => {
    if (!message) return undefined;
    const timer = window.setTimeout(() => setMessage(null), 3000);
    return () => window.clearTimeout(timer);
  }, [message]);

  useEffect(() => {
    function handleWindowDragOver(event) {
      if (!dragPreviewRef.current) return;
      const deltaX = event.clientX - dragPositionRef.current.x;
      const deltaY = event.clientY - dragPositionRef.current.y;
      dragPositionRef.current = { x: event.clientX, y: event.clientY };

      const tilt = Math.max(-12, Math.min(12, deltaX * 0.6));
      const lift = Math.max(-8, Math.min(8, -deltaY * 0.25));
      dragPreviewRef.current.style.left = `${event.clientX + 18}px`;
      dragPreviewRef.current.style.top = `${event.clientY + 18}px`;
      dragPreviewRef.current.style.transform = `rotate(${tilt}deg) translateY(${lift}px)`;
    }

    function handleWindowDrop() {
      clearDragPreview();
    }

    window.addEventListener("dragover", handleWindowDragOver);
    window.addEventListener("drop", handleWindowDrop);

    return () => {
      window.removeEventListener("dragover", handleWindowDragOver);
      window.removeEventListener("drop", handleWindowDrop);
    };
  }, []);

  const upcomingJobs = useMemo(() => {
    const today = board?.today || getLocalTodayIso();
    return [...jobs]
      .filter((job) => job.date >= today)
      .sort((left, right) => {
        if (left.date !== right.date) return left.date.localeCompare(right.date);
        return left.customerName.localeCompare(right.customerName);
      })
      .slice(0, 8);
  }, [board?.today, jobs]);

  const jobsById = useMemo(() => {
    return new Map(jobs.map((job) => [job.id, job]));
  }, [jobs]);

  const holidaysById = useMemo(() => {
    return new Map(holidays.map((holiday) => [holiday.id, holiday]));
  }, [holidays]);

  const holidayYearLabel = useMemo(() => getHolidayYearLabel(holidayYearStart), [holidayYearStart]);
  const holidayRows = useMemo(
    () => buildHolidayYearRows(holidays, holidayYearStart, holidayEvents),
    [holidayEvents, holidayYearStart, holidays]
  );

  const backdropPointerStartedRef = useRef(false);

  function resetForm(nextDate = board?.today || getLocalTodayIso()) {
    setEditingId("");
    setForm({ ...EMPTY_FORM, date: nextDate });
    setJobModalDate("");
    setOrderLookupOpen(false);
    setOrderLookupQuery("");
    setOrderLookupResults([]);
    setOrderLookupError("");
  }

  function preserveScrollPosition() {
    const currentScrollY = window.scrollY;
    window.requestAnimationFrame(() => {
      window.scrollTo({ top: currentScrollY, behavior: "auto" });
    });
  }

  function openJobModal(nextDate, nextForm, nextEditingId = "") {
    setEditingId(nextEditingId);
    setForm(nextForm);
    setJobModalDate(nextDate);
    preserveScrollPosition();
  }

  function handleBackdropPointerDown(event) {
    backdropPointerStartedRef.current = event.target === event.currentTarget;
  }

  function handleBackdropClick(event, onClose) {
    const shouldClose = backdropPointerStartedRef.current && event.target === event.currentTarget;
    backdropPointerStartedRef.current = false;
    if (shouldClose) {
      onClose();
    }
  }

  async function handleLogin(event) {
    event.preventDefault();
    try {
      setLoginLoading(true);
      setLoginError("");
      const response = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          displayName: loginDisplayName,
          password: loginPassword
        })
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not sign in.");
      }

      setCurrentUser(payload.user);
      setLoginPassword("");
      window.location.replace(getHomePathForUser(payload.user));
    } catch (error) {
      console.error(error);
      setLoginError(error.message || "Could not sign in.");
    } finally {
      setLoginLoading(false);
    }
  }

  async function handleLogout() {
    try {
      await fetch("/api/auth/logout", { method: "POST" });
    } catch (error) {
      console.error(error);
    } finally {
      setCurrentUser(null);
      setBoard(null);
      setJobs([]);
      setHolidays([]);
      setHolidayRequests([]);
      setApprovedHolidayRequests([]);
      setLoginPassword("");
      setLoginError("");
      window.location.replace(isClientRoute ? "/client" : "/");
    }
  }

  async function handlePermissionChange(userId, appKey, value) {
    const targetUser = loginUsers.find((entry) => entry.id === userId);
    if (!targetUser || !currentUser?.canManagePermissions) return;

    const nextPermissions = {
      board: getPermissionForApp(targetUser, "board"),
      installer: getPermissionForApp(targetUser, "installer"),
      holidays: getPermissionForApp(targetUser, "holidays"),
      attendance: getPermissionForApp(targetUser, "attendance"),
      mileage: getPermissionForApp(targetUser, "mileage"),
      vanEstimator: getPermissionForApp(targetUser, "vanEstimator"),
      [appKey]: value
    };

    setPermissionSavingKey(`${userId}:${appKey}`);

    try {
      const response = await fetch(`/api/auth/users/${encodeURIComponent(userId)}/permissions`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(nextPermissions)
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not update permissions.");
      }

      setLoginUsers((current) =>
        current.map((entry) => (entry.id === userId ? { ...entry, ...payload.user } : entry))
      );
      setMessage(createMessage(`Updated ${targetUser.displayName}'s permissions.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not update permissions.", "error"));
    } finally {
      setPermissionSavingKey("");
    }
  }

  async function handleUpdateAttendanceProfile(userId, attendanceProfile) {
    const targetUser = loginUsers.find((entry) => entry.id === userId);
    if (!targetUser || !currentUser?.canManagePermissions) return;

    setPermissionSavingKey(`${userId}:attendance-profile`);

    try {
      const response = await fetch(`/api/auth/users/${encodeURIComponent(userId)}/attendance-profile`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(attendanceProfile)
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not update attendance settings.");
      }

      setLoginUsers((current) =>
        current.map((entry) => (entry.id === userId ? { ...entry, ...payload.user } : entry))
      );
      if (currentUser?.id === userId) {
        setCurrentUser((existing) => ({ ...existing, ...payload.user }));
      }
      if (showAttendance) {
        const attendanceResponse = await fetch(`/api/attendance?month=${encodeURIComponent(attendanceMonthId)}`);
        if (attendanceResponse.ok) {
          const attendancePayload = await attendanceResponse.json();
          setAttendanceData(attendancePayload);
        }
      }
      setMessage(createMessage(`Updated ${targetUser.displayName}'s attendance settings.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not update attendance settings.", "error"));
    } finally {
      setPermissionSavingKey("");
    }
  }

  async function handleCreateUser({ displayName, role, password }) {
    setPermissionSavingKey("create-user");
    try {
      const response = await fetch("/api/auth/users", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ displayName, role, password })
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not create user.");
      }

      setLoginUsers((current) => [...current, payload.user].sort((left, right) => left.displayName.localeCompare(right.displayName)));
      setMessage(createMessage(`Added ${payload.user.displayName}.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not create user.", "error"));
      throw error;
    } finally {
      setPermissionSavingKey("");
    }
  }

  async function handleResetUserPassword(userId, password) {
    setPermissionSavingKey(`${userId}:password`);
    try {
      const response = await fetch(`/api/auth/users/${encodeURIComponent(userId)}/password`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ password })
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not update password.");
      }

      setLoginUsers((current) =>
        current.map((entry) => (entry.id === userId ? { ...entry, ...payload.user } : entry))
      );
      setMessage(createMessage(`Updated ${payload.user.displayName}'s password.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not update password.", "error"));
      throw error;
    } finally {
      setPermissionSavingKey("");
    }
  }

  async function handleDeleteUser(user) {
    if (!window.confirm(`Delete ${user.displayName}?`)) return;
    setPermissionSavingKey(`${user.id}:delete`);
    try {
      const response = await fetch(`/api/auth/users/${encodeURIComponent(user.id)}`, {
        method: "DELETE"
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not delete user.");
      }

      setLoginUsers((current) => current.filter((entry) => entry.id !== user.id));
      setMessage(createMessage(`Deleted ${user.displayName}.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not delete user.", "error"));
    } finally {
      setPermissionSavingKey("");
    }
  }

  async function markNotificationRead(notificationId) {
    setNotifications((current) =>
      current.map((entry) =>
        entry.id === notificationId
          ? {
              ...entry,
              read: true
            }
          : entry
      )
    );

    try {
      const response = await fetch(`/api/notifications/${encodeURIComponent(notificationId)}/read`, {
        method: "PATCH"
      });
      if (!response.ok) {
        throw new Error("Could not update notification.");
      }
      const payload = await response.json();
      setNotifications(Array.isArray(payload) ? payload : []);
    } catch (error) {
      await refreshNotifications();
      throw error;
    }
  }

  async function markAllNotificationsRead() {
    try {
      setNotifications((current) => current.map((entry) => ({ ...entry, read: true })));
      const response = await fetch("/api/notifications/read-all", { method: "POST" });
      if (!response.ok) {
        throw new Error("Could not update notifications.");
      }
      const payload = await response.json();
      setNotifications(Array.isArray(payload) ? payload : []);
    } catch (error) {
      await refreshNotifications();
      console.error(error);
      setMessage(createMessage(error.message || "Could not update notifications.", "error"));
    }
  }

  async function openNotification(notification) {
    try {
      if (!notification?.read) {
        await markNotificationRead(notification.id);
      }
    } catch (error) {
      console.error(error);
    }

    if (notification?.link) {
      window.location.assign(notification.link);
    }
  }

  function editJob(job) {
    openJobModal(job.date || "Unscheduled", {
      id: job.id,
      date: job.date,
      orderReference: job.orderReference || "",
      customerName: job.customerName || "",
      description: job.description || "",
      contact: job.contact || "",
      number: job.number || "",
      address: job.address || "",
      installers: Array.isArray(job.installers)
        ? job.installers
        : typeof job.installers === "string" && job.installers.trim()
          ? job.installers.split(/[,/]+/).map((item) => item.trim()).filter(Boolean)
          : [],
      customInstaller: job.customInstaller || "",
      jobType: job.jobType || "Install",
      customJobType: job.customJobType || "",
      isPlaceholder: Boolean(job.isPlaceholder),
      notes: job.notes || ""
    }, job.id);
  }

  async function refreshData() {
    const [jobsResponse, holidaysResponse] = await Promise.all([fetch("/api/jobs"), fetch("/api/holidays")]);
    if (!jobsResponse.ok || !holidaysResponse.ok) throw new Error("Could not refresh data.");
    const [nextJobs, nextHolidays] = await Promise.all([jobsResponse.json(), holidaysResponse.json()]);
    setJobs(Array.isArray(nextJobs) ? nextJobs : []);
    setHolidays(Array.isArray(nextHolidays) ? nextHolidays : []);
  }

  async function refreshHolidayData() {
    const response = await fetch(`/api/holiday-requests?yearStart=${encodeURIComponent(holidayYearStart)}`);
    if (!response.ok) throw new Error("Could not refresh holiday calendar.");
      const payload = await response.json();
      setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
      setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
      setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
      setHolidayStaff(normalizeHolidayStaffEntries(payload.holidayStaff));
      setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
      setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
      setHolidayYearStart(Number(payload.holidayYearStart || holidayYearStart));
      setCurrentHolidayYearStart(Number(payload.currentHolidayYearStart || getCurrentHolidayYearStart()));
    }

  async function refreshNotifications() {
    const response = await fetch("/api/notifications");
    if (!response.ok) throw new Error("Could not refresh notifications.");
    const payload = await response.json();
    setNotifications(Array.isArray(payload) ? payload : []);
  }

  async function saveAttendanceEntry({ person, date, clockIn, clockOut, adminNote = "" }) {
    setAttendanceSavingKey(`${person}:${date}`);
    try {
      const response = await fetch("/api/attendance/entries", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ person, date, clockIn, clockOut, adminNote })
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not save attendance.");
      }
      setAttendanceData(payload);
      setMessage(createMessage(`Updated attendance for ${person} on ${formatJobDate(date)}.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not save attendance.", "error"));
    } finally {
      setAttendanceSavingKey("");
    }
  }

  async function submitAttendanceExplanation({ date, note }) {
    setAttendanceNoteSavingKey(date);
    try {
      const response = await fetch("/api/attendance/explanations", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ date, note })
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not send attendance note.");
      }
      setAttendanceData(payload);
      await refreshNotifications();
      setMessage(createMessage("Attendance note sent.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not send attendance note.", "error"));
    } finally {
      setAttendanceNoteSavingKey("");
    }
  }

  async function searchCoreBridgeOrders(searchTerm = orderLookupQuery) {
    if (isClientMode) return;

    try {
      setOrderLookupLoading(true);
      setOrderLookupError("");
      const query = String(searchTerm || "").trim();
      const params = new URLSearchParams();
      if (query) params.set("q", query);
      const url = params.toString() ? `/api/corebridge/orders?${params.toString()}` : "/api/corebridge/orders";
      const response = await fetch(url);
      const raw = await response.text();
      let payload = {};

      try {
        payload = raw ? JSON.parse(raw) : {};
      } catch (error) {
        throw new Error("CoreBridge returned an invalid response.");
      }

      if (!response.ok) {
        throw new Error(payload.detail || payload.error || "Could not load CoreBridge orders.");
      }

      setOrderLookupResults(Array.isArray(payload.orders) ? payload.orders : []);
    } catch (error) {
      console.error(error);
      setOrderLookupResults([]);
      setOrderLookupError(error.message || "Could not load CoreBridge orders.");
    } finally {
      setOrderLookupLoading(false);
    }
  }

  async function submitHolidayRequest() {
    const person = canEditHolidays(currentUser)
      ? holidayRequestForm.person
      : getHolidayStaffPersonForUser(currentUser);

    if (!person || !holidayRequestForm.startDate || !holidayRequestForm.endDate) {
      setMessage(createMessage("Choose a person plus start and end dates.", "error"));
      return;
    }

    setHolidayRequestSaving(true);

    try {
      const response = await fetch("/api/holiday-requests", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          person,
          holidayYearStart,
          startDate: holidayRequestForm.startDate,
          endDate: holidayRequestForm.endDate,
          duration: holidayRequestForm.duration,
          notes: holidayRequestForm.notes
        })
      });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "Could not send holiday request.");
        setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
        setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
        setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
        setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
        setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
        setHolidayRequestForm(EMPTY_HOLIDAY_REQUEST_FORM);
        setHolidayRequestOpen(false);
        await refreshNotifications();
        setMessage(createMessage("Holiday request sent.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not send holiday request.", "error"));
    } finally {
      setHolidayRequestSaving(false);
    }
  }

  async function reviewHolidayRequest(requestId, status) {
    try {
      const response = await fetch(`/api/holiday-requests/${requestId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ status })
      });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "Could not update holiday request.");
        setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
        setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
        setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
        setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
        setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
        await refreshNotifications();
        setMessage(createMessage(`Holiday request ${status}.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not update holiday request.", "error"));
    }
  }

  async function saveHolidayAllowance(person, updates) {
    if (!canEditHolidays(currentUser)) return;
    const existing = holidayAllowances.find((entry) => entry.person === person);
    const savingField = Object.keys(updates)[0] || "";
    setHolidayAllowanceSavingKey(`${person}:${savingField}`);

    try {
      const response = await fetch("/api/holiday-allowances", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...existing,
          person,
          yearStart: holidayYearStart,
          ...updates
        })
      });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "Could not save holiday allowance.");
        setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
        setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
        setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
        setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
        setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
        setMessage(createMessage(`Updated ${person}'s holiday allowance.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not save holiday allowance.", "error"));
    } finally {
      setHolidayAllowanceSavingKey("");
    }
  }

  function changeHolidayAllowanceDraft(person, updates) {
    setHolidayAllowances((current) =>
      current.map((entry) =>
        entry.person === person
          ? {
              ...entry,
              ...updates
            }
          : entry
      )
    );
  }

  async function openOrderLookup() {
    setOrderLookupOpen(true);
    setOrderLookupError("");
  }

  async function applyCoreBridgeOrder(order) {
    let resolvedOrder = order;

    try {
      if (order?.id) {
        setOrderLookupLoading(true);
        setOrderLookupError("");
        const url = `/api/corebridge/orders/${encodeURIComponent(order.id)}`;
        const response = await fetch(url);
        const raw = await response.text();
        const payload = raw ? JSON.parse(raw) : {};
        if (!response.ok) {
          throw new Error(payload.detail || payload.error || "Could not load CoreBridge order detail.");
        }
        resolvedOrder = payload;
      }
    } catch (error) {
      console.error(error);
      setOrderLookupError(error.message || "Could not load CoreBridge order detail.");
      return;
    } finally {
      setOrderLookupLoading(false);
    }

    setForm((current) => ({
      ...current,
      orderReference: resolvedOrder.orderReference ?? "",
      customerName: resolvedOrder.customerName ?? "",
      description: resolvedOrder.description ?? "",
      contact: resolvedOrder.contact ?? "",
      number: resolvedOrder.number ?? "",
      address: resolvedOrder.address ?? "",
      notes: resolvedOrder.notes ?? ""
    }));
    setOrderLookupOpen(false);
    setMessage(createMessage("Order details copied into the job form.", "success"));
  }

  function applyBoardPayloadToState(payload, fallbackJobId = "") {
    if (payload?.board) setBoard(payload.board);
    if (Array.isArray(payload?.jobs)) setJobs(payload.jobs);
    if (Array.isArray(payload?.holidays)) setHolidays(payload.holidays);

    if (payload?.job) {
      setActiveClientJob(payload.job);
      return;
    }

    if (fallbackJobId && Array.isArray(payload?.jobs)) {
      const matched = payload.jobs.find((job) => String(job.id || "") === String(fallbackJobId));
      if (matched) {
        setActiveClientJob(matched);
      }
    }
  }

  async function cancelHolidayRequest() {
    const targetHolidayRequest = cancellableHolidayRequests.find(
      (request) => String(request.id || "") === String(holidayCancelForm.requestId || "")
    );
    if (!targetHolidayRequest) {
      setMessage(createMessage("Choose an approved holiday request to cancel.", "error"));
      return;
    }

    setHolidayRequestSaving(true);
    try {
      const response = await fetch("/api/holiday-requests", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          person: targetHolidayRequest.person,
          holidayYearStart,
          startDate: targetHolidayRequest.startDate,
          endDate: targetHolidayRequest.endDate,
          duration: targetHolidayRequest.duration,
          notes: holidayCancelForm.notes,
          action: "cancel",
          targetRequestId: targetHolidayRequest.id
        })
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not send cancellation request.");
      setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
      setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
      setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
      setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
      setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
      setHolidayCancelForm(EMPTY_HOLIDAY_CANCEL_FORM);
      setHolidayCancelOpen(false);
      await refreshNotifications();
      setMessage(createMessage("Holiday cancellation request sent.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not send cancellation request.", "error"));
    } finally {
      setHolidayRequestSaving(false);
    }
  }

  async function completeJob(jobId) {
    const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/complete`, {
      method: "POST"
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Could not mark the job as complete.");
    }
    applyBoardPayloadToState(payload, jobId);
    return payload;
  }

  async function snaggingJob(jobId) {
    const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/snagging`, {
      method: "POST"
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Could not mark the job as snagging.");
    }
    applyBoardPayloadToState(payload, jobId);
    return payload;
  }

  async function clearSnaggingJob(jobId) {
    const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/unsnagging`, {
      method: "POST"
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Could not remove the snagging tag.");
    }
    applyBoardPayloadToState(payload, jobId);
    return payload;
  }

  async function undoCompleteJob(jobId) {
    const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/uncomplete`, {
      method: "POST"
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Could not undo the completed status.");
    }
    applyBoardPayloadToState(payload, jobId);
    return payload;
  }

  async function uploadJobPhotos(jobId, files) {
    for (const file of files) {
      const prepared = await compressPhotoForUpload(file);
      const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/photos`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(prepared)
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || `Could not upload ${file.name}.`);
      }
      applyBoardPayloadToState(payload, jobId);
    }
  }

  async function markClientJobComplete(job, uploadFiles = []) {
    if (!job?.id) return;
    const files = Array.from(uploadFiles || []);
    if (files.length) {
      setClientPhotoUploading(true);
    }

    try {
      if (!job.isCompleted) {
        await completeJob(job.id);
      }
      if (files.length) {
        await uploadJobPhotos(job.id, files);
      }
      setClientCompletePrompt(false);
      setMessage(
        createMessage(
          files.length ? "Job marked complete and photos uploaded." : "Job marked complete.",
          "success"
        )
      );
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not complete the job.", "error"));
    } finally {
      setClientPhotoUploading(false);
      if (clientPhotoInputRef.current) {
        clientPhotoInputRef.current.value = "";
      }
    }
  }

  async function markAdminJobComplete(job, uploadFiles = []) {
    if (!job?.id) return;
    const files = Array.from(uploadFiles || []);
    if (files.length) {
      setAdminPhotoUploading(true);
    }

    try {
      if (!job.isCompleted) {
        await completeJob(job.id);
      }
      if (files.length) {
        await uploadJobPhotos(job.id, files);
      }
      setAdminCompletePrompt(false);
      setMessage(
        createMessage(
          files.length ? "Job marked complete and photos uploaded." : "Job marked complete.",
          "success"
        )
      );
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not complete the job.", "error"));
    } finally {
      setAdminPhotoUploading(false);
      if (adminPhotoInputRef.current) {
        adminPhotoInputRef.current.value = "";
      }
    }
  }

  async function markAdminJobSnagging(job) {
    if (!job?.id) return;
    try {
      await snaggingJob(job.id);
      setAdminCompletePrompt(false);
      setMessage(createMessage("Job marked as snagging.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not mark the job as snagging.", "error"));
    }
  }

  async function removeAdminJobSnagging(job) {
    if (!job?.id) return;
    try {
      await clearSnaggingJob(job.id);
      setAdminCompletePrompt(false);
      setMessage(createMessage("Snagging removed.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not remove snagging.", "error"));
    }
  }

  async function undoClientJobComplete(job) {
    if (!job?.id) return;
    try {
      await undoCompleteJob(job.id);
      setClientCompletePrompt(false);
      setMessage(createMessage("Job marked as not complete.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not undo completion.", "error"));
    }
  }

  async function undoAdminJobComplete(job) {
    if (!job?.id) return;
    try {
      await undoCompleteJob(job.id);
      setAdminCompletePrompt(false);
      setMessage(createMessage("Job marked as not complete.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not undo completion.", "error"));
    }
  }

  async function exportJob(job, setExportingState) {
    if (!job?.id) return;
    setExportingState(true);
    try {
      const printWindow = window.open("", "_blank");
      if (!printWindow) {
        throw new Error("Please allow pop-ups so the PDF export can open.");
      }

      const uploadedBy = [...new Set((job.photos || []).map((photo) => String(photo.uploadedByName || "").trim()).filter(Boolean))].join(", ") || "-";
      const installers = getInstallerDisplayList(job).join(", ") || "-";
      const firstPagePhotos = (job.photos || []).slice(0, 2);
      const remainingPhotoPages = [];
      for (let index = 2; index < (job.photos || []).length; index += 6) {
        remainingPhotoPages.push((job.photos || []).slice(index, index + 6));
      }

      const renderPhotoTile = (photo, index, extraClass = "") => `
        <figure class="photo-tile ${extraClass}">
          <div class="photo-frame">
            <img src="${escapeHtml(photo.url || buildJobPhotoUrl(job.id, photo.id))}" alt="Job photo ${index + 1}" />
          </div>
        </figure>
      `;

      const html = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>${escapeHtml(job.orderReference || job.customerName || "Job Export")}</title>
    <style>
      @font-face {
        font-family: 'Faricy';
        src: url('${window.location.origin}/fonts/FARICYNEW-MEDIUM.TTF') format('truetype');
        font-weight: 500;
      }
      @font-face {
        font-family: 'Faricy';
        src: url('${window.location.origin}/fonts/FARICYNEW-BOLD.TTF') format('truetype');
        font-weight: 700;
      }
      :root {
        color-scheme: light;
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        font-family: 'Faricy', Arial, sans-serif;
        color: #1f2937;
        background: white;
      }
      .page {
        width: 210mm;
        min-height: 297mm;
        padding: 14mm;
        page-break-after: always;
      }
      .page:last-child { page-break-after: auto; }
      .header {
        display: flex;
        justify-content: space-between;
        align-items: start;
        gap: 16mm;
        margin-bottom: 8mm;
      }
      .header-copy h1 {
        margin: 0 0 2mm;
        font-size: 22px;
        line-height: 1.05;
      }
      .header-copy p {
        margin: 0;
        font-size: 12px;
      }
      .brand {
        width: 52mm;
        flex: 0 0 auto;
      }
      .brand img {
        display: block;
        width: 100%;
        height: auto;
      }
      .summary-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 4mm 8mm;
        margin-bottom: 8mm;
      }
      .summary-item {
        border: 1px solid #e2e8f0;
        border-radius: 4mm;
        padding: 3.2mm 3.6mm;
        min-height: 16mm;
      }
      .summary-item.wide {
        grid-column: 1 / -1;
      }
      .summary-item strong {
        display: block;
        margin-bottom: 1.2mm;
        color: #475569;
        font-size: 9px;
        letter-spacing: 0.08em;
        text-transform: uppercase;
      }
      .summary-item span {
        font-size: 12px;
        line-height: 1.35;
      }
      .photo-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 6mm;
      }
      .photo-grid.first-page {
        margin-top: 4mm;
      }
      .photo-grid.extra-page {
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 5mm;
      }
      .photo-tile {
        margin: 0;
      }
      .photo-tile.first-page-photo .photo-frame {
        aspect-ratio: 1 / 1;
      }
      .photo-frame {
        border: 1px solid #dbe2ea;
        border-radius: 4mm;
        overflow: hidden;
        aspect-ratio: 1 / 1;
        background: #f8fafc;
        display: flex;
        align-items: center;
        justify-content: center;
      }
      .photo-frame img {
        width: 100%;
        height: 100%;
        object-fit: cover;
        display: block;
      }
      @page {
        size: A4;
        margin: 0;
      }
      @media print {
        body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      }
    </style>
  </head>
  <body>
    <section class="page">
      <div class="header">
        <div class="header-copy">
          <h1>${escapeHtml(job.customerName || "Job Export")}</h1>
          <p>${escapeHtml(job.description || "")}</p>
        </div>
        <div class="brand">
          <img src="${window.location.origin}/branding/signs-express-logo.svg" alt="Signs Express logo" />
        </div>
      </div>
      <div class="summary-grid">
        <div class="summary-item"><strong>Order Ref</strong><span>${escapeHtml(job.orderReference || "-")}</span></div>
        <div class="summary-item"><strong>Completion Date</strong><span>${escapeHtml(formatJobDate(job.date) || "-")}</span></div>
        <div class="summary-item"><strong>Job Type</strong><span>${escapeHtml(getJobTypeLabel(job))}</span></div>
        <div class="summary-item"><strong>Installers</strong><span>${escapeHtml(installers)}</span></div>
        <div class="summary-item"><strong>Contact</strong><span>${escapeHtml(job.contact || "-")}</span></div>
        <div class="summary-item"><strong>Number</strong><span>${escapeHtml(job.number || "-")}</span></div>
        <div class="summary-item wide"><strong>Address</strong><span>${escapeHtml(job.address || "-")}</span></div>
        <div class="summary-item wide"><strong>Photos Uploaded By</strong><span>${escapeHtml(uploadedBy)}</span></div>
      </div>
      ${firstPagePhotos.length ? `<div class="photo-grid first-page">${firstPagePhotos.map((photo, index) => renderPhotoTile(photo, index, "first-page-photo")).join("")}</div>` : ""}
    </section>
    ${remainingPhotoPages.map((pagePhotos) => `
      <section class="page">
        <div class="photo-grid extra-page">${pagePhotos.map((photo, index) => renderPhotoTile(photo, index)).join("")}</div>
      </section>
    `).join("")}
    <script>
      const images = Array.from(document.images);
      Promise.all(images.map((image) => image.complete ? Promise.resolve() : new Promise((resolve) => {
        image.addEventListener('load', resolve, { once: true });
        image.addEventListener('error', resolve, { once: true });
      }))).then(() => {
        setTimeout(() => {
          window.focus();
          window.print();
          setTimeout(() => {
            window.close();
          }, 300);
        }, 250);
      });
    </script>
  </body>
</html>`;

      printWindow.document.open();
      printWindow.document.write(html);
      printWindow.document.close();
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not export the job.", "error"));
    } finally {
      setExportingState(false);
    }
  }

  async function deleteJobPhoto(job, photoId) {
    if (!job?.id || !photoId) return;
    try {
      const response = await fetch(
        `/api/jobs/${encodeURIComponent(job.id)}/photos/${encodeURIComponent(photoId)}`,
        { method: "DELETE" }
      );
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "Could not delete the photo.");
      }
      applyBoardPayloadToState(payload, job.id);
      setMessage(createMessage("Photo removed.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not delete the photo.", "error"));
    }
  }

  async function handleSubmit(event) {
    event.preventDefault();
    if (isClientMode) return;

    if (!form.customerName.trim()) {
      setMessage(createMessage("Enter the customer name.", "error"));
      return;
    }

    setSaving(true);

    try {
      const existingJob = editingId ? jobsById.get(editingId) : null;
      const response = await fetch("/api/jobs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...existingJob,
          id: editingId || form.id || undefined,
          date: form.date,
          orderReference: form.orderReference.trim(),
          customerName: form.customerName.trim(),
          description: form.description.trim(),
          contact: form.contact.trim(),
          number: form.number.trim(),
          address: form.address.trim(),
          installers: form.installers,
          customInstaller: form.customInstaller.trim(),
          jobType: form.jobType,
          customJobType: form.customJobType.trim(),
          isPlaceholder: Boolean(form.isPlaceholder),
          notes: form.notes.trim()
        })
      });

      if (!response.ok) throw new Error("Could not save the job.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      setMessage(createMessage(editingId ? "Job updated." : "Job added.", "success"));
      resetForm(form.date);
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not save the job.", "error"));
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete(jobId) {
    if (isClientMode) return;
    try {
      const response = await fetch(`/api/jobs/${jobId}`, { method: "DELETE" });
      if (!response.ok) throw new Error("Could not delete the job.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      if (editingId === jobId) resetForm();
      setMessage(createMessage("Job deleted.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not delete the job.", "error"));
    }
  }

  async function moveJobToDate(jobId, nextDate) {
    if (isClientMode) return;
    const job = jobsById.get(jobId);
    const normalizedNextDate = nextDate === UNSCHEDULED_DROP_ZONE ? "" : nextDate;
    if (!job || normalizedNextDate === undefined || normalizedNextDate === null || job.date === normalizedNextDate) {
      setDropDate("");
      setDraggingJobId("");
      return;
    }

    try {
      const response = await fetch("/api/jobs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...job,
          date: normalizedNextDate
        })
      });

      if (!response.ok) throw new Error("Could not move the job.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      if (editingId === jobId) {
        setForm((current) => ({ ...current, date: normalizedNextDate }));
      }
      setMessage(
        createMessage(
          normalizedNextDate ? `Job moved to ${normalizedNextDate}.` : "Job moved to unscheduled.",
          "success"
        )
      );
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not move the job.", "error"));
    } finally {
      setDropDate("");
      setDraggingJobId("");
    }
  }

  async function duplicateJobToDate(jobId, nextDate) {
    if (isClientMode) return;
    const job = jobsById.get(jobId);
    const normalizedNextDate = nextDate === UNSCHEDULED_DROP_ZONE ? "" : nextDate;
    if (!job || normalizedNextDate === undefined || normalizedNextDate === null) {
      setDuplicatingJobId("");
      setDropDate("");
      return;
    }

    try {
      const response = await fetch("/api/jobs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...job,
          id: undefined,
          date: normalizedNextDate,
          isCompleted: false,
          completedAt: "",
          completedByUserId: "",
          completedByName: "",
          photos: []
        })
      });

      if (!response.ok) throw new Error("Could not duplicate the job.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      setMessage(
        createMessage(
          normalizedNextDate ? `Job copied to ${normalizedNextDate}.` : "Job copied to unscheduled.",
          "success"
        )
      );
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not duplicate the job.", "error"));
    } finally {
      setDuplicatingJobId("");
      setDropDate("");
      clearDragPreview();
    }
  }

  function getJobTypeMeta(jobType) {
    return JOB_TYPES.find((option) => option.value === jobType) || JOB_TYPES[JOB_TYPES.length - 1];
  }

  function getJobTypeLabel(job) {
    return job.jobType === "Other" ? job.customJobType || "Other" : job.jobType;
  }

  function toggleInstaller(value) {
    setForm((current) => ({
      ...current,
      installers: current.installers.includes(value)
        ? current.installers.filter((item) => item !== value)
        : [...current.installers, value]
    }));
  }

  function getInstallerMeta(value) {
    return INSTALLER_OPTIONS.find((option) => option.value === value) || { value, colorClass: "installer-custom" };
  }

  function getInstallerDisplayList(item) {
    const source = Array.isArray(item.installers)
      ? item.installers
      : typeof item.installers === "string" && item.installers.trim()
        ? item.installers.split(/[,/]+/).map((entry) => entry.trim()).filter(Boolean)
        : [];

    const visible = source.filter((entry) => entry !== "Custom");
    if (source.includes("Custom") && item.customInstaller) {
      visible.push(item.customInstaller);
    }
    return visible;
  }

  function getHolidayPersonColor(person) {
    return HOLIDAY_PERSON_COLORS[person] || "holiday-person-black";
  }

  function buildDragPreview(element) {
    const preview = element.cloneNode(true);
    preview.classList.add("drag-preview");
    preview.style.width = `${element.offsetWidth}px`;
    preview.style.position = "fixed";
    preview.style.top = "0";
    preview.style.left = "0";
    document.body.appendChild(preview);
    return preview;
  }

  function getTransparentDragImage() {
    if (!transparentDragImageRef.current) {
      const image = new Image();
      image.src =
        "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==";
      transparentDragImageRef.current = image;
    }

    return transparentDragImageRef.current;
  }

  function clearDragPreview() {
    if (dragPreviewRef.current) {
      dragPreviewRef.current.remove();
      dragPreviewRef.current = null;
    }
  }

  async function saveHoliday(date, person, duration = "Full Day", id) {
    if (isClientMode) return;
    try {
      const response = await fetch("/api/holidays", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          id,
          date,
          person,
          duration
        })
      });
      if (!response.ok) throw new Error("Could not save holiday.");

        const payload = await response.json();
        setBoard(payload.board);
        setJobs(payload.jobs);
        setHolidays(payload.holidays);
        if (showHolidays) {
          await refreshHolidayData();
        }
        setMessage(createMessage("Holiday updated.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not save holiday.", "error"));
    }
  }

  async function deleteHoliday(holidayId, options = {}) {
    if (isClientMode) return;
    try {
      const params = new URLSearchParams();
      if (options.date) params.set("date", options.date);
      if (options.person) params.set("person", options.person);
      const url = params.toString() ? `/api/holidays/${holidayId}?${params.toString()}` : `/api/holidays/${holidayId}`;
      const response = await fetch(url, { method: "DELETE" });
      if (!response.ok) throw new Error("Could not delete holiday.");

        const payload = await response.json();
        setBoard(payload.board);
        setJobs(payload.jobs);
        setHolidays(payload.holidays);
        if (showHolidays) {
          await refreshHolidayData();
        }
        setMessage(createMessage("Holiday removed.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not delete holiday.", "error"));
    }
  }

  async function cycleHoliday(date, person) {
    if (isClientMode) return;
    const normalizedPerson = getHolidayStaffIdentityKey(person);
    const matchingEntries = holidays.filter(
      (item) =>
        String(item.date || "") === String(date || "") &&
        getHolidayStaffIdentityKey(item.person) === normalizedPerson
    );
    const existing =
      matchingEntries.find((item) => !isBirthdayHoliday(item)) ||
      matchingEntries[0] ||
      null;

    if (existing && isBirthdayHoliday(existing)) return;

    if (!existing) {
      await saveHoliday(date, person, "Full Day");
      return;
    }

    if (existing.duration === "Full Day") {
      await saveHoliday(date, person, "Morning", existing.id);
      return;
    }

    if (existing.duration === "Morning") {
      await saveHoliday(date, person, "Afternoon", existing.id);
      return;
    }

    await deleteHoliday(existing.id, { date, person });
  }

  async function submitHolidayEvent() {
    if (!canEditHolidays(currentUser)) return;
    setHolidayEventSaving(true);

    try {
      const response = await fetch("/api/holiday-events", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(holidayEventForm)
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not save calendar event.");
      setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
      setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
      setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
      setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
      setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
      setHolidayEventForm(EMPTY_HOLIDAY_EVENT_FORM);
      setHolidayEventOpen(false);
      setMessage(createMessage("Calendar event updated.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not save calendar event.", "error"));
    } finally {
      setHolidayEventSaving(false);
    }
  }

  async function deleteHolidayEvent(eventId) {
    if (!canEditHolidays(currentUser)) return;
    setHolidayEventSaving(true);

    try {
      const response = await fetch(
        `/api/holiday-events/${encodeURIComponent(eventId)}?yearStart=${encodeURIComponent(holidayYearStart)}`,
        { method: "DELETE" }
      );
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "Could not delete calendar event.");
      setHolidays(Array.isArray(payload.holidays) ? payload.holidays : []);
      setHolidayRequests(Array.isArray(payload.holidayRequests) ? payload.holidayRequests : []);
      setApprovedHolidayRequests(Array.isArray(payload.approvedHolidayRequests) ? payload.approvedHolidayRequests : []);
      setHolidayAllowances(Array.isArray(payload.holidayAllowances) ? payload.holidayAllowances : []);
      setHolidayEvents(Array.isArray(payload.holidayEvents) ? payload.holidayEvents : []);
      setHolidayEventForm(EMPTY_HOLIDAY_EVENT_FORM);
      setHolidayEventOpen(false);
      setMessage(createMessage("Calendar event removed.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage(error.message || "Could not delete calendar event.", "error"));
    } finally {
      setHolidayEventSaving(false);
    }
  }

  async function handleHolidayChipClick(date, person, closePicker = false) {
    await cycleHoliday(date, person);
    if (closePicker) {
      setActiveHolidayDate("");
    }
  }

  async function handleManualRefresh() {
    try {
      const [boardResponse] = await Promise.all([fetch("/api/board"), refreshData()]);
      if (!boardResponse.ok) throw new Error("Could not refresh the board.");
      setBoard(await boardResponse.json());
      setMessage(createMessage("Board refreshed.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not refresh the board.", "error"));
    }
  }

  if (!authChecked) {
    return (
      <div className="app-shell">
        <div className="page">
          <section className="panel auth-panel">
            <div className="board-loading">Checking login...</div>
          </section>
        </div>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div className="app-shell">
        <div className="page">
          <section className="panel auth-panel">
            <div className="auth-brand">
              <img className="auth-logo" src="/branding/signs-express-logo.svg" alt="Signs Express" />
            </div>
            <form className="job-form auth-form" onSubmit={handleLogin}>
              <label>
                Username
                <input
                  type="text"
                  value={loginDisplayName}
                  placeholder="Enter your full name"
                  onChange={(event) => setLoginDisplayName(event.target.value)}
                />
              </label>

              <label>
                Password
                <input
                  type="password"
                  value={loginPassword}
                  placeholder="Enter your password"
                  onChange={(event) => setLoginPassword(event.target.value)}
                />
              </label>

              {loginError ? <div className="flash error">{loginError}</div> : null}

              <div className="form-actions">
                <button className="primary-button" type="submit" disabled={loginLoading || !loginDisplayName || !loginPassword}>
                  {loginLoading ? "Signing in..." : "Sign in"}
                </button>
              </div>
            </form>
          </section>
        </div>
      </div>
    );
  }

  if (showHostLanding) {
    return (
        <HostLandingPage
          currentUser={currentUser}
          onLogout={handleLogout}
          users={loginUsers}
          savingKey={permissionSavingKey}
          onChangePermission={handlePermissionChange}
          onUpdateAttendanceProfile={handleUpdateAttendanceProfile}
          onCreateUser={handleCreateUser}
          onResetPassword={handleResetUserPassword}
          onDeleteUser={handleDeleteUser}
          notifications={notifications}
        />
    );
  }

  if (showClientLanding) {
    return (
      <ClientLandingPage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
      />
    );
  }

  if (showNotifications) {
    return (
      <NotificationsPage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
        onOpenNotification={openNotification}
        onMarkNotificationRead={markNotificationRead}
        onMarkAllNotificationsRead={markAllNotificationsRead}
      />
    );
  }

  if (showMileage) {
    return (
      <MileagePage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
        onRefreshNotifications={refreshNotifications}
      />
    );
  }

  if (showVanEstimator) {
    return (
      <VinylEstimatorPage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
      />
    );
  }

  if (showAttendance) {
    return (
      <AttendancePage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
        attendanceData={attendanceData}
        loading={attendanceLoading}
        attendanceMonthId={attendanceMonthId}
        setAttendanceMonthId={setAttendanceMonthId}
        attendanceSavingKey={attendanceSavingKey}
        attendanceNoteSavingKey={attendanceNoteSavingKey}
        attendanceFocusDate={attendanceNotificationDate}
        onSaveAttendanceEntry={saveAttendanceEntry}
        onSubmitAttendanceExplanation={submitAttendanceExplanation}
      />
    );
  }

  if (showHolidays) {
    return (
      <HolidaysPage
        currentUser={currentUser}
        onLogout={handleLogout}
        notifications={notifications}
        holidays={holidays}
        holidayRequests={holidayRequests}
        approvedHolidayRequests={approvedHolidayRequests}
        holidayStaff={holidayStaff}
        holidayAllowances={holidayAllowances}
        holidayEvents={holidayEvents}
        holidayRows={holidayRows}
        holidayYearStart={holidayYearStart}
        holidayYearLabel={holidayYearLabel}
        currentHolidayYearStart={currentHolidayYearStart}
        setHolidayYearStart={setHolidayYearStart}
          holidayRequestOpen={holidayRequestOpen}
          setHolidayRequestOpen={setHolidayRequestOpen}
          holidayRequestForm={holidayRequestForm}
          setHolidayRequestForm={setHolidayRequestForm}
          holidayRequestSaving={holidayRequestSaving}
          holidayCancelOpen={holidayCancelOpen}
          setHolidayCancelOpen={setHolidayCancelOpen}
          holidayCancelForm={holidayCancelForm}
          setHolidayCancelForm={setHolidayCancelForm}
          holidayEventOpen={holidayEventOpen}
          setHolidayEventOpen={setHolidayEventOpen}
          holidayEventForm={holidayEventForm}
        setHolidayEventForm={setHolidayEventForm}
        holidayEventSaving={holidayEventSaving}
        holidayAllowanceSavingKey={holidayAllowanceSavingKey}
        onChangeHolidayAllowanceDraft={changeHolidayAllowanceDraft}
        onSaveHolidayAllowance={saveHolidayAllowance}
        onToggleHolidayDate={cycleHoliday}
        onSubmitHolidayEvent={submitHolidayEvent}
          onDeleteHolidayEvent={deleteHolidayEvent}
          onSubmitHolidayRequest={submitHolidayRequest}
          onReviewHolidayRequest={reviewHolidayRequest}
          onCancelHolidayRequest={cancelHolidayRequest}
        />
      );
    }

  if (showInstallerDirectory) {
    return <InstallerDirectoryHost currentUser={currentUser} onLogout={handleLogout} readOnly={!installerEditable} />;
  }

  return (
    <div className={`app-shell ${isClientMode ? "client-mode" : "editor-mode"}`}>
      <div className="page">
        <MainNavBar
          currentUser={currentUser}
          active="board"
          onLogout={handleLogout}
          notifications={notifications}
        />

        <div className="layout">
          <section className="panel board-panel board-panel-full">
            {loading || !board ? (
              <div className="board-loading">Loading the shared installation board...</div>
            ) : (
              <div className="board board-with-history">
                <div className="board-history-launch board-history-launch-top">
                  <div className="board-history-actions">
                    <button
                      className="ghost-button board-history-button"
                      type="button"
                      onClick={() => setPreviousMonthDepth((current) => Math.min(6, current + 1))}
                    >
                      Previous months
                    </button>
                    <button className="ghost-button board-history-button" type="button" onClick={resetBoardWindow}>
                      Current month
                    </button>
                  </div>
                </div>

                {board.weeks.map((week) => (
                  <section key={week.id} className="week-block">
                    <header className="week-header">
                      <strong>{week.label}</strong>
                    </header>

                    {week.rows.map((row) => (
                      <article
                        key={row.isoDate}
                        className={[
                          "board-row",
                          row.isToday ? "is-today" : "",
                          row.bankHoliday ? "is-bank-holiday" : "",
                          row.isPast ? "is-past" : "",
                          dropDate === row.isoDate ? "is-drop-target" : ""
                        ].join(" ").trim()}
                        onDragOver={(event) => {
                          if (isClientMode) return;
                          event.preventDefault();
                          if (draggingJobId) setDropDate(row.isoDate);
                        }}
                        onDragLeave={() => {
                          if (dropDate === row.isoDate) setDropDate("");
                        }}
                        onDrop={(event) => {
                          if (isClientMode) return;
                          event.preventDefault();
                          const duplicateJobId = event.dataTransfer.getData("job-copy");
                          if (duplicateJobId || duplicatingJobId) {
                            duplicateJobToDate(duplicateJobId || duplicatingJobId, row.isoDate);
                            return;
                          }
                          const jobId = event.dataTransfer.getData("text/plain") || draggingJobId;
                          moveJobToDate(jobId, row.isoDate);
                        }}
                      >
                          <div
                            className="date-cell"
                          onClick={() => {
                            if (!isClientMode) {
                              setActiveHolidayDate((current) => (current === row.isoDate ? "" : row.isoDate));
                            }
                          }}
                          title={row.fullDateLabel}
                          >
                            <div className="date-heading">
                              <span className="date-day">{row.dayLabel}</span>
                              <strong className="date-number">{row.dayNumber}</strong>
                            </div>
                            {row.isToday ? <span className="date-today-pill">Today</span> : null}
                            {isClientMode && row.staffHolidays.length ? (
                            <div className="mobile-holiday-inline">
                              {row.staffHolidays.map((holiday) => (
                                <span
                                  key={`mobile-${holiday.id}`}
                                  className={`mobile-holiday-chip ${getHolidayPersonColor(holiday.person)} ${isBirthdayHoliday(holiday) ? "holiday-birthday-token" : ""}`}
                                >
                                  {getHolidayDisplayToken(holiday.person)}
                                  {holiday.duration === "Morning" ? " AM" : holiday.duration === "Afternoon" ? " PM" : ""}
                                </span>
                              ))}
                            </div>
                          ) : null}
                          {!isClientMode && row.staffHolidays.length ? (
                            <div className="date-holiday-summary" onClick={(event) => event.stopPropagation()}>
                              {row.staffHolidays.map((holiday) => {
                                const durationLabel =
                                  holiday.duration === "Morning"
                                    ? ".AM"
                                    : holiday.duration === "Afternoon"
                                      ? ".PM"
                                      : "";
                                return (
                                  <button
                                    key={`summary-${holiday.id}`}
                                    type="button"
                                    className={`date-holiday-chip ${getHolidayPersonColor(holiday.person)} ${isBirthdayHoliday(holiday) ? "holiday-birthday-token" : ""} active`}
                                    onClick={(event) => {
                                      event.stopPropagation();
                                      if (!isBirthdayHoliday(holiday)) {
                                        handleHolidayChipClick(row.isoDate, holiday.person, false);
                                      }
                                    }}
                                    disabled={isBirthdayHoliday(holiday)}
                                  >
                                    {getHolidayDisplayToken(holiday.person)}{durationLabel}
                                  </button>
                                );
                              })}
                            </div>
                          ) : null}
                          {!isClientMode && activeHolidayDate === row.isoDate ? (
                              <div className="date-holiday-popover" onClick={(event) => event.stopPropagation()}>
                                {holidayStaff.map((entry) => {
                                  const name = entry.person;
                                  const existing = row.staffHolidays.find((holiday) => holiday.person === name);
                                  const durationLabel =
                                    existing?.duration === "Morning"
                                      ? ".AM"
                                      : existing?.duration === "Afternoon"
                                      ? ".PM"
                                      : "";
                                const initials = getHolidayDisplayToken(name);
                                return (
                                  <button
                                    key={`${row.isoDate}-${name}`}
                                    type="button"
                                    className={`date-holiday-chip ${getHolidayPersonColor(name)} ${existing ? "active" : ""} ${isBirthdayHoliday(existing) ? "holiday-birthday-token" : ""}`}
                                    onClick={(event) => {
                                      event.stopPropagation();
                                      if (!isBirthdayHoliday(existing)) {
                                        handleHolidayChipClick(row.isoDate, name, true);
                                      }
                                    }}
                                    disabled={isBirthdayHoliday(existing)}
                                  >
                                    {initials}{durationLabel}
                                  </button>
                                );
                              })}
                            </div>
                          ) : null}
                          {!isClientMode && row.bankHoliday ? <span className="date-holiday-chip date-bank-holiday">{row.bankHoliday}</span> : null}
                          {Array.isArray(row.holidayEvents) && row.holidayEvents.length ? (
                            <div className="date-calendar-events">
                              {row.holidayEvents.map((event) => (
                                <span key={`board-event-${event.id}`} className="date-calendar-event-chip">
                                  {event.title}
                                </span>
                              ))}
                            </div>
                          ) : null}
                        </div>

                        <div className="jobs-cell">
                          <button
                            type="button"
                            className={`jobs-lane-button ${row.jobs.length === 0 ? "is-empty" : ""}`}
                            disabled={isClientMode}
                            onClick={() => {
                              if (isClientMode) return;
                              openJobModal(row.isoDate, {
                                ...EMPTY_FORM,
                                date: row.isoDate,
                                jobType: form.jobType || "Install"
                              });
                            }}
                          >
                            {row.jobs.length === 0 ? <span className="muted">No jobs booked</span> : <span className="lane-add-label">{isClientMode ? "View only" : "Click anywhere here to add another job"}</span>}
                          </button>

                          {row.jobs.length > 0 ? (
                            <div className="job-stack">
                                {row.jobs.map((job) =>
                                  renderJobCardContent({
                                    job,
                                    isCondensed: row.isPast,
                                    isClientMode,
                                    draggingJobId,
                                  getJobTypeMeta,
                                  getJobTypeLabel,
                                  getInstallerDisplayList,
                                  getInstallerMeta,
                                  editJob,
                                  handleDelete,
                                  setActiveClientJob,
                                  buildDragPreview,
                                  getTransparentDragImage,
                                  clearDragPreview,
                                  dragPreviewRef,
                                  dragPositionRef,
                                  setDraggingJobId,
                                  duplicatingJobId,
                                  setDuplicatingJobId,
                                  setDropDate
                                })
                              )}
                            </div>
                          ) : null}
                        </div>
                      </article>
                    ))}
                  </section>
                ))}

                <div className="board-history-launch board-history-launch-bottom">
                  <div className="board-history-actions">
                    <button
                      className="ghost-button board-history-button"
                      type="button"
                      onClick={() => setFutureMonthDepth((current) => Math.min(6, current + 1))}
                    >
                      Future months
                    </button>
                  </div>
                </div>

                <section
                  className={`unscheduled-section panel ${dropDate === UNSCHEDULED_DROP_ZONE ? "is-drop-target" : ""}`}
                  onDragOver={(event) => {
                    if (isClientMode) return;
                    event.preventDefault();
                    if (draggingJobId || duplicatingJobId) setDropDate(UNSCHEDULED_DROP_ZONE);
                  }}
                  onDragLeave={() => {
                    if (dropDate === UNSCHEDULED_DROP_ZONE) setDropDate("");
                  }}
                  onDrop={(event) => {
                    if (isClientMode) return;
                    event.preventDefault();
                    const duplicateJobId = event.dataTransfer.getData("job-copy");
                    if (duplicateJobId || duplicatingJobId) {
                      duplicateJobToDate(duplicateJobId || duplicatingJobId, UNSCHEDULED_DROP_ZONE);
                      return;
                    }
                    const jobId = event.dataTransfer.getData("text/plain") || draggingJobId;
                    moveJobToDate(jobId, UNSCHEDULED_DROP_ZONE);
                  }}
                >
                  <div className="unscheduled-head">
                    <div>
                      <h3>Unscheduled</h3>
                      <p>Keep jobs here until you have a firm date, then drag them into the calendar or edit the date.</p>
                    </div>
                    {!isClientMode ? (
                      <button
                        className="ghost-button"
                        type="button"
                        onClick={() => {
                          openJobModal("Unscheduled", {
                            ...EMPTY_FORM,
                            date: "",
                            jobType: form.jobType || "Install"
                          });
                        }}
                      >
                        Add unscheduled job
                      </button>
                    ) : null}
                  </div>

                  {board.unscheduled?.length ? (
                    <div className="job-stack unscheduled-job-stack">
                        {board.unscheduled.map((job) =>
                          renderJobCardContent({
                            job,
                            isCondensed: false,
                            isClientMode,
                          draggingJobId,
                          getJobTypeMeta,
                          getJobTypeLabel,
                          getInstallerDisplayList,
                          getInstallerMeta,
                          editJob,
                          handleDelete,
                          setActiveClientJob,
                          buildDragPreview,
                          getTransparentDragImage,
                          clearDragPreview,
                          dragPreviewRef,
                          dragPositionRef,
                          setDraggingJobId,
                          duplicatingJobId,
                          setDuplicatingJobId,
                          setDropDate
                        })
                      )}
                    </div>
                  ) : (
                    <div className="unscheduled-empty">No unscheduled jobs.</div>
                  )}
                </section>
              </div>
            )}
          </section>
        </div>
      </div>
      {!isClientMode && jobModalDate ? (
        <div
          className="modal-backdrop"
          onPointerDown={handleBackdropPointerDown}
          onClick={(event) => handleBackdropClick(event, () => resetForm())}
        >
          <div className="modal job-modal" onPointerDown={() => { backdropPointerStartedRef.current = false; }} onClick={(event) => event.stopPropagation()}>
            <div className="modal-head job-modal-head">
              <button className="icon-button" type="button" onClick={() => resetForm()}>
                x
              </button>
            </div>

            <form className="job-form job-form-scroll" onSubmit={handleSubmit}>
              <label>
                Date
                <input
                  type="date"
                  value={form.date}
                  onChange={(event) => setForm((current) => ({ ...current, date: event.target.value }))}
                />
              </label>

              <div className="corebridge-lookup-bar">
                <button className="ghost-button" type="button" onClick={() => openOrderLookup()}>
                  Find order
                </button>
                <span className="muted">Pull live order details from CoreBridge where available.</span>
              </div>

              <label>
                Order reference
                <div className="order-reference-row">
                  <input
                    type="text"
                    value={form.orderReference}
                    onChange={(event) => setForm((current) => ({ ...current, orderReference: event.target.value }))}
                  />
                  <button
                    type="button"
                    className={`placeholder-toggle-button ${form.isPlaceholder ? "active" : ""}`}
                    onClick={() => setForm((current) => ({ ...current, isPlaceholder: !current.isPlaceholder }))}
                  >
                    Add as Placeholder
                  </button>
                </div>
              </label>

              <label>
                Customer name
                <input
                  type="text"
                  value={form.customerName}
                  onChange={(event) => setForm((current) => ({ ...current, customerName: event.target.value }))}
                />
              </label>

              <label>
                Description
                <input
                  type="text"
                  value={form.description}
                  onChange={(event) => setForm((current) => ({ ...current, description: event.target.value }))}
                />
              </label>

              <div className="split-fields">
                <label>
                  Contact
                  <input
                    type="text"
                    value={form.contact}
                    onChange={(event) => setForm((current) => ({ ...current, contact: event.target.value }))}
                  />
                </label>

                <label>
                  Number
                  <input
                    type="text"
                    value={form.number}
                    onChange={(event) => setForm((current) => ({ ...current, number: event.target.value }))}
                  />
                </label>
              </div>

              <label>
                Address
                <input
                  type="text"
                  value={form.address}
                  onChange={(event) => setForm((current) => ({ ...current, address: event.target.value }))}
                />
              </label>

              <label>
                Installers
                <div className="installer-picker">
                  {INSTALLER_OPTIONS.map((option) => (
                    <button
                      key={option.value}
                      type="button"
                      className={`installer-chip ${option.colorClass} ${form.installers.includes(option.value) ? "active" : ""}`}
                      onClick={() => toggleInstaller(option.value)}
                    >
                      {option.value}
                    </button>
                  ))}
                </div>
              </label>

              {form.installers.includes("Custom") ? (
                <label>
                  Custom installer
                  <input
                    type="text"
                    value={form.customInstaller}
                    onChange={(event) => setForm((current) => ({ ...current, customInstaller: event.target.value }))}
                  />
                </label>
              ) : null}

              <label>
                Job type
                <select
                  value={form.jobType}
                  onChange={(event) => setForm((current) => ({ ...current, jobType: event.target.value }))}
                >
                  {JOB_TYPES.map((option) => (
                    <option key={option.value} value={option.value}>
                      {option.value}
                    </option>
                  ))}
                </select>
              </label>

              {form.jobType === "Other" ? (
                <label>
                  Other job type
                  <input
                    type="text"
                    value={form.customJobType}
                    onChange={(event) => setForm((current) => ({ ...current, customJobType: event.target.value }))}
                  />
                </label>
              ) : null}

              <label>
                Notes
                <textarea
                  rows="4"
                  value={form.notes}
                  onChange={(event) => setForm((current) => ({ ...current, notes: event.target.value }))}
                />
              </label>

              {activeAdminJob ? (
                <>
                  <div className="client-job-summary admin-job-summary">
                    {activeAdminJob.orderReference ? <span className="job-summary-pill">{activeAdminJob.orderReference}</span> : null}
                    <span className={`job-summary-pill ${activeAdminJob.isCompleted ? "is-complete" : ""}`}>
                      {activeAdminJob.isCompleted ? "Completed" : getJobTypeLabel(activeAdminJob)}
                    </span>
                    {activeAdminJob.isSnagging ? <span className="job-summary-pill is-snagging">Snagging</span> : null}
                    {activeAdminJob.isPlaceholder ? <span className="job-summary-pill is-placeholder">Placeholder</span> : null}
                    {Array.isArray(activeAdminJob.photos) && activeAdminJob.photos.length ? (
                      <span className="job-summary-pill is-photos">{activeAdminJob.photos.length} photo{activeAdminJob.photos.length === 1 ? "" : "s"}</span>
                    ) : null}
                  </div>

                  <div className="detail-grid client-detail-grid admin-job-detail-grid">
                    {activeAdminJob.completedAt ? (
                      <div className="detail-card">
                        <strong>Completed At</strong>
                        <p>{formatDateTime(activeAdminJob.completedAt)}</p>
                      </div>
                    ) : null}
                    {activeAdminJob.completedByName ? (
                      <div className="detail-card">
                        <strong>Completed By</strong>
                        <p>{activeAdminJob.completedByName}</p>
                      </div>
                    ) : null}
                    <div className="detail-card detail-card-wide">
                      <strong>Photos</strong>
                      {Array.isArray(activeAdminJob.photos) && activeAdminJob.photos.length ? (
                        <div className="job-photo-grid">
                          {activeAdminJob.photos.map((photo) => (
                            <div key={photo.id} className="job-photo-tile">
                              <a
                                className="job-photo-link"
                                href={photo.url || buildJobPhotoUrl(activeAdminJob.id, photo.id)}
                                target="_blank"
                                rel="noreferrer"
                                onClick={(event) => event.stopPropagation()}
                              >
                                <img
                                  src={photo.url || buildJobPhotoUrl(activeAdminJob.id, photo.id)}
                                  alt={photo.fileName || "Job photo"}
                                  loading="lazy"
                                />
                              </a>
                              <div className="job-photo-meta">
                                <small>{photo.uploadedByName || "Uploaded photo"}</small>
                                <button
                                  className="text-button danger"
                                  type="button"
                                  onClick={() => {
                                    deleteJobPhoto(activeAdminJob, photo.id);
                                  }}
                                >
                                  Delete
                                </button>
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <p>No photos uploaded yet.</p>
                      )}
                    </div>
                  </div>

                  </>
                ) : null}

              <div className="form-actions job-form-actions">
                {activeAdminJob && !activeAdminJob.isCompleted && !adminCompletePrompt ? (
                  <button
                    className="success-button"
                    type="button"
                    onClick={() => setAdminCompletePrompt(true)}
                    disabled={adminPhotoUploading || adminExporting}
                  >
                    Mark as Complete
                  </button>
                ) : null}
                {activeAdminJob && !activeAdminJob.isCompleted && !activeAdminJob.isSnagging && !adminCompletePrompt ? (
                  <button
                    className="snagging-button"
                    type="button"
                    onClick={() => markAdminJobSnagging(activeAdminJob)}
                    disabled={adminPhotoUploading || adminExporting}
                  >
                    Snagging
                  </button>
                ) : null}
                {activeAdminJob && activeAdminJob.isSnagging && !adminCompletePrompt ? (
                  <button
                    className="snagging-button is-active"
                    type="button"
                    onClick={() => removeAdminJobSnagging(activeAdminJob)}
                    disabled={adminPhotoUploading || adminExporting}
                  >
                    Remove Snagging
                  </button>
                ) : null}
                <button className="primary-button" type="submit" disabled={saving}>
                  {saving ? "Saving..." : editingId ? "Update job" : "Add job"}
                </button>
                <button className="ghost-button" type="button" onClick={() => resetForm()}>
                  Cancel
                </button>
                {activeAdminJob?.isCompleted ? (
                  <>
                    <button
                      className="ghost-button"
                      type="button"
                      onClick={() => adminPhotoInputRef.current?.click()}
                      disabled={adminPhotoUploading}
                    >
                      {adminPhotoUploading ? "Uploading..." : "Upload photos"}
                    </button>
                    <button
                      className="ghost-button"
                      type="button"
                      onClick={() => undoAdminJobComplete(activeAdminJob)}
                      disabled={adminPhotoUploading || adminExporting}
                    >
                      Undo complete
                    </button>
                  </>
                ) : null}
                {activeAdminJob ? (
                  <button
                    className="ghost-button"
                    type="button"
                    onClick={() => exportJob(activeAdminJob, setAdminExporting)}
                    disabled={adminExporting || adminPhotoUploading}
                  >
                    {adminExporting ? "Exporting..." : "Export"}
                  </button>
                ) : null}
                {activeAdminJob && !activeAdminJob.isCompleted && adminCompletePrompt ? (
                  <div className="client-complete-prompt form-complete-prompt">
                    <span>Would you like to upload job photos?</span>
                    <div className="client-complete-prompt-actions">
                      <button
                        className="ghost-button"
                        type="button"
                        onClick={() => markAdminJobComplete(activeAdminJob)}
                        disabled={adminPhotoUploading}
                      >
                        No
                      </button>
                      <button
                        className="success-button"
                        type="button"
                        onClick={() => adminPhotoInputRef.current?.click()}
                        disabled={adminPhotoUploading}
                      >
                        {adminPhotoUploading ? "Uploading..." : "Yes"}
                      </button>
                    </div>
                  </div>
                ) : null}
              </div>
            </form>
            <input
              ref={adminPhotoInputRef}
              className="visually-hidden"
              type="file"
              accept="image/*"
              multiple
              onChange={async (event) => {
                const files = Array.from(event.target.files || []);
                if (!files.length || !activeAdminJob) return;
                await markAdminJobComplete(activeAdminJob, files);
              }}
            />
          </div>
        </div>
      ) : null}
      {!isClientMode && orderLookupOpen ? (
        <div
          className="modal-backdrop"
          onPointerDown={handleBackdropPointerDown}
          onClick={(event) => handleBackdropClick(event, () => setOrderLookupOpen(false))}
        >
          <div className="modal order-lookup-modal" onPointerDown={() => { backdropPointerStartedRef.current = false; }} onClick={(event) => event.stopPropagation()}>
            <div className="modal-head">
              <div>
                <h3>Find CoreBridge Order</h3>
                <p>Search live orders and copy the matching details into this job.</p>
              </div>
              <button className="icon-button" type="button" onClick={() => setOrderLookupOpen(false)}>
                x
              </button>
            </div>

            <div className="order-lookup-toolbar">
              <input
                type="text"
                value={orderLookupQuery}
                placeholder="Search by order ref, customer, phone or address"
                onChange={(event) => setOrderLookupQuery(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter") {
                    event.preventDefault();
                    searchCoreBridgeOrders(orderLookupQuery);
                  }
                }}
              />
              <button className="primary-button" type="button" onClick={() => searchCoreBridgeOrders(orderLookupQuery)} disabled={orderLookupLoading}>
                {orderLookupLoading ? "Searching..." : "Search"}
              </button>
            </div>

            {orderLookupError ? <div className="flash error">{orderLookupError}</div> : null}

            <div className="order-lookup-results">
              {orderLookupLoading ? (
                <div className="board-loading compact">Looking up CoreBridge orders...</div>
              ) : orderLookupResults.length ? (
                orderLookupResults.map((order) => (
                  <div
                    key={`${order.id}-${order.orderReference}-${order.customerName}`}
                    className="order-result-card"
                  >
                    <div className="order-result-top">
                      <strong>{order.orderReference || "No order ref"}</strong>
                      {order.status ? <span className="job-tag job-type-other">{order.status}</span> : null}
                    </div>
                    <p className="order-result-customer">{order.customerName || "Unnamed customer"}</p>
                    <p>{order.description || "No description"}</p>
                    <div className="order-result-meta">
                      <span><b>Contact:</b> {order.contact || "-"}</span>
                      <span><b>Number:</b> {order.number || "-"}</span>
                    </div>
                    <p className="order-result-address"><b>Address:</b> {order.address || "-"}</p>
                    <div className="order-result-actions">
                      <button className="primary-button" type="button" onClick={() => applyCoreBridgeOrder(order)}>
                        Use this order
                      </button>
                    </div>
                  </div>
                ))
              ) : (
                <div className="board-loading compact">No CoreBridge orders found yet.</div>
              )}
            </div>
          </div>
        </div>
      ) : null}
      {isClientMode && activeClientJob ? (
        <div
          className="modal-backdrop"
          onPointerDown={handleBackdropPointerDown}
          onClick={(event) => handleBackdropClick(event, () => setActiveClientJob(null))}
        >
          <div className="modal client-detail-modal" onPointerDown={() => { backdropPointerStartedRef.current = false; }} onClick={(event) => event.stopPropagation()}>
            <div className="modal-head">
              <div>
                <h3>{activeClientJob.customerName}</h3>
                <p>{activeClientJob.description || "No description"}</p>
              </div>
              <button className="icon-button" type="button" onClick={() => setActiveClientJob(null)}>
                x
              </button>
            </div>
            <div className="client-detail-scroll">
              <div className="client-job-summary">
                {activeClientJob.orderReference ? <span className="job-summary-pill">{activeClientJob.orderReference}</span> : null}
                <span className={`job-summary-pill ${activeClientJob.isCompleted ? "is-complete" : ""}`}>
                  {activeClientJob.isCompleted ? "Completed" : getJobTypeLabel(activeClientJob)}
                </span>
                {activeClientJob.isSnagging ? <span className="job-summary-pill is-snagging">Snagging</span> : null}
                {activeClientJob.isPlaceholder ? <span className="job-summary-pill is-placeholder">Placeholder</span> : null}
                {Array.isArray(activeClientJob.photos) && activeClientJob.photos.length ? (
                  <span className="job-summary-pill is-photos">{activeClientJob.photos.length} photo{activeClientJob.photos.length === 1 ? "" : "s"}</span>
                ) : null}
              </div>
              <div className="detail-grid client-detail-grid">
                <div className="detail-card">
                  <strong>Contact</strong>
                  <p>{activeClientJob.contact || "-"}</p>
                </div>
                <div className="detail-card">
                  <strong>Number</strong>
                  <p>{activeClientJob.number || "-"}</p>
                </div>
                <div className="detail-card">
                  <strong>Installers</strong>
                  <p>{getInstallerDisplayList(activeClientJob).join(", ") || "-"}</p>
                </div>
                <div className="detail-card">
                  <strong>Placeholder</strong>
                  <p>{activeClientJob.isPlaceholder ? "Yes" : "No"}</p>
                </div>
                {activeClientJob.completedAt ? (
                  <div className="detail-card">
                    <strong>Completed At</strong>
                    <p>{formatDateTime(activeClientJob.completedAt)}</p>
                  </div>
                ) : null}
                {activeClientJob.completedByName ? (
                  <div className="detail-card">
                    <strong>Completed By</strong>
                    <p>{activeClientJob.completedByName}</p>
                  </div>
                ) : null}
                <div className="detail-card detail-card-wide">
                  <strong>Address</strong>
                  <p>{activeClientJob.address || "-"}</p>
                </div>
                <div className="detail-card detail-card-wide">
                  <strong>Notes</strong>
                  <p>{activeClientJob.notes || "-"}</p>
                </div>
                <div className="detail-card detail-card-wide">
                  <strong>Photos</strong>
                  {Array.isArray(activeClientJob.photos) && activeClientJob.photos.length ? (
                    <div className="job-photo-grid">
                      {activeClientJob.photos.map((photo) => (
                        <div key={photo.id} className="job-photo-tile">
                          <a
                            className="job-photo-link"
                            href={photo.url || buildJobPhotoUrl(activeClientJob.id, photo.id)}
                            target="_blank"
                            rel="noreferrer"
                            onClick={(event) => event.stopPropagation()}
                          >
                            <img
                              src={photo.url || buildJobPhotoUrl(activeClientJob.id, photo.id)}
                              alt={photo.fileName || "Job photo"}
                              loading="lazy"
                            />
                          </a>
                          <div className="job-photo-meta">
                            <small>{photo.uploadedByName || "Uploaded photo"}</small>
                            <button
                              className="text-button danger"
                              type="button"
                              onClick={(event) => {
                                event.stopPropagation();
                                deleteJobPhoto(activeClientJob, photo.id);
                              }}
                            >
                              Delete
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <p>No photos uploaded yet.</p>
                  )}
                </div>
              </div>
              <div className="client-job-actions">
                {!activeClientJob.isCompleted ? (
                  <>
                    {!clientCompletePrompt ? (
                      <button
                        className="primary-button"
                        type="button"
                        onClick={() => setClientCompletePrompt(true)}
                        disabled={clientPhotoUploading || clientExporting}
                      >
                        Mark as Complete
                      </button>
                    ) : (
                      <div className="client-complete-prompt">
                        <span>Would you like to upload job photos?</span>
                        <div className="client-complete-prompt-actions">
                          <button
                            className="ghost-button"
                            type="button"
                            onClick={() => markClientJobComplete(activeClientJob)}
                            disabled={clientPhotoUploading}
                          >
                            No
                          </button>
                          <button
                            className="primary-button"
                            type="button"
                            onClick={() => clientPhotoInputRef.current?.click()}
                            disabled={clientPhotoUploading}
                          >
                            {clientPhotoUploading ? "Uploading..." : "Yes"}
                          </button>
                        </div>
                      </div>
                    )}
                  </>
                ) : (
                  <>
                    <button
                      className="ghost-button"
                      type="button"
                      onClick={() => clientPhotoInputRef.current?.click()}
                      disabled={clientPhotoUploading}
                    >
                      {clientPhotoUploading ? "Uploading..." : "Upload photos"}
                    </button>
                    <button
                      className="ghost-button"
                      type="button"
                      onClick={() => undoClientJobComplete(activeClientJob)}
                      disabled={clientPhotoUploading || clientExporting}
                    >
                      Undo complete
                    </button>
                  </>
                )}

                <button
                  className="ghost-button"
                  type="button"
                  onClick={() => exportJob(activeClientJob, setClientExporting)}
                  disabled={clientExporting || clientPhotoUploading}
                >
                  {clientExporting ? "Exporting..." : "Export"}
                </button>
              </div>
            </div>
            <input
              ref={clientPhotoInputRef}
              className="visually-hidden"
              type="file"
              accept="image/*"
              multiple
              onChange={async (event) => {
                const files = Array.from(event.target.files || []);
                if (!files.length) return;
                await markClientJobComplete(activeClientJob, files);
              }}
            />
          </div>
        </div>
      ) : null}
    </div>
  );
}
