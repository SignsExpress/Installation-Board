import { useEffect, useMemo, useRef, useState } from "react";

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

const STAFF_NAMES = ["Matt R", "Dawn D", "Tom V-B", "Amber H", "Eddy D'A", "Paul M", "Kyle W", "Matt C", "Keilan C"];
const HOLIDAY_DURATIONS = ["Morning", "Afternoon", "Full Day"];
const HOLIDAY_PERSON_COLORS = {
  "Matt R": "holiday-person-black",
  "Dawn D": "holiday-person-black",
  "Tom V-B": "holiday-person-black",
  "Amber H": "holiday-person-black",
  "Eddy D'A": "holiday-person-black",
  "Paul M": "holiday-person-green",
  "Kyle W": "holiday-person-green",
  "Matt C": "holiday-person-red",
  "Keilan C": "holiday-person-red"
};

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
  notes: ""
};

function getLocalTodayIso() {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/London",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).format(new Date());
}

function createMessage(text, tone = "info") {
  return { text, tone, id: `${Date.now()}-${Math.random()}` };
}

function toInitials(name) {
  return String(name || "")
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part[0])
    .join("");
}

export default function App() {
  const pathname = typeof window !== "undefined" ? window.location.pathname.toLowerCase() : "/";
  const isClientMode = pathname.startsWith("/client");
  const [board, setBoard] = useState(null);
  const [jobs, setJobs] = useState([]);
  const [holidays, setHolidays] = useState([]);
  const [form, setForm] = useState({ ...EMPTY_FORM, date: getLocalTodayIso() });
  const [editingId, setEditingId] = useState("");
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState(null);
  const [draggingJobId, setDraggingJobId] = useState("");
  const [draggingHolidayId, setDraggingHolidayId] = useState("");
  const [dropDate, setDropDate] = useState("");
  const [activeHolidayDate, setActiveHolidayDate] = useState("");
  const [activeHolidayId, setActiveHolidayId] = useState("");
  const [holidayForm, setHolidayForm] = useState({ person: STAFF_NAMES[0], duration: "Full Day" });
  const [jobModalDate, setJobModalDate] = useState("");
  const [activeClientJob, setActiveClientJob] = useState(null);
  const dragPreviewRef = useRef(null);
  const transparentDragImageRef = useRef(null);
  const dragPositionRef = useRef({ x: 0, y: 0 });

  useEffect(() => {
    let active = true;

    async function loadBoard() {
      try {
        setLoading(true);
        const [boardResponse, jobsResponse, holidaysResponse] = await Promise.all([
          fetch("/api/board"),
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
  }, []);

  useEffect(() => {
    const stream = new EventSource("/api/events");

    function handleUpdate(event) {
      try {
        setBoard(JSON.parse(event.data));
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
  }, []);

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

  function resetForm(nextDate = board?.today || getLocalTodayIso()) {
    setEditingId("");
    setForm({ ...EMPTY_FORM, date: nextDate });
    setJobModalDate("");
  }

  function editJob(job) {
    setEditingId(job.id);
    setJobModalDate(job.date);
    setForm({
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
      notes: job.notes || ""
    });
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function refreshData() {
    const [jobsResponse, holidaysResponse] = await Promise.all([fetch("/api/jobs"), fetch("/api/holidays")]);
    if (!jobsResponse.ok || !holidaysResponse.ok) throw new Error("Could not refresh data.");
    const [nextJobs, nextHolidays] = await Promise.all([jobsResponse.json(), holidaysResponse.json()]);
    setJobs(Array.isArray(nextJobs) ? nextJobs : []);
    setHolidays(Array.isArray(nextHolidays) ? nextHolidays : []);
  }

  async function handleSubmit(event) {
    event.preventDefault();
    if (isClientMode) return;

    if (!form.date || !form.customerName.trim()) {
      setMessage(createMessage("Pick a date and enter the customer name.", "error"));
      return;
    }

    setSaving(true);

    try {
      const response = await fetch("/api/jobs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
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
    if (!job || !nextDate || job.date === nextDate) {
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
          date: nextDate
        })
      });

      if (!response.ok) throw new Error("Could not move the job.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      if (editingId === jobId) {
        setForm((current) => ({ ...current, date: nextDate }));
      }
      setMessage(createMessage(`Job moved to ${nextDate}.`, "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not move the job.", "error"));
    } finally {
      setDropDate("");
      setDraggingJobId("");
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

  async function saveHoliday(date) {
    if (isClientMode) return;
    try {
      const response = await fetch("/api/holidays", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          date,
          person: holidayForm.person,
          duration: holidayForm.duration
        })
      });
      if (!response.ok) throw new Error("Could not save holiday.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      setMessage(createMessage("Holiday added.", "success"));
      setHolidayForm({ person: STAFF_NAMES[0], duration: "Full Day" });
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not save holiday.", "error"));
    }
  }

  async function deleteHoliday(holidayId) {
    if (isClientMode) return;
    try {
      const response = await fetch(`/api/holidays/${holidayId}`, { method: "DELETE" });
      if (!response.ok) throw new Error("Could not delete holiday.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      setMessage(createMessage("Holiday removed.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not delete holiday.", "error"));
    }
  }

  async function duplicateHolidayToDate(holidayId, nextDate) {
    if (isClientMode) return;
    const holiday = holidaysById.get(holidayId);
    if (!holiday || !nextDate) {
      setDraggingHolidayId("");
      setDropDate("");
      return;
    }

    try {
      const response = await fetch("/api/holidays", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          date: nextDate,
          person: holiday.person,
          duration: holiday.duration
        })
      });
      if (!response.ok) throw new Error("Could not duplicate holiday.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      setMessage(createMessage("Holiday copied.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not duplicate holiday.", "error"));
    } finally {
      setDraggingHolidayId("");
      setDropDate("");
      clearDragPreview();
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

  return (
    <div className={`app-shell ${isClientMode ? "client-mode" : "editor-mode"}`}>
      <div className="page">
        <section className="hero">
          <div className="hero-brand">
            <img className="hero-logo" src="/branding/signs-express-logo.svg" alt="Signs Express" />
            <div className="hero-copy">
              <h1>Installation Board</h1>
            </div>
          </div>
        </section>

        <div className="layout">
          <section className="panel board-panel board-panel-full">
            <div className="panel-head">
              <div>
                <p className="panel-kicker">Live board</p>
                <h2>{isClientMode ? "Rolling installation calendar - View only" : "Rolling installation calendar"}</h2>
              </div>
              <div className="board-range">{board ? `${board.start} to ${board.end}` : "Loading"}</div>
            </div>

            {loading || !board ? (
              <div className="board-loading">Loading the shared installation board...</div>
            ) : (
              <div className="board">
                {board.weeks.map((week) => (
                  <section key={week.id} className="week-block">
                    <header className="week-header">
                      <span>Week</span>
                      <strong>{week.label}</strong>
                    </header>

                    <div className="board-header">
                      <div>Date</div>
                      <div>Holidays</div>
                      <div>Jobs</div>
                    </div>

                    {week.rows.map((row) => (
                      <article
                        key={row.isoDate}
                        className={[
                          "board-row",
                          row.isToday ? "is-today" : "",
                          row.bankHoliday || row.staffHolidays.length ? "is-holiday" : "",
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
                          const holidayId = event.dataTransfer.getData("holiday-copy");
                          if (holidayId || draggingHolidayId) {
                            duplicateHolidayToDate(holidayId || draggingHolidayId, row.isoDate);
                            return;
                          }
                          const jobId = event.dataTransfer.getData("text/plain") || draggingJobId;
                          moveJobToDate(jobId, row.isoDate);
                        }}
                      >
                        <button
                          type="button"
                          className="date-cell"
                          onClick={() => setForm((current) => ({ ...current, date: row.isoDate }))}
                          title={row.fullDateLabel}
                        >
                          <span className="date-day">{row.dayLabel}</span>
                          <strong className="date-number">{row.dayNumber}</strong>
                          {isClientMode && row.staffHolidays.length ? (
                            <div className="mobile-holiday-inline">
                              {row.staffHolidays.map((holiday) => (
                                <span key={`mobile-${holiday.id}`} className={`mobile-holiday-chip ${getHolidayPersonColor(holiday.person)}`}>
                                  {toInitials(holiday.person)}
                                  {holiday.duration === "Morning" ? " AM" : holiday.duration === "Afternoon" ? " PM" : ""}
                                </span>
                              ))}
                            </div>
                          ) : null}
                        </button>

                        <div className="holiday-cell" onClick={() => setActiveHolidayDate((current) => (current === row.isoDate ? "" : row.isoDate))}>
                          <div className="holiday-stack">
                            {row.bankHoliday ? <span className="holiday-pill">{row.bankHoliday}</span> : null}
                            {row.staffHolidays.length ? (
                              <div className="holiday-grid">
                                {row.staffHolidays.map((holiday) => (
                                  (() => {
                                    const durationLabel =
                                      holiday.duration === "Morning" ? "AM" : holiday.duration === "Afternoon" ? "PM" : "";
                                    return (
                                  <button
                                    key={holiday.id}
                                    type="button"
                                    className={`staff-holiday-pill ${getHolidayPersonColor(holiday.person)} ${activeHolidayId === holiday.id ? "active" : ""}`}
                                    onClick={(event) => {
                                      if (isClientMode) return;
                                      event.stopPropagation();
                                      setActiveHolidayDate(row.isoDate);
                                      setActiveHolidayId((current) => (current === holiday.id ? "" : holiday.id));
                                    }}
                                  >
                                    <span>{holiday.person}</span>
                                    {durationLabel ? <b>{durationLabel}</b> : null}
                                    <span
                                      className="holiday-duplicate"
                                      draggable={!isClientMode}
                                      onDragStart={(event) => {
                                        if (isClientMode) return;
                                        event.stopPropagation();
                                        event.dataTransfer.setData("holiday-copy", holiday.id);
                                        event.dataTransfer.effectAllowed = "copy";
                                        const preview = buildDragPreview(event.currentTarget.closest(".staff-holiday-pill"));
                                        dragPreviewRef.current = preview;
                                        dragPositionRef.current = { x: event.clientX, y: event.clientY };
                                        preview.style.left = `${event.clientX + 18}px`;
                                        preview.style.top = `${event.clientY + 18}px`;
                                        event.dataTransfer.setDragImage(getTransparentDragImage(), 0, 0);
                                        setDraggingHolidayId(holiday.id);
                                      }}
                                      onDragEnd={() => {
                                        if (isClientMode) return;
                                        setDraggingHolidayId("");
                                        setDropDate("");
                                        clearDragPreview();
                                      }}
                                      title="Drag to copy"
                                    >
                                      +
                                    </span>
                                  </button>
                                    );
                                  })()
                                ))}
                              </div>
                            ) : null}
                            {!row.bankHoliday && row.staffHolidays.length === 0 ? <span className="muted">{isClientMode ? "-" : "Click to add holiday"}</span> : null}
                            {!isClientMode && activeHolidayDate === row.isoDate && activeHolidayId ? (
                              <div className="holiday-entry-actions" onClick={(event) => event.stopPropagation()}>
                                <button className="danger-inline-button" type="button" onClick={() => { deleteHoliday(activeHolidayId); setActiveHolidayId(""); }}>
                                  Delete Selected
                                </button>
                              </div>
                            ) : null}
                            {!isClientMode && activeHolidayDate === row.isoDate ? (
                              <div className="holiday-editor" onClick={(event) => event.stopPropagation()}>
                                <label>
                                  Name
                                  <select value={holidayForm.person} onChange={(event) => setHolidayForm((current) => ({ ...current, person: event.target.value }))}>
                                    {STAFF_NAMES.map((name) => (
                                      <option key={name} value={name}>
                                        {name}
                                      </option>
                                    ))}
                                  </select>
                                </label>
                                <div className="duration-group">
                                  {HOLIDAY_DURATIONS.map((duration) => (
                                    <button
                                      key={duration}
                                      type="button"
                                      className={`duration-button ${holidayForm.duration === duration ? "active" : ""}`}
                                      onClick={() => setHolidayForm((current) => ({ ...current, duration }))}
                                    >
                                      {duration}
                                    </button>
                                  ))}
                                </div>
                                <button className="ghost-button holiday-save" type="button" onClick={() => saveHoliday(row.isoDate)}>
                                  Add Holiday
                                </button>
                              </div>
                            ) : null}
                          </div>
                        </div>

                        <div className="jobs-cell">
                          <button
                            type="button"
                            className={`jobs-lane-button ${row.jobs.length === 0 ? "is-empty" : ""}`}
                            disabled={isClientMode}
                            onClick={() => {
                              if (isClientMode) return;
                              setJobModalDate(row.isoDate);
                              setForm((current) => ({
                                ...EMPTY_FORM,
                                date: row.isoDate,
                                jobType: current.jobType || "Install"
                              }));
                              setEditingId("");
                            }}
                          >
                            {row.jobs.length === 0 ? <span className="muted">No jobs booked</span> : <span className="lane-add-label">{isClientMode ? "View only" : "Click anywhere here to add another job"}</span>}
                          </button>

                          {row.jobs.length > 0 ? (
                            <div className="job-stack">
                              {row.jobs.map((job) => {
                                const meta = getJobTypeMeta(job.jobType);
                                return (
                                  <div
                                    key={job.id}
                                    className={`job-card ${meta.colorClass}-card ${draggingJobId === job.id ? "is-dragging" : ""}`}
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
                                    {(() => {
                                      const installerLabels = getInstallerDisplayList(job);
                                      return (
                                        <>
                                    <div className="job-card-top">
                                      <div className="job-title-wrap">
                                        <strong className="job-title-line">
                                          {job.orderReference ? <span className="job-ref-inline">{job.orderReference}</span> : null}
                                          <span className="job-customer-inline">{job.customerName}</span>
                                        </strong>
                                        <p>{job.description || "No description"}</p>
                                      </div>
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
                                        </>
                                      ) : null}
                                    </div>
                                        </>
                                      );
                                    })()}
                                  </div>
                                );
                              })}
                            </div>
                          ) : null}
                        </div>
                      </article>
                    ))}
                  </section>
                ))}
              </div>
            )}
          </section>
        </div>
      </div>
      {!isClientMode && jobModalDate ? (
        <div className="modal-backdrop" onClick={() => resetForm()}>
          <div className="modal job-modal" onClick={(event) => event.stopPropagation()}>
            <div className="modal-head">
              <div>
                <h3>{editingId ? "Edit Job" : "Add Job"}</h3>
                <p>{jobModalDate}</p>
              </div>
              <button className="icon-button" type="button" onClick={() => resetForm()}>
                x
              </button>
            </div>

            <form className="job-form" onSubmit={handleSubmit}>
              <label>
                Date
                <input
                  type="date"
                  value={form.date}
                  onChange={(event) => setForm((current) => ({ ...current, date: event.target.value }))}
                />
              </label>

              <label>
                Order reference
                <input
                  type="text"
                  value={form.orderReference}
                  onChange={(event) => setForm((current) => ({ ...current, orderReference: event.target.value }))}
                />
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

              <div className="form-actions">
                <button className="primary-button" type="submit" disabled={saving}>
                  {saving ? "Saving..." : editingId ? "Update job" : "Add job"}
                </button>
                <button className="ghost-button" type="button" onClick={() => resetForm()}>
                  Cancel
                </button>
              </div>
            </form>
          </div>
        </div>
      ) : null}
      {isClientMode && activeClientJob ? (
        <div className="modal-backdrop" onClick={() => setActiveClientJob(null)}>
          <div className="modal client-detail-modal" onClick={(event) => event.stopPropagation()}>
            <div className="modal-head">
              <div>
                <h3>{activeClientJob.customerName}</h3>
                <p>{activeClientJob.description || "No description"}</p>
              </div>
              <button className="icon-button" type="button" onClick={() => setActiveClientJob(null)}>
                x
              </button>
            </div>
            <div className="detail-grid">
              <div className="detail-card">
                <strong>Order Ref</strong>
                <p>{activeClientJob.orderReference || "-"}</p>
              </div>
              <div className="detail-card">
                <strong>Job Type</strong>
                <p>{getJobTypeLabel(activeClientJob)}</p>
              </div>
              <div className="detail-card detail-card-wide">
                <strong>Installers</strong>
                <p>{getInstallerDisplayList(activeClientJob).join(", ") || "-"}</p>
              </div>
              <div className="detail-card">
                <strong>Contact</strong>
                <p>{activeClientJob.contact || "-"}</p>
              </div>
              <div className="detail-card">
                <strong>Number</strong>
                <p>{activeClientJob.number || "-"}</p>
              </div>
              <div className="detail-card detail-card-wide">
                <strong>Address</strong>
                <p>{activeClientJob.address || "-"}</p>
              </div>
              <div className="detail-card detail-card-wide">
                <strong>Notes</strong>
                <p>{activeClientJob.notes || "-"}</p>
              </div>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}
