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

const STAFF_NAMES = ["Matt R", "Dawn D", "Tom V-B", "Amber H", "Eddy D'A", "Paul M", "Kyle W", "Matt C", "Keilan C"];
const HOLIDAY_DURATIONS = ["Morning", "Afternoon", "Full Day"];

const EMPTY_FORM = {
  id: "",
  date: "",
  orderReference: "",
  customerName: "",
  contact: "",
  number: "",
  address: "",
  installers: "",
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

export default function App() {
  const [board, setBoard] = useState(null);
  const [jobs, setJobs] = useState([]);
  const [holidays, setHolidays] = useState([]);
  const [form, setForm] = useState({ ...EMPTY_FORM, date: getLocalTodayIso() });
  const [editingId, setEditingId] = useState("");
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState(null);
  const [draggingJobId, setDraggingJobId] = useState("");
  const [dropDate, setDropDate] = useState("");
  const [activeHolidayDate, setActiveHolidayDate] = useState("");
  const [holidayForm, setHolidayForm] = useState({ person: STAFF_NAMES[0], duration: "Full Day" });
  const dragPreviewRef = useRef(null);
  const transparentDragImageRef = useRef(null);

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
      dragPreviewRef.current.style.left = `${event.clientX + 18}px`;
      dragPreviewRef.current.style.top = `${event.clientY + 18}px`;
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

  function resetForm(nextDate = board?.today || getLocalTodayIso()) {
    setEditingId("");
    setForm({ ...EMPTY_FORM, date: nextDate });
  }

  function editJob(job) {
    setEditingId(job.id);
    setForm({
      id: job.id,
      date: job.date,
      orderReference: job.orderReference || "",
      customerName: job.customerName || "",
      contact: job.contact || "",
      number: job.number || "",
      address: job.address || "",
      installers: job.installers || "",
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
          contact: form.contact.trim(),
          number: form.number.trim(),
          address: form.address.trim(),
          installers: form.installers.trim(),
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
    <div className="app-shell">
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
          <section className="panel form-panel">
            <div className="panel-head">
              <div>
                <p className="panel-kicker">Board editor</p>
                <h2>{editingId ? "Edit job" : "Add a job"}</h2>
              </div>
              <button className="ghost-button" type="button" onClick={() => resetForm()}>
                Clear
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
                <textarea
                  rows="3"
                  value={form.address}
                  onChange={(event) => setForm((current) => ({ ...current, address: event.target.value }))}
                />
              </label>

              <label>
                Installers
                <input
                  type="text"
                  value={form.installers}
                  onChange={(event) => setForm((current) => ({ ...current, installers: event.target.value }))}
                />
              </label>

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
                <button className="ghost-button" type="button" onClick={handleManualRefresh}>
                  Refresh
                </button>
              </div>
            </form>

            {message && <div className={`flash ${message.tone}`}>{message.text}</div>}

            <div className="upcoming">
              <div className="panel-head compact">
                <div>
                  <p className="panel-kicker">Next up</p>
                  <h3>Upcoming jobs</h3>
                </div>
              </div>

              {upcomingJobs.length === 0 ? (
                <p className="muted">No upcoming jobs saved yet.</p>
              ) : (
                <div className="upcoming-list">
                  {upcomingJobs.map((job) => (
                    <button key={job.id} className="upcoming-item" type="button" onClick={() => editJob(job)}>
                      <span>{job.date}</span>
                      <strong>{job.customerName}</strong>
                      <small>{job.orderReference || getJobTypeLabel(job)}</small>
                    </button>
                  ))}
                </div>
              )}
            </div>
          </section>

          <section className="panel board-panel">
            <div className="panel-head">
              <div>
                <p className="panel-kicker">Live board</p>
                <h2>Rolling installation calendar</h2>
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
                          event.preventDefault();
                          if (draggingJobId) setDropDate(row.isoDate);
                        }}
                        onDragLeave={() => {
                          if (dropDate === row.isoDate) setDropDate("");
                        }}
                        onDrop={(event) => {
                          event.preventDefault();
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
                        </button>

                        <div className="holiday-cell" onClick={() => setActiveHolidayDate((current) => (current === row.isoDate ? "" : row.isoDate))}>
                          <div className="holiday-stack">
                            {row.bankHoliday ? <span className="holiday-pill">{row.bankHoliday}</span> : null}
                            {row.staffHolidays.map((holiday) => (
                              <div key={holiday.id} className="staff-holiday-pill">
                                <span>{holiday.person}</span>
                                <b>{holiday.duration}</b>
                                <button
                                  type="button"
                                  className="holiday-delete"
                                  onClick={(event) => {
                                    event.stopPropagation();
                                    deleteHoliday(holiday.id);
                                  }}
                                >
                                  x
                                </button>
                              </div>
                            ))}
                            {!row.bankHoliday && row.staffHolidays.length === 0 ? <span className="muted">Click to add holiday</span> : null}
                            {activeHolidayDate === row.isoDate ? (
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
                          {row.jobs.length === 0 ? (
                            <span className="muted">No jobs booked</span>
                          ) : (
                            <div className="job-stack">
                              {row.jobs.map((job) => {
                                const meta = getJobTypeMeta(job.jobType);
                                return (
                                  <div
                                    key={job.id}
                                    className={`job-card ${draggingJobId === job.id ? "is-dragging" : ""}`}
                                    draggable
                                  onDragStart={(event) => {
                                    event.dataTransfer.setData("text/plain", job.id);
                                    event.dataTransfer.effectAllowed = "move";
                                    const preview = buildDragPreview(event.currentTarget);
                                    dragPreviewRef.current = preview;
                                    preview.style.left = `${event.clientX + 18}px`;
                                    preview.style.top = `${event.clientY + 18}px`;
                                    event.dataTransfer.setDragImage(getTransparentDragImage(), 0, 0);
                                    setDraggingJobId(job.id);
                                  }}
                                  onDragEnd={() => {
                                    setDraggingJobId("");
                                    setDropDate("");
                                    clearDragPreview();
                                  }}
                                >
                                    <div className="job-card-top">
                                      <div>
                                        <strong>{job.customerName}</strong>
                                        <p>{job.orderReference || "No order reference"}</p>
                                      </div>
                                      <span className={`job-tag ${meta.colorClass}`}>{getJobTypeLabel(job)}</span>
                                    </div>
                                    <div className="job-meta-grid">
                                      <p><b>Contact:</b> {job.contact || "-"}</p>
                                      <p><b>Number:</b> {job.number || "-"}</p>
                                      <p><b>Installers:</b> {job.installers || "-"}</p>
                                    </div>
                                    {job.address ? <p className="job-notes"><b>Address:</b> {job.address}</p> : null}
                                    {job.notes ? <p className="job-notes"><b>Notes:</b> {job.notes}</p> : null}
                                    <div className="job-actions">
                                      <button className="text-button" type="button" onClick={() => editJob(job)}>
                                        Edit
                                      </button>
                                      <button className="text-button danger" type="button" onClick={() => handleDelete(job.id)}>
                                        Delete
                                      </button>
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          )}
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
    </div>
  );
}
