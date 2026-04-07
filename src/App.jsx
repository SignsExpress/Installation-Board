import { useEffect, useMemo, useState } from "react";

const EMPTY_FORM = {
  id: "",
  date: "",
  title: "",
  crew: "",
  category: "Install",
  notes: ""
};

const CATEGORY_OPTIONS = ["Install", "Delivery", "Survey", "Access", "Holiday", "Other"];

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
  const [form, setForm] = useState({ ...EMPTY_FORM, date: getLocalTodayIso() });
  const [editingId, setEditingId] = useState("");
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState(null);

  useEffect(() => {
    let active = true;

    async function loadBoard() {
      try {
        setLoading(true);
        const [boardResponse, jobsResponse] = await Promise.all([fetch("/api/board"), fetch("/api/jobs")]);
        if (!boardResponse.ok || !jobsResponse.ok) {
          throw new Error("Could not load the installation board.");
        }

        const [boardData, jobsData] = await Promise.all([boardResponse.json(), jobsResponse.json()]);
        if (!active) return;
        setBoard(boardData);
        setJobs(Array.isArray(jobsData) ? jobsData : []);
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

  const upcomingJobs = useMemo(() => {
    const today = board?.today || getLocalTodayIso();
    return [...jobs]
      .filter((job) => job.date >= today)
      .sort((left, right) => {
        if (left.date !== right.date) return left.date.localeCompare(right.date);
        return left.title.localeCompare(right.title);
      })
      .slice(0, 8);
  }, [board?.today, jobs]);

  function resetForm(nextDate = board?.today || getLocalTodayIso()) {
    setEditingId("");
    setForm({ ...EMPTY_FORM, date: nextDate });
  }

  function editJob(job) {
    setEditingId(job.id);
    setForm({
      id: job.id,
      date: job.date,
      title: job.title || "",
      crew: job.crew || "",
      category: job.category || "Install",
      notes: job.notes || ""
    });
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function refreshJobs() {
    const response = await fetch("/api/jobs");
    if (!response.ok) throw new Error("Could not refresh jobs.");
    const nextJobs = await response.json();
    setJobs(Array.isArray(nextJobs) ? nextJobs : []);
  }

  async function handleSubmit(event) {
    event.preventDefault();

    if (!form.date || !form.title.trim()) {
      setMessage(createMessage("Pick a date and enter a job title.", "error"));
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
          title: form.title.trim(),
          crew: form.crew.trim(),
          category: form.category,
          notes: form.notes.trim()
        })
      });

      if (!response.ok) throw new Error("Could not save the job.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
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
      if (editingId === jobId) resetForm();
      setMessage(createMessage("Job deleted.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not delete the job.", "error"));
    }
  }

  async function handleManualRefresh() {
    try {
      const [boardResponse] = await Promise.all([fetch("/api/board"), refreshJobs()]);
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
          <div className="hero-copy">
            <p className="eyebrow">Signs Express installation planning</p>
            <h1>Installation Board</h1>
            <p className="lede">
              A live shared board that mirrors the paper pad: date down the left, bank holidays in the middle, jobs on the
              right, and a rolling view of last week, this week, next week, and the week after.
            </p>
          </div>
          <div className="hero-stats">
            <div className="stat-card">
              <span>Weeks shown</span>
              <strong>{board?.weeks?.length || 4}</strong>
            </div>
            <div className="stat-card">
              <span>Jobs saved</span>
              <strong>{jobs.length}</strong>
            </div>
            <div className="stat-card">
              <span>Live status</span>
              <strong>Realtime</strong>
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
                Job title
                <input
                  type="text"
                  placeholder="Moor Park remove"
                  value={form.title}
                  onChange={(event) => setForm((current) => ({ ...current, title: event.target.value }))}
                />
              </label>

              <div className="split-fields">
                <label>
                  Crew / initials
                  <input
                    type="text"
                    placeholder="MC + KC"
                    value={form.crew}
                    onChange={(event) => setForm((current) => ({ ...current, crew: event.target.value }))}
                  />
                </label>

                <label>
                  Type
                  <select
                    value={form.category}
                    onChange={(event) => setForm((current) => ({ ...current, category: event.target.value }))}
                  >
                    {CATEGORY_OPTIONS.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))}
                  </select>
                </label>
              </div>

              <label>
                Notes
                <textarea
                  rows="4"
                  placeholder="Access kit, timing, customer note, delivery details"
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
                      <strong>{job.title}</strong>
                      <small>{job.crew || job.category}</small>
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
                          row.holiday ? "is-holiday" : "",
                          row.isPast ? "is-past" : ""
                        ].join(" ").trim()}
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

                        <div className="holiday-cell">
                          {row.holiday ? <span className="holiday-pill">{row.holiday}</span> : <span className="muted">-</span>}
                        </div>

                        <div className="jobs-cell">
                          {row.jobs.length === 0 ? (
                            <span className="muted">No jobs booked</span>
                          ) : (
                            <div className="job-stack">
                              {row.jobs.map((job) => (
                                <div key={job.id} className="job-card">
                                  <div className="job-card-top">
                                    <div>
                                      <strong>{job.title}</strong>
                                      <p>{job.crew || "Crew not added yet"}</p>
                                    </div>
                                    <span className="job-tag">{job.category}</span>
                                  </div>
                                  {job.notes ? <p className="job-notes">{job.notes}</p> : null}
                                  <div className="job-actions">
                                    <button className="text-button" type="button" onClick={() => editJob(job)}>
                                      Edit
                                    </button>
                                    <button className="text-button danger" type="button" onClick={() => handleDelete(job.id)}>
                                      Delete
                                    </button>
                                  </div>
                                </div>
                              ))}
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
