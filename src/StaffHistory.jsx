import { useEffect, useMemo, useState } from "react";
import "./staff-history.css";

const api = "/api/staff-history";
const DAY = 86400000;
const statuses = { current: "Current", former: "Former", unverified: "Unverified" };
const reasons = {
  unknown: "Not recorded", resigned: "Resigned / gave notice", redundancy: "Redundancy",
  dismissed: "Dismissed", retired: "Retired", "contract-ended": "Contract ended", other: "Other"
};
const views = ["Directory", "Timeline", "Service", "Starts & finishes", "Headcount", "Leaving reasons"];
const blankSpell = () => ({ start: "", end: "", startLabel: "", endLabel: "", notice: "", reason: "unknown", notes: "", source: "" });
const blankPerson = () => ({ name: "", role: "", status: "current", notes: "", spells: [blankSpell()] });
const dateText = (date) => date ? new Intl.DateTimeFormat("en-GB", { dateStyle: "medium", timeZone: "UTC" }).format(new Date(date)) : "Unknown";
const today = () => new Intl.DateTimeFormat("sv-SE", { timeZone: "Europe/London" }).format(new Date());
const dateValue = (date) => Date.parse(date);
const currentSpell = (person, index) => person.status === "current" && index === person.spells.length - 1;

function serviceFor(person, index, asOf) {
  const spell = person.spells[index];
  const end = spell.end || (currentSpell(person, index) ? asOf : "");
  if (!spell.start || !end || end < spell.start) return null;
  return { days: (dateValue(end) - dateValue(spell.start)) / DAY, approximate: Boolean(spell.startLabel || spell.endLabel) };
}

function duration(days) {
  if (days === 0) return "Same day";
  if (days < 62) return `${days}d`;
  const months = Math.round(days / 30.4375);
  const years = Math.floor(months / 12);
  const remainder = months % 12;
  return [years ? `${years}y` : "", remainder ? `${remainder}m` : ""].filter(Boolean).join(" ");
}

function describeService(service) {
  return service ? `${service.approximate ? "≈ " : ""}${duration(service.days)}` : "Unknown";
}

function combinedService(person, asOf) {
  const spells = person.spells.map((_, index) => serviceFor(person, index, asOf));
  const known = spells.filter(Boolean);
  if (!known.length) return null;
  return {
    days: known.reduce((sum, spell) => sum + spell.days, 0),
    approximate: known.some((spell) => spell.approximate),
    partial: known.length !== spells.length
  };
}

function finishText(person, spell, index) {
  if (currentSpell(person, index)) return "Current";
  return spell.endLabel || dateText(spell.end);
}

async function request(path = "", method = "GET", body) {
  const response = await fetch(api + path, {
    method,
    cache: "no-store",
    headers: { "X-Staff-History": "1", ...(body ? { "Content-Type": "application/json" } : {}) },
    body: body ? JSON.stringify(body) : undefined
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload.error || "Could not access Staff History.");
  return payload;
}

async function preparePhoto(file) {
  if (file.size > 20 * 1024 * 1024) throw new Error("Choose a photo under 20 MB.");
  const bitmap = await createImageBitmap(file);
  try {
    const canvas = document.createElement("canvas");
    const size = Math.min(480, bitmap.width, bitmap.height);
    canvas.width = size;
    canvas.height = size;
    const crop = Math.min(bitmap.width, bitmap.height);
    const context = canvas.getContext("2d");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, size, size);
    context.drawImage(bitmap, (bitmap.width - crop) / 2, (bitmap.height - crop) / 2, crop, crop, 0, 0, size, size);
    return canvas.toDataURL("image/jpeg", 0.78).split(",")[1];
  } finally {
    bitmap.close();
  }
}

function Avatar({ person }) {
  const initials = person.name.trim().split(/\s+/).slice(0, 2).map((name) => name[0]).join("").toUpperCase();
  return <span className="sh-avatar" aria-hidden="true">
    {person.photoVersion
      ? <img loading="lazy" alt="" src={`${api}/people/${person.id}/photo?v=${person.photoVersion}`} />
      : initials}
  </span>;
}

function PersonButton({ person, onOpen }) {
  return <button className="sh-person" type="button" onClick={() => onOpen(person)}>
    <Avatar person={person} /><span>{person.name}</span>
  </button>;
}

function Directory({ people, asOf, onOpen }) {
  return <div className="sh-table-wrap"><table className="sh-table">
    <thead><tr><th>Employee</th><th>Status</th><th>Started</th><th>Finished</th><th>Service</th><th>Leaving reason</th></tr></thead>
    <tbody>{people.flatMap((person) => person.spells.map((spell, index) => <tr key={`${person.id}-${index}`}>
      <td><PersonButton person={person} onOpen={onOpen} />{person.spells.length > 1 && <small>Spell {index + 1} of {person.spells.length}</small>}</td>
      <td><span className={`sh-tag ${currentSpell(person, index) ? "current" : ""}`}>{currentSpell(person, index) ? "Current" : spell.end || person.status === "former" ? "Former" : "Unverified"}</span></td>
      <td>{spell.startLabel || dateText(spell.start)}</td>
      <td>{finishText(person, spell, index)}</td>
      <td>{describeService(serviceFor(person, index, asOf))}</td>
      <td>{currentSpell(person, index) ? "—" : reasons[spell.reason]}</td>
    </tr>))}</tbody>
  </table></div>;
}

function Timeline({ people, asOf, onOpen }) {
  const years = people.flatMap((person) => person.spells.flatMap((spell) => [spell.start, spell.end])).filter(Boolean).map((date) => Number(date.slice(0, 4)));
  const last = Math.max(Number(asOf.slice(0, 4)), ...years) + 1;
  const first = Math.min(last - 1, ...years);
  const min = Date.UTC(first, 0, 1);
  const max = Date.UTC(last, 0, 1);
  const percent = (date) => 100 * (dateValue(date) - min) / (max - min);
  const step = Math.max(1, Math.ceil((last - first) / 7));
  const ticks = Array.from({ length: Math.ceil((last - first) / step) }, (_, i) => first + i * step);
  return <div className="sh-timeline">
    <div className="sh-timeline-row sh-timeline-head"><strong>Employee</strong><div className="sh-track">{ticks.map((year) =>
      <span key={year} style={{ left: percent(`${year}-01-01`) + "%" }}>{year}</span>)}</div></div>
    {people.map((person) => <div className="sh-timeline-row" key={person.id}>
      <PersonButton person={person} onOpen={onOpen} />
      <div className="sh-lanes">{person.spells.map((spell, index) => {
        const ongoing = currentSpell(person, index);
        const end = spell.end || (ongoing ? asOf : "");
        const label = `${spell.startLabel || dateText(spell.start)} → ${finishText(person, spell, index)}`;
        if (!spell.start && !end) return <small key={index}>{label}</small>;
        const start = spell.start || end;
        const width = end ? Math.max(0, percent(end) - percent(start)) : 0;
        return <div className="sh-track" key={index}>
          <button type="button" title={label} aria-label={`${person.name}: ${label}`} onClick={() => onOpen(person)}
            className={`sh-span ${ongoing ? "current" : ""} ${spell.startLabel || spell.endLabel ? "approximate" : ""}`}
            style={{ left: percent(start) + "%", width: width + "%" }} />
        </div>;
      })}</div>
    </div>)}
    <p className="sh-caption">Green shows service to date. A dot indicates one known date. Stripes flag approximate or disputed dates.</p>
  </div>;
}

function ServiceChart({ people, asOf, onOpen }) {
  const totals = people.map((person) => ({ person, total: combinedService(person, asOf) }))
    .filter(({ total }) => total).sort((a, b) => b.total.days - a.total.days);
  const max = Math.max(1, ...totals.map(({ total }) => total.days));
  return <div className="sh-ranking">
    <p className="sh-caption">Combined time across each person’s spells, excluding gaps. Current service runs to {dateText(asOf)}.</p>
    {!totals.length && <p>No calculable service for this selection.</p>}
    {totals.map(({ person, total }) => <div className="sh-rank" key={person.id}>
      <PersonButton person={person} onOpen={onOpen} />
      <div className="sh-meter"><span className={person.status === "current" ? "current" : ""} style={{ width: `${total.days / max * 100}%` }} /></div>
      <strong>{describeService(total)}{total.partial ? " known" : ""}</strong>
    </div>)}
  </div>;
}

function YearChart({ people, headcount, asOf, onOpen }) {
  const [selected, setSelected] = useState(null);
  const events = people.flatMap((person) => person.spells.map((spell, index) => ({ person, spell, index })));
  const dated = events.flatMap(({ spell }) => [spell.start, spell.end]).filter(Boolean);
  if (!dated.length) return <p>No dated records in this selection.</p>;
  const years = dated.map((date) => Number(date.slice(0, 4)));
  const first = Math.min(...years);
  const last = Math.max(Number(asOf.slice(0, 4)), ...years);
  const points = Array.from({ length: last - first + 1 }, (_, i) => {
    const year = first + i;
    const snapshot = year === Number(asOf.slice(0, 4)) ? asOf : `${year}-12-31`;
    const starts = events.filter(({ spell }) => spell.start.startsWith(String(year)) && !spell.startLabel);
    const finishes = events.filter(({ spell }) => spell.end.startsWith(String(year)) && !spell.endLabel);
    const covered = events.filter(({ person, spell, index }) => spell.start && !spell.startLabel && spell.start <= snapshot &&
      ((spell.end && !spell.endLabel && spell.end >= snapshot) || currentSpell(person, index)));
    const active = [...new Map(covered.map((entry) => [entry.person.id, entry])).values()];
    return { year, starts, finishes, active };
  });
  const max = Math.max(1, ...points.flatMap((point) => headcount ? [point.active.length] : [point.starts.length, point.finishes.length]));
  const point = points.find((entry) => entry.year === selected) || points.at(-1);
  const groups = headcount ? [["People in recorded service", point.active]] : [["Starts", point.starts], ["Finishes", point.finishes]];
  return <>
    <p className="sh-caption">{headcount
      ? "Recorded service at each year end, or today for this year. Includes spells with known boundaries and confirmed current staff. Earlier records may be incomplete."
      : "Counts employment spells by exact date. Approximate and disputed dates stay visible in the directory."}</p>
    <div className="sh-year-scroll"><div className="sh-year-chart">
      {points.map((entry) => <button key={entry.year} type="button" aria-pressed={point.year === entry.year} onClick={() => setSelected(entry.year)}>
        <span className="sh-year-bars" aria-hidden="true">{(headcount ? [entry.active] : [entry.starts, entry.finishes]).map((items, index) =>
          <span key={index} className={index === 1 ? "finish" : ""} style={{ height: `${items.length / max * 100}%` }} />)}</span>
        <small>{entry.year}</small>
        <span className="sh-year-count">{headcount ? entry.active.length : `${entry.starts.length} / ${entry.finishes.length}`}</span>
      </button>)}
    </div></div>
    <div className="sh-breakdown"><h3>{point.year}</h3>{groups.map(([label, items]) => <div key={label}>
      <h4>{label} · {items.length}</h4>
      <div className="sh-chips">{items.map(({ person, index }) => <PersonButton key={`${person.id}-${index}`} person={person} onOpen={onOpen} />)}
        {!items.length && <small>No records</small>}
      </div>
    </div>)}</div>
  </>;
}

function Reasons({ people, onOpen }) {
  const events = people.flatMap((person) => person.spells.flatMap((spell, index) =>
    !currentSpell(person, index) && (spell.end || spell.endLabel || person.status === "former")
      ? [{ person, spell, index }] : []));
  const [selected, setSelected] = useState(null);
  const counts = Object.entries(reasons).map(([key, label]) => ({
    key, label, items: events.filter(({ spell }) => spell.reason === key)
  })).sort((a, b) => b.items.length - a.items.length);
  const group = counts.find(({ key }) => key === selected) || counts[0];
  const max = Math.max(1, ...counts.map(({ items }) => items.length));
  return <div className="sh-ranking">
    <p className="sh-caption">Reasons are recorded per employment spell. Missing evidence is counted as not recorded.</p>
    {counts.map(({ key, label, items }) => <button type="button" key={key} className="sh-rank sh-reason" aria-pressed={group.key === key} onClick={() => setSelected(key)}>
      <span>{label}</span><span className="sh-meter"><span style={{ width: `${items.length / max * 100}%` }} /></span><strong>{items.length}</strong>
    </button>)}
    <div className="sh-breakdown"><h3>{group.label}</h3><div className="sh-chips">
      {group.items.map(({ person, spell, index }) => <div key={`${person.id}-${index}`}><PersonButton person={person} onOpen={onOpen} />
        <small>{spell.notice ? `Notice: ${dateText(spell.notice)} · ` : ""}{finishText(person, spell, index)}</small></div>)}
      {!group.items.length && <p>No recorded departures in this category.</p>}
    </div></div>
  </div>;
}

function Editor({ person, asOf, busy, onClose, onSave, onPhoto }) {
  const [draft, setDraft] = useState(person);
  useEffect(() => { document.getElementById("staff-history-editor")?.scrollIntoView({ block: "start" }); }, []);
  const edit = (key, value) => setDraft((current) => ({ ...current, [key]: value }));
  const editSpell = (index, key, value) => setDraft((current) => ({
    ...current, spells: current.spells.map((spell, i) => i === index ? { ...spell, [key]: value } : spell)
  }));
  return <section className="panel sh-editor" id="staff-history-editor" aria-label="Employee record">
    <header><div className="sh-profile"><Avatar person={person} /><div><h2>{person.name || "Add employee"}</h2>
      <small>Combined service: {describeService(combinedService(draft, asOf))}</small></div></div>
      <button type="button" onClick={onClose}>Close</button></header>
    {person.id && <div className="sh-photo-actions">
      <label className="sh-file">Upload photo<input type="file" accept="image/*" disabled={busy}
        onChange={(event) => { const file = event.target.files?.[0]; event.target.value = ""; if (file) onPhoto(person, file); }} /></label>
      {person.photoVersion && <button type="button" disabled={busy} onClick={() => onPhoto(person, null)}>Remove photo</button>}
    </div>}
    <form onSubmit={(event) => { event.preventDefault(); onSave(draft); }}>
      <div className="sh-form-grid">
        <label>Name<input value={draft.name} required maxLength={120} onChange={(event) => edit("name", event.target.value)} /></label>
        <label>Role<input value={draft.role} maxLength={120} onChange={(event) => edit("role", event.target.value)} /></label>
        <label>Status<select value={draft.status} onChange={(event) => edit("status", event.target.value)}>
          {Object.entries(statuses).map(([value, label]) => <option key={value} value={value}>{label}</option>)}
        </select></label>
      </div>
      <label>Employee notes<textarea rows={2} maxLength={3000} value={draft.notes} onChange={(event) => edit("notes", event.target.value)} /></label>
      {draft.spells.map((spell, index) => <fieldset key={index}>
        <legend>Employment spell {index + 1} · {describeService(serviceFor(draft, index, asOf))}</legend>
        <div className="sh-form-grid">
          {[["start", "Start date"], ["end", "Finish date"], ["notice", "Notice given / received"]].map(([key, label]) =>
            <label key={key}>{label}<input type="date" value={spell[key]} onChange={(event) => editSpell(index, key, event.target.value)} /></label>)}
          <label>Leaving reason<select value={spell.reason} onChange={(event) => editSpell(index, "reason", event.target.value)}>
            {Object.entries(reasons).map(([value, label]) => <option key={value} value={value}>{label}</option>)}
          </select></label>
          {[["startLabel", "Start date qualifier"], ["endLabel", "Finish date qualifier"]].map(([key, label]) =>
            <label key={key}>{label}<input value={spell[key]} placeholder="Optional approximate or disputed date" maxLength={100} onChange={(event) => editSpell(index, key, event.target.value)} /></label>)}
        </div>
        <label>Evidence / source<input value={spell.source} maxLength={500} onChange={(event) => editSpell(index, "source", event.target.value)} /></label>
        <label>Spell notes<textarea rows={2} value={spell.notes} maxLength={2000} onChange={(event) => editSpell(index, "notes", event.target.value)} /></label>
        {draft.spells.length > 1 && <button type="button" onClick={() => edit("spells", draft.spells.filter((_, i) => i !== index))}>Remove this spell</button>}
      </fieldset>)}
      <div className="sh-actions">
        <button type="button" disabled={draft.spells.length >= 30} onClick={() => edit("spells", [...draft.spells, blankSpell()])}>Add spell</button>
        <button type="submit" className="primary-button" disabled={busy}>{busy ? "Saving…" : "Save employee"}</button>
      </div>
    </form>
  </section>;
}

export default function StaffHistory({ navigation }) {
  const [people, setPeople] = useState(null);
  const [view, setView] = useState(views[0]);
  const [filter, setFilter] = useState("all");
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState(null);
  const [editorKey, setEditorKey] = useState(0);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState(null);
  const asOf = today();

  function load() {
    setMessage(null);
    request().then(({ people: records }) => setPeople(records)).catch((error) => setMessage({ error: error.message }));
  }
  useEffect(load, []);

  const filtered = useMemo(() => (people || []).filter((person) =>
    (filter === "all" || person.status === filter) &&
    `${person.name} ${person.role}`.toLowerCase().includes(search.trim().toLowerCase())
  ).sort((a, b) => Object.keys(statuses).indexOf(a.status) - Object.keys(statuses).indexOf(b.status) ||
    a.name.localeCompare(b.name, "en-GB")), [people, filter, search]);

  const employee = selected === "new" ? blankPerson() : people?.find((person) => person.id === selected);
  function open(person) {
    setSelected(person.id);
    setEditorKey((key) => key + 1);
    setMessage(null);
  }
  function updateRecord(saved) {
    setPeople((records) => records.some((person) => person.id === saved.id)
      ? records.map((person) => person.id === saved.id ? saved : person) : [...records, saved]);
  }
  async function perform(action) {
    setBusy(true);
    setMessage(null);
    try { await action(); }
    catch (error) { setMessage({ error: error.message }); }
    finally { setBusy(false); }
  }
  function save(draft) {
    return perform(async () => {
      const saved = await request(draft.id ? `/people/${draft.id}` : "/people", draft.id ? "PUT" : "POST", draft);
      updateRecord(saved);
      open(saved);
      setMessage({ success: "Employee saved." });
    });
  }
  function photo(person, file) {
    return perform(async () => {
      const path = `/people/${person.id}/photo`;
      const saved = file ? await request(path, "PUT", { jpeg: await preparePhoto(file) }) : await request(path, "DELETE");
      updateRecord(saved);
      setMessage({ success: file ? "Photo saved." : "Photo removed." });
    });
  }
  function importFile(file) {
    return perform(async () => {
      if (file.size > 2 * 1024 * 1024) throw new Error("Choose a Staff History JSON file under 2 MB.");
      const data = JSON.parse(await file.text());
      const result = await request("/import", "POST", data);
      setPeople(result.people);
      setMessage({ success: `${result.added} added · ${result.skipped} existing records kept.` });
    });
  }

  return <div className="app-shell">{navigation}<main className="page staff-history">
    <section className="panel sh-main">
      <div className="sh-top"><div><div className="sh-kicker">PRIVATE MODULE</div><h1>Staff History</h1></div>
        {people && <div className="sh-actions">
          <label className="sh-file">Import history<input type="file" accept=".json,application/json" disabled={busy}
            onChange={(event) => { const file = event.target.files?.[0]; event.target.value = ""; if (file) importFile(file); }} /></label>
          <button type="button" className="primary-button" disabled={busy} onClick={() => { setSelected("new"); setEditorKey((key) => key + 1); }}>Add employee</button>
        </div>}
      </div>
      {message && <p className={`flash ${message.error ? "error" : "success"}`} role={message.error ? "alert" : "status"}>{message.error || message.success}</p>}
      {!people && (message?.error ? <button type="button" onClick={load}>Retry</button> : <p>Loading staff history…</p>)}
      {employee && <Editor key={editorKey} person={employee} asOf={asOf} busy={busy} onClose={() => setSelected(null)} onSave={save} onPhoto={photo} />}
      {people && !people.length && <p className="sh-empty">Import the staff history file or add an employee to start the register.</p>}
      {people?.length > 0 && <>
        <div className="sh-stats">
          <span><strong>{people.length}</strong> people</span>
          <span><strong>{people.filter((person) => person.status === "current").length}</strong> current</span>
          <span><strong>{people.reduce((sum, person) => sum + person.spells.length, 0)}</strong> spells</span>
          <span><strong>{people.filter((person) => person.status === "unverified").length}</strong> unverified</span>
        </div>
        <nav className="sh-views" aria-label="Staff history views">{views.map((label) =>
          <button key={label} type="button" aria-pressed={view === label} onClick={() => setView(label)}>{label}</button>)}</nav>
        <div className="sh-controls">
          <label>Search staff<input type="search" value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Name or role" /></label>
          <label>Status<select value={filter} onChange={(event) => setFilter(event.target.value)}><option value="all">All staff</option>
            {Object.entries(statuses).map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select></label>
        </div>
        {!filtered.length ? <p>No employees match these filters.</p> : <>
          {view === "Directory" && <Directory people={filtered} asOf={asOf} onOpen={open} />}
          {view === "Timeline" && <Timeline people={filtered} asOf={asOf} onOpen={open} />}
          {view === "Service" && <ServiceChart people={filtered} asOf={asOf} onOpen={open} />}
          {view === "Starts & finishes" && <YearChart key={view} people={filtered} asOf={asOf} onOpen={open} />}
          {view === "Headcount" && <YearChart key={view} people={filtered} headcount asOf={asOf} onOpen={open} />}
          {view === "Leaving reasons" && <Reasons people={filtered} onOpen={open} />}
        </>}
      </>}
    </section>
  </main></div>;
}
