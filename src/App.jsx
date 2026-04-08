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

const STAFF_NAMES = ["Matt R", "Dawn D", "Tom V-B", "Amber H", "Eddy D'A", "Paul M", "Kyle W", "Matt C", "Keilan C"];
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

function MainNavBar({ currentUser, active = "home", onLogout }) {
  function goTo(path) {
    window.location.assign(path);
  }

  const isClientUser = currentUser?.role === "client";
  const homePath = isClientUser ? "/client" : "/";
  const boardPath = isClientUser ? "/client/board" : "/board";
  const installerPath = "/installer";

  return (
    <nav className="panel host-nav">
      <div className="host-nav-links">
        <button
          type="button"
          className={`host-nav-link ${active === "home" ? "active" : ""}`}
          onClick={() => goTo(homePath)}
        >
          Home
        </button>
        <button
          type="button"
          className={`host-nav-link ${active === "board" ? "active" : ""}`}
          onClick={() => goTo(boardPath)}
        >
          Installation Board
        </button>
        <button
          type="button"
          className={`host-nav-link ${active === "installer" ? "active" : ""} ${isClientUser ? "disabled" : ""}`}
          onClick={() => {
            if (isClientUser) return;
            goTo(installerPath);
          }}
          disabled={isClientUser}
        >
          Subcontractor Directory
        </button>
      </div>
      <div className="host-nav-meta">
        <span className="host-nav-user">Logged in as <strong>{currentUser.displayName}</strong></span>
        <button className="host-nav-logout" type="button" onClick={onLogout}>
          Log out
        </button>
      </div>
    </nav>
  );
}

function HostLandingPage({ currentUser, onLogout }) {
  function goTo(path) {
    window.location.assign(path);
  }

  return (
    <div className="app-shell host-landing-shell">
      <div className="page host-landing-page">
        <div className="host-landing-brand">
          <img className="hero-logo host-landing-brand-logo" src="/branding/signs-express-logo.svg" alt="Signs Express" />
        </div>
        <MainNavBar currentUser={currentUser} active="home" onLogout={onLogout} />

        <section className="panel host-landing-panel">
          <div className="host-landing-actions">
            <button className="host-launch-card" type="button" onClick={() => goTo("/installer")}>
              <strong>Subcontractor Database</strong>
            </button>

            <button className="host-launch-card" type="button" onClick={() => goTo("/board")}>
              <strong>Installation Board</strong>
            </button>
          </div>
        </section>
      </div>
    </div>
  );
}

function ClientLandingPage({ currentUser, onLogout }) {
  function goTo(path) {
    window.location.assign(path);
  }

  return (
    <div className="app-shell host-landing-shell client-landing-shell">
      <div className="page host-landing-page">
        <div className="host-landing-brand">
          <img className="hero-logo host-landing-brand-logo" src="/branding/signs-express-logo.svg" alt="Signs Express" />
        </div>
        <MainNavBar currentUser={currentUser} active="home" onLogout={onLogout} />

        <section className="panel host-landing-panel">
          <div className="host-landing-actions">
            <button className="host-launch-card disabled" type="button" disabled>
              <strong>Subcontractor Directory</strong>
            </button>

            <button className="host-launch-card" type="button" onClick={() => goTo("/client/board")}>
              <strong>Installation Board</strong>
            </button>
          </div>
        </section>
      </div>
    </div>
  );
}

export default function App() {
  const pathname = typeof window !== "undefined" ? window.location.pathname.toLowerCase() : "/";
  const isClientRoute = pathname.startsWith("/client");
  const isClientBoardRoute = pathname.startsWith("/client/board");
  const isInstallerRoute = pathname.startsWith("/installer");
  const isBoardRoute = pathname.startsWith("/board");
  const [board, setBoard] = useState(null);
  const [jobs, setJobs] = useState([]);
  const [holidays, setHolidays] = useState([]);
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
  const [orderLookupOpen, setOrderLookupOpen] = useState(false);
  const [orderLookupQuery, setOrderLookupQuery] = useState("");
  const [orderLookupLoading, setOrderLookupLoading] = useState(false);
  const [orderLookupResults, setOrderLookupResults] = useState([]);
  const [orderLookupError, setOrderLookupError] = useState("");
  const [orderLookupDebugMode, setOrderLookupDebugMode] = useState(false);
  const [activeCoreBridgeDebugOrder, setActiveCoreBridgeDebugOrder] = useState(null);
  const [coreBridgeDebugView, setCoreBridgeDebugView] = useState("fields");
  const [currentUser, setCurrentUser] = useState(null);
  const [authChecked, setAuthChecked] = useState(false);
  const [loginUsers, setLoginUsers] = useState([]);
  const [loginDisplayName, setLoginDisplayName] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginError, setLoginError] = useState("");
  const dragPreviewRef = useRef(null);
  const transparentDragImageRef = useRef(null);
  const dragPositionRef = useRef({ x: 0, y: 0 });
  const isHostUser = currentUser?.role === "host";
  const isClientUser = currentUser?.role === "client";
  const isClientMode = isClientUser;
  const showInstallerDirectory = Boolean(currentUser && isHostUser && isInstallerRoute);
  const showBoard = Boolean(currentUser && (isBoardRoute || (isClientUser && isClientBoardRoute)));
  const showHostLanding = Boolean(currentUser && isHostUser && !isInstallerRoute && !isBoardRoute);
  const showClientLanding = Boolean(currentUser && isClientUser && !isClientBoardRoute);

  useEffect(() => {
    let active = true;

    async function loadAuth() {
      try {
        const [usersResponse, meResponse] = await Promise.all([
          fetch("/api/auth/users"),
          fetch("/api/auth/me")
        ]);

        const usersPayload = usersResponse.ok ? await usersResponse.json() : [];
        const mePayload = meResponse.ok ? await meResponse.json() : null;

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
    if (!currentUser || !showBoard) return undefined;
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
  }, [currentUser, showBoard]);

  useEffect(() => {
    if (!currentUser) return;
    if (currentUser.role === "client" && !isClientRoute) {
      window.location.replace("/client");
      return;
    }

    if (currentUser.role === "host" && isClientRoute) {
      window.location.replace("/");
    }
  }, [currentUser, isClientRoute]);

  useEffect(() => {
    if (!currentUser || !showBoard) return undefined;
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
  }, [currentUser, showBoard]);

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
    setOrderLookupOpen(false);
    setOrderLookupQuery("");
    setOrderLookupResults([]);
    setOrderLookupError("");
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
      if (payload.user?.role === "client" && !isClientRoute) {
        window.location.replace("/client");
        return;
      }
      if (payload.user?.role === "host" && isClientRoute) {
        window.location.replace("/");
      }
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
      setLoginPassword("");
      setLoginError("");
      if (isClientRoute) {
        window.location.replace("/client");
      } else {
        window.location.replace("/");
      }
    }
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

  async function searchCoreBridgeOrders(searchTerm = orderLookupQuery, debugMode = orderLookupDebugMode) {
    if (isClientMode) return;

    try {
      setOrderLookupLoading(true);
      setOrderLookupError("");
      const query = String(searchTerm || "").trim();
      const params = new URLSearchParams();
      if (query) params.set("q", query);
      if (debugMode) params.set("debug", "1");
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

  async function openOrderLookup() {
    setOrderLookupOpen(true);
    setOrderLookupError("");
    setActiveCoreBridgeDebugOrder(null);
    setCoreBridgeDebugView("fields");
  }

  async function applyCoreBridgeOrder(order) {
    let resolvedOrder = order;

    try {
      if (order?.id) {
        setOrderLookupLoading(true);
        setOrderLookupError("");
        const params = new URLSearchParams();
        if (orderLookupDebugMode) params.set("debug", "1");
        const url = params.toString()
          ? `/api/corebridge/orders/${encodeURIComponent(order.id)}?${params.toString()}`
          : `/api/corebridge/orders/${encodeURIComponent(order.id)}`;
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
    setActiveCoreBridgeDebugOrder(null);
    setMessage(createMessage("Order details copied into the job form.", "success"));
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

  async function duplicateJobToDate(jobId, nextDate) {
    if (isClientMode) return;
    const job = jobsById.get(jobId);
    if (!job || !nextDate) {
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
          date: nextDate
        })
      });

      if (!response.ok) throw new Error("Could not duplicate the job.");

      const payload = await response.json();
      setBoard(payload.board);
      setJobs(payload.jobs);
      setHolidays(payload.holidays);
      setMessage(createMessage(`Job copied to ${nextDate}.`, "success"));
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
      setMessage(createMessage("Holiday updated.", "success"));
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

  async function cycleHoliday(date, person) {
    if (isClientMode) return;
    const existing = holidays.find((item) => item.date === date && item.person === person);
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

    await deleteHoliday(existing.id);
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
    return <HostLandingPage currentUser={currentUser} onLogout={handleLogout} />;
  }

  if (showClientLanding) {
    return <ClientLandingPage currentUser={currentUser} onLogout={handleLogout} />;
  }

  if (showInstallerDirectory) {
    return <InstallerDirectoryHost currentUser={currentUser} onLogout={handleLogout} />;
  }

  return (
    <div className={`app-shell ${isClientMode ? "client-mode" : "editor-mode"}`}>
      <div className="page">
        <MainNavBar currentUser={currentUser} active="board" onLogout={handleLogout} />

        <div className="layout">
          <section className="panel board-panel board-panel-full">
            {loading || !board ? (
              <div className="board-loading">Loading the shared installation board...</div>
            ) : (
              <div className="board">
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
                                    className={`date-holiday-chip ${getHolidayPersonColor(holiday.person)} active`}
                                    onClick={(event) => {
                                      event.stopPropagation();
                                      handleHolidayChipClick(row.isoDate, holiday.person, false);
                                    }}
                                  >
                                    {toInitials(holiday.person)}{durationLabel}
                                  </button>
                                );
                              })}
                            </div>
                          ) : null}
                          {!isClientMode && activeHolidayDate === row.isoDate ? (
                            <div className="date-holiday-popover" onClick={(event) => event.stopPropagation()}>
                              {STAFF_NAMES.map((name) => {
                                const existing = row.staffHolidays.find((holiday) => holiday.person === name);
                                const durationLabel =
                                  existing?.duration === "Morning"
                                    ? ".AM"
                                    : existing?.duration === "Afternoon"
                                      ? ".PM"
                                      : "";
                                const initials = toInitials(name);
                                return (
                                  <button
                                    key={`${row.isoDate}-${name}`}
                                    type="button"
                                    className={`date-holiday-chip ${getHolidayPersonColor(name)} ${existing ? "active" : ""}`}
                                    onClick={(event) => {
                                      event.stopPropagation();
                                      handleHolidayChipClick(row.isoDate, name, true);
                                    }}
                                  >
                                    {initials}{durationLabel}
                                  </button>
                                );
                              })}
                            </div>
                          ) : null}
                          {!isClientMode && row.bankHoliday ? <span className="date-holiday-chip date-bank-holiday">{row.bankHoliday}</span> : null}
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

              <div className="corebridge-lookup-bar">
                <button className="ghost-button" type="button" onClick={() => openOrderLookup()}>
                  Find order
                </button>
                <span className="muted">Pull live order details from CoreBridge where available.</span>
              </div>

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
      {!isClientMode && orderLookupOpen ? (
        <div className="modal-backdrop" onClick={() => setOrderLookupOpen(false)}>
          <div className="modal order-lookup-modal" onClick={(event) => event.stopPropagation()}>
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
              <button
                className={`ghost-button ${orderLookupDebugMode ? "active-debug-toggle" : ""}`}
                type="button"
                onClick={() => {
                  setOrderLookupDebugMode((current) => !current);
                  setActiveCoreBridgeDebugOrder(null);
                }}
              >
                {orderLookupDebugMode ? "Debug on" : "Debug off"}
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
                    {orderLookupDebugMode ? (
                      <div className="order-result-debug">
                        <span><b>Detail fetch:</b> {order._detailFetched ? `ok (${order._detailOrderId})` : `fallback (${order._detailError || "no detail"})`}</span>
                        <span><b>Normalized description:</b> {order.description || "-"}</span>
                        <span><b>Normalized number:</b> {order.number || "-"}</span>
                        <span><b>Normalized address:</b> {order.address || "-"}</span>
                      </div>
                    ) : null}
                    <div className="order-result-actions">
                      <button className="primary-button" type="button" onClick={() => applyCoreBridgeOrder(order)}>
                        Use this order
                      </button>
                      {orderLookupDebugMode ? (
                        <button
                  className="ghost-button"
                          type="button"
                          onClick={() => {
                            setActiveCoreBridgeDebugOrder(order);
                            setCoreBridgeDebugView("fields");
                          }}
                        >
                          Inspect fields
                        </button>
                      ) : null}
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
      {!isClientMode && activeCoreBridgeDebugOrder ? (
        <div className="modal-backdrop" onClick={() => setActiveCoreBridgeDebugOrder(null)}>
          <div className="modal corebridge-debug-modal" onClick={(event) => event.stopPropagation()}>
            <div className="modal-head">
              <div>
                <h3>CoreBridge Field Debug</h3>
                <p>
                  {activeCoreBridgeDebugOrder.orderReference || "No order ref"} -{" "}
                  {activeCoreBridgeDebugOrder.customerName || "Unnamed customer"}
                </p>
              </div>
              <button className="icon-button" type="button" onClick={() => setActiveCoreBridgeDebugOrder(null)}>
                x
              </button>
            </div>
            <div className="corebridge-debug-switches">
              <button
                type="button"
                className={`ghost-button ${coreBridgeDebugView === "fields" ? "active-debug-toggle" : ""}`}
                onClick={() => setCoreBridgeDebugView("fields")}
              >
                Flattened fields
              </button>
              <button
                type="button"
                className={`ghost-button ${coreBridgeDebugView === "raw" ? "active-debug-toggle" : ""}`}
                onClick={() => setCoreBridgeDebugView("raw")}
              >
                Raw JSON
              </button>
            </div>
            <div className="corebridge-debug-help">
              <strong>Tell me which variable has the right value.</strong>
              <span>
                {coreBridgeDebugView === "fields"
                  ? "Left column is the field name from CoreBridge. Right column is the value."
                  : "Open Raw JSON and search for the exact text, for example Wigton - Valve Markers."}
              </span>
            </div>
            {coreBridgeDebugView === "fields" ? (
              <div className="corebridge-debug-table">
                {(activeCoreBridgeDebugOrder.debugFields || []).map((field) => (
                  <div key={field.key} className="corebridge-debug-row">
                    <div className="corebridge-debug-key">{field.key}</div>
                    <div className="corebridge-debug-value">{field.value}</div>
                  </div>
                ))}
              </div>
            ) : (
              <pre className="corebridge-debug-raw">{activeCoreBridgeDebugOrder.debugRaw || "{}"}</pre>
            )}
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
