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
  isPlaceholder: false,
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

function renderJobCardContent({
  job,
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
      className={`job-card ${meta.colorClass}-card ${job.isPlaceholder ? "is-placeholder" : ""} ${draggingJobId === job.id ? "is-dragging" : ""}`}
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
    </div>
  );
}

function getPermissionForApp(user, key) {
  const fallback =
    key === "board"
      ? user?.role === "host"
        ? "admin"
        : "user"
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

function usesHostShell(user) {
  return Boolean(user && (canAccessInstaller(user) || canEditBoard(user) || user.canManagePermissions));
}

function getHomePathForUser(user) {
  return usesHostShell(user) ? "/" : "/client";
}

function getBoardPathForUser(user) {
  return canEditBoard(user) ? "/board" : "/client/board";
}

function buildBoardUrl(startIso = "", endIso = "") {
  if (startIso && endIso) {
    return `/api/board?start=${encodeURIComponent(startIso)}&end=${encodeURIComponent(endIso)}`;
  }
  return "/api/board";
}

function MainNavBar({ currentUser, active = "home", onLogout }) {
  function goTo(path) {
    window.location.assign(path);
  }

  const boardAllowed = canAccessBoard(currentUser);
  const installerAllowed = canAccessInstaller(currentUser);
  const homePath = getHomePathForUser(currentUser);
  const boardPath = getBoardPathForUser(currentUser);
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
          className={`host-nav-link ${active === "board" ? "active" : ""} ${!boardAllowed ? "disabled" : ""}`}
          onClick={() => {
            if (!boardAllowed) return;
            goTo(boardPath);
          }}
          disabled={!boardAllowed}
        >
          Installation Board
        </button>
        <button
          type="button"
          className={`host-nav-link ${active === "installer" ? "active" : ""} ${!installerAllowed ? "disabled" : ""}`}
          onClick={() => {
            if (!installerAllowed) return;
            goTo(installerPath);
          }}
          disabled={!installerAllowed}
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

function PermissionsPanel({ currentUser, users, savingKey, onChangePermission }) {
  const visibleUsers = [...users].sort((left, right) => left.displayName.localeCompare(right.displayName));

  return (
    <section className="panel permissions-panel">
      <div className="permissions-head">
        <h3>Permissions</h3>
        <p>Choose each person&apos;s access level for the Installation Board and Subcontractor Directory.</p>
      </div>
      <div className="permissions-grid">
        {visibleUsers.map((user) => {
          const isSelf = user.id === currentUser.id;
          const boardPermission = getPermissionForApp(user, "board");
          const installerPermission = getPermissionForApp(user, "installer");

          return (
            <article key={user.id} className="permissions-user-card">
              <div className="permissions-user-head">
                <strong>{user.displayName}</strong>
                {isSelf ? <span className="permissions-owner-pill">Owner</span> : null}
              </div>

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
            </article>
          );
        })}
      </div>
    </section>
  );
}

function HostLandingPage({ currentUser, onLogout, users, savingKey, onChangePermission }) {
  const [permissionsOpen, setPermissionsOpen] = useState(false);

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
          {currentUser?.canManagePermissions ? (
            <div className="host-landing-tools">
              <button className="ghost-button permissions-launch-button" type="button" onClick={() => setPermissionsOpen(true)}>
                Manage permissions
              </button>
            </div>
          ) : null}
          <div className="host-landing-actions">
            <button
              className={`host-launch-card ${!canAccessInstaller(currentUser) ? "disabled" : ""}`}
              type="button"
              onClick={() => {
                if (!canAccessInstaller(currentUser)) return;
                goTo("/installer");
              }}
              disabled={!canAccessInstaller(currentUser)}
            >
              <strong>Subcontractor Database</strong>
            </button>

            <button
              className={`host-launch-card ${!canAccessBoard(currentUser) ? "disabled" : ""}`}
              type="button"
              onClick={() => {
                if (!canAccessBoard(currentUser)) return;
                goTo(getBoardPathForUser(currentUser));
              }}
              disabled={!canAccessBoard(currentUser)}
            >
              <strong>Installation Board</strong>
            </button>
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
            />
          </div>
        </div>
      ) : null}
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

            <button className="host-launch-card" type="button" onClick={() => goTo(getBoardPathForUser(currentUser))}>
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
  const [currentUser, setCurrentUser] = useState(null);
  const [authChecked, setAuthChecked] = useState(false);
  const [loginUsers, setLoginUsers] = useState([]);
  const [loginDisplayName, setLoginDisplayName] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginError, setLoginError] = useState("");
  const [permissionSavingKey, setPermissionSavingKey] = useState("");
  const [previousMonthDepth, setPreviousMonthDepth] = useState(0);
  const [futureMonthDepth, setFutureMonthDepth] = useState(0);
  const dragPreviewRef = useRef(null);
  const transparentDragImageRef = useRef(null);
  const dragPositionRef = useRef({ x: 0, y: 0 });
  const boardEditable = canEditBoard(currentUser);
  const installerEditable = canEditInstaller(currentUser);
  const hostShellMode = usesHostShell(currentUser);
  const isClientMode = currentUser ? !boardEditable : false;
  const showInstallerDirectory = Boolean(currentUser && canAccessInstaller(currentUser) && isInstallerRoute);
  const showBoard = Boolean(
    currentUser &&
      canAccessBoard(currentUser) &&
      ((boardEditable && isBoardRoute) || (!boardEditable && isClientBoardRoute))
  );
  const showHostLanding = Boolean(currentUser && hostShellMode && !isInstallerRoute && !isBoardRoute && !isClientBoardRoute);
  const showClientLanding = Boolean(currentUser && !hostShellMode && canAccessBoard(currentUser) && !isClientBoardRoute);

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
    if (!currentUser) return;
    const nextHomePath = getHomePathForUser(currentUser);
    const nextBoardPath = getBoardPathForUser(currentUser);

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

    if (!hostShellMode && !isClientRoute) {
      window.location.replace(nextHomePath);
      return;
    }

    if ((isBoardRoute || isClientBoardRoute) && nextBoardPath !== window.location.pathname) {
      window.location.replace(nextBoardPath);
    }
  }, [currentUser, isClientRoute, isClientBoardRoute, isInstallerRoute, isBoardRoute, hostShellMode]);

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

  async function handleSubmit(event) {
    event.preventDefault();
    if (isClientMode) return;

    if (!form.customerName.trim()) {
      setMessage(createMessage("Enter the customer name.", "error"));
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
    return (
      <HostLandingPage
        currentUser={currentUser}
        onLogout={handleLogout}
        users={loginUsers}
        savingKey={permissionSavingKey}
        onChangePermission={handlePermissionChange}
      />
    );
  }

  if (showClientLanding) {
    return <ClientLandingPage currentUser={currentUser} onLogout={handleLogout} />;
  }

  if (showInstallerDirectory) {
    return <InstallerDirectoryHost currentUser={currentUser} onLogout={handleLogout} readOnly={!installerEditable} />;
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

                <section className="unscheduled-section panel">
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
              <div className="detail-card">
                <strong>Placeholder</strong>
                <p>{activeClientJob.isPlaceholder ? "Yes" : "No"}</p>
              </div>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}
