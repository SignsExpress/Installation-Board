import { useEffect, useMemo, useState } from "react";
import InteractiveMap from "./InteractiveMap";
import StarRating from "./StarRating";
import { DEFAULT_INSTALLER_FORM, FILTER_OPTIONS, REGION_MAP, REGIONS } from "./directoryData";
import { createId, getRegionFromAddress, normalizeInstaller } from "./directoryUtils";

const DEMO_COMPANIES = [
  "Lorem Installations Ltd",
  "Ipsum Access Group",
  "Dolor Sign Systems",
  "Sit Amet Services",
  "Consectetur Display Co"
];

const DEMO_NAMES = ["Alex Mercer", "Jordan Blake", "Taylor Reed", "Casey Morgan", "Jamie Carter"];

function cls(...items) {
  return items.filter(Boolean).join(" ");
}

function FilterButtons({ activeIds, onToggle, className = "" }) {
  return (
    <div className={cls("chip-grid filter-bar", className)}>
      {FILTER_OPTIONS.map((filter) => (
        <button
          key={filter.id}
          type="button"
          className={cls("chip-button", activeIds.includes(filter.id) && "active")}
          onClick={() => onToggle(filter.id)}
        >
          {filter.label}
        </button>
      ))}
    </div>
  );
}

function createMessage(text, tone = "success") {
  return { text, tone, id: `${Date.now()}-${Math.random()}` };
}

export default function InstallerDirectoryHostV2({ currentUser, onLogout, readOnly = false, navigation = null }) {
  const [installers, setInstallers] = useState([]);
  const [requests, setRequests] = useState([]);
  const [selectedRegion, setSelectedRegion] = useState(null);
  const [hoveredRegion, setHoveredRegion] = useState(null);
  const [activeInstaller, setActiveInstaller] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [reviewingRequestId, setReviewingRequestId] = useState(null);
  const [form, setForm] = useState(DEFAULT_INSTALLER_FORM);
  const [message, setMessage] = useState(null);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState("");
  const [activeFilters, setActiveFilters] = useState([]);
  const [showBy, setShowBy] = useState("business");
  const [trailMode, setTrailMode] = useState(false);
  const [lookupAddress, setLookupAddress] = useState("");
  const [lookupMessage, setLookupMessage] = useState("");
  const [serverInfo, setServerInfo] = useState(null);
  const [isFormOpen, setIsFormOpen] = useState(false);

  useEffect(() => {
    let active = true;

    async function loadAll() {
      setLoading(true);
      try {
        const [installersResponse, requestsResponse, statusResponse] = await Promise.all([
          fetch("/api/installers"),
          readOnly ? Promise.resolve({ ok: false }) : fetch("/api/requests"),
          fetch("/api/installers/status")
        ]);

        if (!installersResponse.ok) {
          throw new Error("Could not load installer data.");
        }

        const installersPayload = await installersResponse.json();
        const requestsPayload = requestsResponse.ok ? await requestsResponse.json() : [];
        const statusPayload = statusResponse.ok ? await statusResponse.json() : null;

        if (!active) return;
        setInstallers((Array.isArray(installersPayload) ? installersPayload : []).map(normalizeInstaller));
        setRequests(Array.isArray(requestsPayload) ? requestsPayload : []);
        setServerInfo(statusPayload);
      } catch (error) {
        console.error(error);
        if (active) setMessage(createMessage("Could not reach the shared server.", "error"));
      } finally {
        if (active) setLoading(false);
      }
    }

    loadAll();
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (!message) return undefined;
    const timer = window.setTimeout(() => setMessage(null), 2200);
    return () => window.clearTimeout(timer);
  }, [message]);

  const filteredInstallers = useMemo(
    () =>
      installers
        .filter((installer) => !activeFilters.length || activeFilters.every((tagId) => installer.tags.includes(tagId)))
        .filter((installer) => {
          if (!searchTerm.trim()) return true;
          const haystack = [installer.company, installer.name, installer.phone, installer.email, installer.address, installer.notes]
            .join(" ")
            .toLowerCase();
          return haystack.includes(searchTerm.trim().toLowerCase());
        })
        .sort((left, right) => {
          const leftPrimary = showBy === "name" ? left.name || left.company || "" : left.company || left.name || "";
          const rightPrimary = showBy === "name" ? right.name || right.company || "" : right.company || right.name || "";
          const primaryCompare = leftPrimary.localeCompare(rightPrimary, undefined, { sensitivity: "base" });
          if (primaryCompare !== 0) return primaryCompare;

          const leftSecondary = showBy === "name" ? left.company || "" : left.name || "";
          const rightSecondary = showBy === "name" ? right.company || "" : right.name || "";
          return leftSecondary.localeCompare(rightSecondary, undefined, { sensitivity: "base" });
        }),
    [activeFilters, installers, searchTerm, showBy]
  );

  const installersByRegion = useMemo(() => {
    const map = {};
    REGIONS.forEach((region) => {
      map[region.id] = filteredInstallers.filter((installer) => installer.regions.includes(region.id));
    });
    return map;
  }, [filteredInstallers]);

  const visibleInstallers = useMemo(
    () => (selectedRegion ? filteredInstallers.filter((installer) => installer.regions.includes(selectedRegion)) : filteredInstallers),
    [filteredInstallers, selectedRegion]
  );

  function navigate(path) {
    window.location.assign(path);
  }

  const boardPath = currentUser?.permissions?.board === "admin" ? "/board" : "/client/board";
  const holidaysAllowed = currentUser?.permissions?.holidays && currentUser.permissions.holidays !== "none";

  function resetForm() {
    setForm(DEFAULT_INSTALLER_FORM);
    setEditingId(null);
    setReviewingRequestId(null);
  }

  function showSavedMessage(text) {
    setMessage(createMessage(text, "success"));
  }

  async function saveInstaller(payload) {
    const response = await fetch("/api/installers", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    if (!response.ok) throw new Error("Could not save installer.");
    setInstallers((await response.json()).map(normalizeInstaller));
  }

  async function removeInstaller(id) {
    const response = await fetch(`/api/installers/${id}`, { method: "DELETE" });
    if (!response.ok) throw new Error("Could not delete installer.");
    setInstallers((await response.json()).map(normalizeInstaller));
  }

  async function submitForm(event) {
    event.preventDefault();
    if (!form.name.trim() && !form.company.trim()) {
      setMessage(createMessage("Add a company or contact name before saving.", "error"));
      return;
    }

    const payload = {
      id: editingId || createId(),
      name: form.name.trim(),
      company: form.company.trim(),
      phone: form.phone.trim(),
      email: form.email.trim(),
      address: form.address.trim(),
      notes: form.notes.trim(),
      rating: Math.max(0, Math.min(5, Number(form.rating) || 0)),
      regions: form.regions,
      tags: form.tags
    };

    try {
      await saveInstaller(payload);

      if (reviewingRequestId) {
        const response = await fetch(`/api/requests/${reviewingRequestId}`, { method: "DELETE" });
        if (response.ok) {
          setRequests(await response.json());
        }
      }

      showSavedMessage(
        editingId ? "Installer updated." : reviewingRequestId ? "Request approved and saved." : "Installer saved."
      );
      resetForm();
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not save installer.", "error"));
    }
  }

  function editInstaller(installer) {
    if (readOnly) return;
    setForm({
      id: installer.id || "",
      name: installer.name || "",
      company: installer.company || "",
      phone: installer.phone || "",
      email: installer.email || "",
      address: installer.address || "",
      notes: installer.notes || "",
      rating: Number(installer.rating) || 0,
      regions: installer.regions || [],
      tags: installer.tags || []
    });
    setEditingId(installer.id);
    setIsFormOpen(true);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function handleDelete(id) {
    if (readOnly) return;
    try {
      await removeInstaller(id);
      showSavedMessage("Installer deleted.");
      if (activeInstaller?.id === id) setActiveInstaller(null);
      if (editingId === id) resetForm();
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not delete installer.", "error"));
    }
  }

  function toggleArrayValue(key, value) {
    setForm((current) => ({
      ...current,
      [key]: current[key].includes(value)
        ? current[key].filter((item) => item !== value)
        : [...current[key], value]
    }));
  }

  function toggleFilter(filterId) {
    setActiveFilters((current) =>
      current.includes(filterId) ? current.filter((id) => id !== filterId) : [...current, filterId]
    );
  }

  function getDisplayInstaller(installer, index) {
    if (!trailMode) {
      return installer;
    }

    return {
      ...installer,
      company: DEMO_COMPANIES[index % DEMO_COMPANIES.length],
      name: DEMO_NAMES[index % DEMO_NAMES.length],
      phone: "07*** *** ***",
      email: `demo${index + 1}@example.com`,
      address: "Lorem House, Ipsum Business Park",
      notes:
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Demo mode hides live subcontractor details without changing the saved data."
    };
  }

  function handleLookupRegion() {
    const match = getRegionFromAddress(lookupAddress);
    if (!match) {
      setLookupMessage("No region match found. Try a postcode or fuller address.");
      return;
    }
    setSelectedRegion(match.id);
    setLookupMessage(`Matched to ${match.label}. Region filter applied.`);
  }

  function reviewRequest(requestItem) {
    if (readOnly) return;
    setForm({
      id: requestItem.id || "",
      name: requestItem.name || "",
      company: requestItem.company || "",
      phone: requestItem.phone || "",
      email: requestItem.email || "",
      address: requestItem.address || "",
      notes: requestItem.notes || "",
      rating: Number(requestItem.rating) || 0,
      regions: requestItem.regions || [],
      tags: requestItem.tags || []
    });
    setReviewingRequestId(requestItem.id);
    setEditingId(null);
    setIsFormOpen(true);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  return (
    <div className="app-shell installer-host-view">
      <div className="page">
        {navigation || (
        <header className="host-nav-shell">
          <nav className="host-nav">
            <div className="host-nav-inner">
              <button type="button" className="host-nav-brand" onClick={() => navigate("/")} aria-label="Go to home">
                <img src="/branding/signs-express-logo.svg" alt="Signs Express" className="host-nav-brand-logo" />
              </button>
              <div className="host-nav-links">
                <button type="button" className="host-nav-link" onClick={() => navigate("/")}>
                  <span className="host-nav-link-label">Home</span>
                </button>
                <button type="button" className="host-nav-link" onClick={() => navigate(boardPath)}>
                  <span className="host-nav-link-label">Installation Board</span>
                </button>
                <button
                  type="button"
                  className={cls("host-nav-link", holidaysAllowed ? "" : "disabled")}
                  onClick={() => {
                    if (!holidaysAllowed) return;
                    navigate("/holidays");
                  }}
                  disabled={!holidaysAllowed}
                >
                  <span className="host-nav-link-label">Holidays</span>
                </button>
                <button type="button" className="host-nav-link active" onClick={() => navigate("/installer")}>
                  <span className="host-nav-link-label">Subcontractor Directory</span>
                </button>
                <button type="button" className="host-nav-link" onClick={() => navigate("/notifications")}>
                  <span className="host-nav-link-label">Notifications</span>
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
        )}

        <div className="workspace-grid">
          <section className="card card-large map-panel-card">
            <div className="section-head">
              <div />
              <div className="map-panel-actions">
                {message ? (
                  <div className={cls("stat-pill", message.tone === "error" ? "" : "success")}>
                    <strong>{message.text}</strong>
                  </div>
                ) : null}
              </div>
            </div>
            <div className="finder-box">
              <label>
                Find nearest location
                <div className="lookup-row">
                  <input
                    value={lookupAddress}
                    onChange={(event) => setLookupAddress(event.target.value)}
                    placeholder="Enter postcode, town or address"
                  />
                  <button type="button" className="primary-button" onClick={handleLookupRegion}>
                    Find nearest location
                  </button>
                  <button
                    type="button"
                    className="ghost-button lookup-clear-button"
                    onClick={() => setSelectedRegion(null)}
                  >
                    Clear region filter
                  </button>
                </div>
              </label>
              {lookupMessage ? <div className="lookup-message">{lookupMessage}</div> : null}
            </div>
            <div className="map-grid">
              <InteractiveMap
                installersByRegion={installersByRegion}
                selectedRegion={selectedRegion}
                setSelectedRegion={setSelectedRegion}
                hoveredRegion={hoveredRegion}
                setHoveredRegion={setHoveredRegion}
              />
              <aside className="side-panel">
                <div className="region-list">
                  {REGIONS.map((region) => (
                    <button
                      key={region.id}
                      className={cls("region-button", selectedRegion === region.id && "active")}
                      onClick={() => setSelectedRegion((current) => (current === region.id ? null : region.id))}
                    >
                      <span>
                        {region.number}. {region.label}
                      </span>
                      <b
                        className={cls(
                          "region-count",
                          (installersByRegion[region.id]?.length || 0) === 0 && "is-zero"
                        )}
                      >
                        {installersByRegion[region.id]?.length || 0}
                      </b>
                    </button>
                  ))}
                  <button
                    type="button"
                    className={cls("region-button trail-region-button", trailMode && "active")}
                    onClick={() => setTrailMode((current) => !current)}
                  >
                    <span>Trail Mode</span>
                    <b className="region-count">{trailMode ? "On" : "Off"}</b>
                  </button>
                </div>
              </aside>
            </div>
          </section>

          <section className="card contacts-panel-card">
            <div className="section-head section-head-stack contacts-panel-head">
              <div className="toolbar-block">
                <div className="toolbar-top-row">
                  <label className="search-box">
                    <input
                      value={searchTerm}
                      onChange={(event) => setSearchTerm(event.target.value)}
                      placeholder="Search by company, name, phone, email or notes"
                    />
                  </label>
                  <div className="subcontractor-count-pill" aria-label={`${visibleInstallers.length} Subcontractors`}>
                    <strong>{visibleInstallers.length} Subcontractors</strong>
                  </div>
                </div>
                <div className="toolbar-inline-row">
                  <div className="inline-filter-group">
                    <span className="show-by-label">Display by</span>
                    <div className="chip-grid compact-chip-grid inline-chip-grid">
                      <button
                        type="button"
                        className={cls("chip-button", showBy === "business" && "active")}
                        onClick={() => setShowBy("business")}
                      >
                        Business Name
                      </button>
                      <button
                        type="button"
                        className={cls("chip-button", showBy === "name" && "active")}
                        onClick={() => setShowBy("name")}
                      >
                        Name
                      </button>
                    </div>
                  </div>
                  <div className="inline-filter-group">
                    <span className="show-by-label">Filters</span>
                    <FilterButtons
                      activeIds={activeFilters}
                      onToggle={toggleFilter}
                      className="contacts-filter-bar inline-chip-grid"
                    />
                  </div>
                </div>
              </div>
            </div>

            {loading ? (
              <div className="empty-state compact-empty-state">Loading saved installers...</div>
            ) : visibleInstallers.length === 0 ? (
              <div className="empty-state compact-empty-state">
                No installers match the current region, search or filter selection.
              </div>
            ) : (
              <div className="contacts-scroll-area">
                <div className="contacts-list-grid">
                  {visibleInstallers.map((installer, index) => {
                    const displayInstaller = getDisplayInstaller(installer, index);
                    return (
                      <article key={installer.id} className="installer-card compact-installer-card">
                        <button className="installer-main" onClick={() => setActiveInstaller(installer)}>
                          <div className="installer-top">
                            <div>
                              <h3>
                                {showBy === "name"
                                  ? displayInstaller.name || displayInstaller.company || "No name added"
                                  : displayInstaller.company || displayInstaller.name || "No company added"}
                              </h3>
                              <p>
                                {showBy === "name"
                                  ? displayInstaller.company || "No company added"
                                  : displayInstaller.name || "No contact name added"}
                              </p>
                              <StarRating rating={displayInstaller.rating || 0} size={14} />
                            </div>
                          </div>
                          <div className="contact-list compact-contact-list">
                            {displayInstaller.phone ? (
                              <div>
                                <span>{displayInstaller.phone}</span>
                              </div>
                            ) : null}
                            {displayInstaller.email ? (
                              <div>
                                <span>{displayInstaller.email}</span>
                              </div>
                            ) : null}
                            {displayInstaller.address ? (
                              <div>
                                <span>{displayInstaller.address}</span>
                              </div>
                            ) : null}
                          </div>
                        </button>
                        <div className="tag-list compact-tag-list">
                          {installer.tags.map((tagId) => (
                            <span key={tagId} className="tag">
                              {FILTER_OPTIONS.find((filter) => filter.id === tagId)?.label || tagId}
                            </span>
                          ))}
                          {installer.regions.map((regionId) => (
                            <span key={regionId} className="tag">
                              {REGION_MAP[regionId]?.label}
                            </span>
                          ))}
                        </div>
                        {!readOnly ? (
                          <div className="card-actions form-actions compact-card-actions">
                            <button className="ghost-button" onClick={() => editInstaller(installer)} disabled={trailMode}>
                              Edit
                            </button>
                            <button className="danger-button" onClick={() => handleDelete(installer.id)}>
                              Delete
                            </button>
                          </div>
                        ) : null}
                      </article>
                    );
                  })}
                </div>
              </div>
            )}
          </section>
        </div>

        {!readOnly ? (
        <section className="card form-card request-form-card collapsed-form-card">
          <button
            type="button"
            className="subcontractor-toggle"
            onClick={() => setIsFormOpen((current) => !current)}
          >
            <span>Add Subcontractor</span>
            <span className={cls("subcontractor-toggle-icon", isFormOpen && "open")}>⌄</span>
          </button>

          {isFormOpen ? (
            <div className="subcontractor-form-wrap">
              <div className="subcontractor-form-top">
                <div className="subcontractor-stats">
                  <div className="stat-pill">
                    <span>Installers</span>
                    <strong>{serverInfo?.installers ?? installers.length}</strong>
                  </div>
                  <div className="stat-pill">
                    <span>Pending requests</span>
                    <strong>{serverInfo?.requests ?? requests.length}</strong>
                  </div>
                </div>
              </div>

              {requests.length > 0 ? (
                <div className="request-panel">
                  <div className="request-panel-head">
                    <strong>Pending requests</strong>
                    <span>{requests.length}</span>
                  </div>
                  <div className="request-list">
                    {requests.map((requestItem) => (
                      <button
                        key={requestItem.id}
                        type="button"
                        className="request-item"
                        onClick={() => reviewRequest(requestItem)}
                      >
                        <strong>{requestItem.company || requestItem.name || "Unnamed request"}</strong>
                        <span>
                          {requestItem.name || "No contact name"}
                          {requestItem.createdAt
                            ? ` - ${new Date(requestItem.createdAt).toLocaleDateString("en-GB")}`
                            : ""}
                        </span>
                      </button>
                    ))}
                  </div>
                </div>
              ) : null}

              {trailMode ? (
                <div className="lookup-message">
                  Trail Mode is on. Live installer details are masked and editing is disabled.
                </div>
              ) : null}

              <form className="installer-form" onSubmit={submitForm}>
                <label>
                  Company
                  <input
                    value={form.company}
                    onChange={(event) => setForm({ ...form, company: event.target.value })}
                    placeholder="ABC Installations Ltd"
                  />
                </label>
                <label>
                  Contact name
                  <input
                    value={form.name}
                    onChange={(event) => setForm({ ...form, name: event.target.value })}
                    placeholder="John Smith"
                  />
                </label>
                <div className="split">
                  <label>
                    Phone
                    <input
                      value={form.phone}
                      onChange={(event) => setForm({ ...form, phone: event.target.value })}
                      placeholder="07700 900000"
                    />
                  </label>
                  <label>
                    Email
                    <input
                      value={form.email}
                      onChange={(event) => setForm({ ...form, email: event.target.value })}
                      placeholder="name@company.co.uk"
                    />
                  </label>
                </div>
                <label>
                  Address
                  <input
                    value={form.address}
                    onChange={(event) => setForm({ ...form, address: event.target.value })}
                    onBlur={() => {
                      if (form.regions.length > 0) return;
                      const match = getRegionFromAddress(form.address);
                      if (!match) return;
                      setForm((current) => ({ ...current, regions: [...current.regions, match.id] }));
                    }}
                    placeholder="Town, county or full address"
                  />
                </label>
                <label>
                  Notes
                  <textarea
                    value={form.notes}
                    onChange={(event) => setForm({ ...form, notes: event.target.value })}
                    placeholder="Access kit, preferred work, coverage notes, RAMS info"
                    rows={3}
                  />
                </label>
                <div>
                  <span className="field-label">Rating</span>
                  <StarRating
                    rating={form.rating}
                    setRating={(value) => setForm({ ...form, rating: value })}
                    editable
                  />
                </div>
                <div>
                  <span className="field-label">Qualifications and status</span>
                  <FilterButtons activeIds={form.tags} onToggle={(tagId) => toggleArrayValue("tags", tagId)} />
                </div>
                <div>
                  <span className="field-label">Regions covered</span>
                  <div className="chip-grid">
                    {REGIONS.map((region) => (
                      <button
                        key={region.id}
                        type="button"
                        className={cls("chip-button", form.regions.includes(region.id) && "active")}
                        onClick={() => toggleArrayValue("regions", region.id)}
                      >
                        {region.number}. {region.label}
                      </button>
                    ))}
                  </div>
                </div>
                <div className="form-actions">
                  <button className="primary-button" type="submit" disabled={trailMode}>
                    {editingId ? "Update installer" : reviewingRequestId ? "Approve and save" : "Save installer"}
                  </button>
                  <button className="ghost-button" type="button" onClick={resetForm}>
                    Clear form
                  </button>
                </div>
              </form>
            </div>
          ) : null}
        </section>
        ) : null}
      </div>

      {activeInstaller ? (
        <div className="modal-backdrop">
          {(() => {
            const displayInstaller = getDisplayInstaller(
              activeInstaller,
              visibleInstallers.findIndex((installer) => installer.id === activeInstaller.id)
            );
            return (
              <div className="modal">
                <div className="modal-head">
                  <div>
                    <h3>{displayInstaller.company || "No company added"}</h3>
                    <p>{displayInstaller.name || "No contact name added"}</p>
                    <StarRating rating={displayInstaller.rating || 0} size={18} />
                  </div>
                  <button className="icon-button" onClick={() => setActiveInstaller(null)}>
                    x
                  </button>
                </div>
                <div className="modal-grid">
                  <div className="info-card">
                    <div className="mini-head">Contact</div>
                    {displayInstaller.name ? <p><strong>Name:</strong> {displayInstaller.name}</p> : null}
                    {displayInstaller.phone ? <p><strong>Phone:</strong> {displayInstaller.phone}</p> : null}
                    {displayInstaller.email ? <p><strong>Email:</strong> {displayInstaller.email}</p> : null}
                    {!displayInstaller.phone && !displayInstaller.email ? (
                      <p>No contact details added.</p>
                    ) : null}
                  </div>
                  <div className="info-card">
                    <div className="mini-head">Business</div>
                    <p><strong>Company:</strong> {displayInstaller.company || "Not added"}</p>
                    <p><strong>Address:</strong> {displayInstaller.address || "Not added"}</p>
                  </div>
                </div>
                <div className="info-card">
                  <div className="mini-head">Qualifications and coverage</div>
                  <div className="tag-list">
                    {activeInstaller.tags.map((tagId) => (
                      <span key={tagId} className="tag active-tag">
                        {FILTER_OPTIONS.find((filter) => filter.id === tagId)?.label || tagId}
                      </span>
                    ))}
                    {activeInstaller.regions.map((regionId) => (
                      <span key={regionId} className="tag active-tag">
                        {REGION_MAP[regionId]?.label}
                      </span>
                    ))}
                  </div>
                </div>
                <div className="info-card">
                  <div className="mini-head">Notes</div>
                  <p className="notes">{displayInstaller.notes || "No notes added."}</p>
                </div>
                {!readOnly ? (
                  <div className="form-actions">
                    <button
                      className="primary-button"
                      onClick={() => {
                        editInstaller(activeInstaller);
                        setActiveInstaller(null);
                      }}
                      disabled={trailMode}
                    >
                      Edit installer
                    </button>
                  </div>
                ) : null}
              </div>
            );
          })()}
        </div>
      ) : null}
    </div>
  );
}
