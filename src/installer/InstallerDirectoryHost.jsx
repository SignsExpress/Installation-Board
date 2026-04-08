import { useEffect, useMemo, useState } from "react";
import InteractiveMap from "./InteractiveMap";
import { DEFAULT_INSTALLER_FORM, FILTER_OPTIONS, REGIONS } from "./directoryData";
import { getRegionFromAddress, normalizeInstaller } from "./directoryUtils";

function createMessage(text, tone = "info") {
  return { text, tone, id: `${Date.now()}-${Math.random()}` };
}

function renderRating(value) {
  const rounded = Math.max(0, Math.min(5, Number(value) || 0));
  return "★★★★★".slice(0, rounded) + "☆☆☆☆☆".slice(rounded, 5);
}

export default function InstallerDirectoryHost({ currentUser, onLogout }) {
  const [installers, setInstallers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState(null);
  const [search, setSearch] = useState("");
  const [selectedRegion, setSelectedRegion] = useState(null);
  const [hoveredRegion, setHoveredRegion] = useState(null);
  const [selectedTag, setSelectedTag] = useState("");
  const [editingId, setEditingId] = useState("");
  const [form, setForm] = useState(DEFAULT_INSTALLER_FORM);

  useEffect(() => {
    let active = true;

    async function loadInstallers() {
      try {
        const response = await fetch("/api/installers");
        if (!response.ok) throw new Error("Could not load installers.");
        const payload = await response.json();
        if (active) {
          setInstallers(Array.isArray(payload) ? payload.map(normalizeInstaller) : []);
        }
      } catch (error) {
        console.error(error);
        if (active) setMessage(createMessage("Could not load the installer directory.", "error"));
      } finally {
        if (active) setLoading(false);
      }
    }

    loadInstallers();
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (!message) return undefined;
    const timer = window.setTimeout(() => setMessage(null), 3000);
    return () => window.clearTimeout(timer);
  }, [message]);

  const installersByRegion = useMemo(() => {
    return installers.reduce((accumulator, installer) => {
      installer.regions.forEach((regionId) => {
        if (!accumulator[regionId]) accumulator[regionId] = [];
        accumulator[regionId].push(installer);
      });
      return accumulator;
    }, {});
  }, [installers]);

  const filteredInstallers = useMemo(() => {
    const needle = search.trim().toLowerCase();

    return installers.filter((installer) => {
      if (selectedRegion && !installer.regions.includes(selectedRegion)) return false;
      if (selectedTag && !installer.tags.includes(selectedTag)) return false;
      if (!needle) return true;

      return [
        installer.name,
        installer.company,
        installer.phone,
        installer.email,
        installer.address,
        installer.notes,
        installer.regions.join(" "),
        installer.tags.join(" ")
      ]
        .join(" ")
        .toLowerCase()
        .includes(needle);
    });
  }, [installers, search, selectedRegion, selectedTag]);

  function navigate(path) {
    window.location.assign(path);
  }

  function resetForm() {
    setEditingId("");
    setForm(DEFAULT_INSTALLER_FORM);
  }

  function startEdit(installer) {
    setEditingId(installer.id);
    setForm({
      id: installer.id,
      name: installer.name || "",
      company: installer.company || "",
      phone: installer.phone || "",
      email: installer.email || "",
      address: installer.address || "",
      notes: installer.notes || "",
      rating: installer.rating || 0,
      regions: Array.isArray(installer.regions) ? installer.regions : [],
      tags: Array.isArray(installer.tags) ? installer.tags : []
    });
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function handleSubmit(event) {
    event.preventDefault();
    if (!form.name.trim()) {
      setMessage(createMessage("Installer name is required.", "error"));
      return;
    }

    try {
      setSaving(true);
      const response = await fetch("/api/installers", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...form,
          id: editingId || form.id || undefined
        })
      });

      if (!response.ok) throw new Error("Could not save installer.");
      const payload = await response.json();
      setInstallers(Array.isArray(payload) ? payload.map(normalizeInstaller) : []);
      setMessage(createMessage(editingId ? "Installer updated." : "Installer added.", "success"));
      resetForm();
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not save installer.", "error"));
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete(installerId) {
    try {
      const response = await fetch(`/api/installers/${installerId}`, { method: "DELETE" });
      if (!response.ok) throw new Error("Could not delete installer.");
      const payload = await response.json();
      setInstallers(Array.isArray(payload) ? payload.map(normalizeInstaller) : []);
      if (editingId === installerId) resetForm();
      setMessage(createMessage("Installer removed.", "success"));
    } catch (error) {
      console.error(error);
      setMessage(createMessage("Could not delete installer.", "error"));
    }
  }

  function toggleRegion(regionId) {
    setForm((current) => ({
      ...current,
      regions: current.regions.includes(regionId)
        ? current.regions.filter((item) => item !== regionId)
        : [...current.regions, regionId]
    }));
  }

  function toggleTag(tagId) {
    setForm((current) => ({
      ...current,
      tags: current.tags.includes(tagId)
        ? current.tags.filter((item) => item !== tagId)
        : [...current.tags, tagId]
    }));
  }

  function handleAddressBlur() {
    if (form.regions.length > 0) return;
    const match = getRegionFromAddress(form.address);
    if (!match) return;

    setForm((current) => ({
      ...current,
      regions: current.regions.includes(match.id) ? current.regions : [...current.regions, match.id]
    }));
  }

  return (
    <div className="app-shell">
      <div className="page">
        <section className="hero">
          <div className="hero-brand">
            <img className="hero-logo" src="/branding/signs-express-logo.svg" alt="Signs Express" />
            <div className="hero-copy">
              <p className="panel-kicker">Host only</p>
              <h1>Subcontractor Installer Database</h1>
              <p className="muted">Signed in as {currentUser.displayName}</p>
            </div>
          </div>
          <div className="hero-user">
            <div className="host-top-actions">
              <button className="ghost-button" type="button" onClick={() => navigate("/")}>
                Home
              </button>
              <button className="ghost-button" type="button" onClick={() => navigate("/board")}>
                Installation Board
              </button>
              <button className="ghost-button" type="button" onClick={onLogout}>
                Log out
              </button>
            </div>
          </div>
        </section>

        {message ? <div className={`flash ${message.tone === "error" ? "error" : "success"}`}>{message.text}</div> : null}

        <div className="layout installer-layout">
          <section className="panel installer-editor-panel">
            <div className="panel-head">
              <div>
                <p className="panel-kicker">Manage installers</p>
                <h2>{editingId ? "Edit installer" : "Add installer"}</h2>
              </div>
            </div>

            <form className="job-form" onSubmit={handleSubmit}>
              <label>
                Name
                <input
                  type="text"
                  value={form.name}
                  onChange={(event) => setForm((current) => ({ ...current, name: event.target.value }))}
                />
              </label>

              <label>
                Company
                <input
                  type="text"
                  value={form.company}
                  onChange={(event) => setForm((current) => ({ ...current, company: event.target.value }))}
                />
              </label>

              <div className="split-fields">
                <label>
                  Phone
                  <input
                    type="text"
                    value={form.phone}
                    onChange={(event) => setForm((current) => ({ ...current, phone: event.target.value }))}
                  />
                </label>

                <label>
                  Email / Website
                  <input
                    type="text"
                    value={form.email}
                    onChange={(event) => setForm((current) => ({ ...current, email: event.target.value }))}
                  />
                </label>
              </div>

              <label>
                Address
                <input
                  type="text"
                  value={form.address}
                  onBlur={handleAddressBlur}
                  onChange={(event) => setForm((current) => ({ ...current, address: event.target.value }))}
                />
              </label>

              <label>
                Notes
                <textarea
                  rows="5"
                  value={form.notes}
                  onChange={(event) => setForm((current) => ({ ...current, notes: event.target.value }))}
                />
              </label>

              <label>
                Rating
                <div className="rating-picker">
                  {[1, 2, 3, 4, 5].map((value) => (
                    <button
                      key={value}
                      type="button"
                      className={`rating-button ${Number(form.rating) === value ? "active" : ""}`}
                      onClick={() => setForm((current) => ({ ...current, rating: value }))}
                    >
                      {value} star{value === 1 ? "" : "s"}
                    </button>
                  ))}
                </div>
              </label>

              <label>
                Regions
                <div className="installer-filter-grid">
                  {REGIONS.map((region) => (
                    <button
                      key={region.id}
                      type="button"
                      className={`installer-filter-chip ${form.regions.includes(region.id) ? "active" : ""}`}
                      onClick={() => toggleRegion(region.id)}
                    >
                      {region.label}
                    </button>
                  ))}
                </div>
              </label>

              <label>
                Tags
                <div className="installer-filter-grid">
                  {FILTER_OPTIONS.map((tag) => (
                    <button
                      key={tag.id}
                      type="button"
                      className={`installer-filter-chip ${form.tags.includes(tag.id) ? "active" : ""}`}
                      onClick={() => toggleTag(tag.id)}
                    >
                      {tag.label}
                    </button>
                  ))}
                </div>
              </label>

              <div className="form-actions">
                <button className="primary-button" type="submit" disabled={saving}>
                  {saving ? "Saving..." : editingId ? "Update installer" : "Add installer"}
                </button>
                <button className="ghost-button" type="button" onClick={resetForm}>
                  Clear
                </button>
              </div>
            </form>
          </section>

          <section className="panel installer-directory-panel">
            <div className="panel-head">
              <div>
                <p className="panel-kicker">Directory</p>
                <h2>Browse installers</h2>
              </div>
              <div className="board-range">{filteredInstallers.length} shown</div>
            </div>

            <div className="installer-toolbar">
              <input
                type="text"
                value={search}
                placeholder="Search by installer, company, phone, address or notes"
                onChange={(event) => setSearch(event.target.value)}
              />
              <button className="ghost-button" type="button" onClick={() => { setSearch(""); setSelectedRegion(null); setSelectedTag(""); }}>
                Reset filters
              </button>
            </div>

            <InteractiveMap
              installersByRegion={installersByRegion}
              selectedRegion={selectedRegion}
              setSelectedRegion={setSelectedRegion}
              hoveredRegion={hoveredRegion}
              setHoveredRegion={setHoveredRegion}
            />

            <div className="installer-filter-grid compact">
              {REGIONS.map((region) => (
                <button
                  key={region.id}
                  type="button"
                  className={`installer-filter-chip ${selectedRegion === region.id ? "active" : ""}`}
                  onMouseEnter={() => setHoveredRegion(region.id)}
                  onMouseLeave={() => setHoveredRegion(null)}
                  onClick={() => setSelectedRegion(selectedRegion === region.id ? null : region.id)}
                >
                  {region.label}
                </button>
              ))}
            </div>

            <div className="installer-filter-grid compact">
              {FILTER_OPTIONS.map((tag) => (
                <button
                  key={tag.id}
                  type="button"
                  className={`installer-filter-chip ${selectedTag === tag.id ? "active" : ""}`}
                  onClick={() => setSelectedTag(selectedTag === tag.id ? "" : tag.id)}
                >
                  {tag.label}
                </button>
              ))}
            </div>

            {loading ? (
              <div className="board-loading">Loading the installer directory...</div>
            ) : (
              <div className="installer-card-grid">
                {filteredInstallers.map((installer) => (
                  <article key={installer.id} className="installer-card">
                    <div className="installer-card-head">
                      <div>
                        <strong>{installer.name}</strong>
                        <p>{installer.company || "No company listed"}</p>
                      </div>
                      <span className="installer-rating">{renderRating(installer.rating)}</span>
                    </div>

                    <div className="installer-card-body">
                      <p><b>Phone:</b> {installer.phone || "-"}</p>
                      <p><b>Email / Website:</b> {installer.email || "-"}</p>
                      <p><b>Address:</b> {installer.address || "-"}</p>
                      <p><b>Notes:</b> {installer.notes || "-"}</p>
                    </div>

                    <div className="installer-badge-row">
                      {installer.regions.map((regionId) => {
                        const region = REGIONS.find((entry) => entry.id === regionId);
                        return (
                          <span key={`${installer.id}-${regionId}`} className="installer-filter-chip active">
                            {region?.label || regionId}
                          </span>
                        );
                      })}
                      {installer.tags.map((tag) => (
                        <span key={`${installer.id}-${tag}`} className="installer-filter-chip">
                          {FILTER_OPTIONS.find((entry) => entry.id === tag)?.label || tag}
                        </span>
                      ))}
                    </div>

                    <div className="job-actions">
                      <button className="text-button" type="button" onClick={() => startEdit(installer)}>
                        Edit
                      </button>
                      <button className="text-button danger" type="button" onClick={() => handleDelete(installer.id)}>
                        Delete
                      </button>
                    </div>
                  </article>
                ))}

                {!filteredInstallers.length ? (
                  <div className="board-loading compact">No installers matched those filters.</div>
                ) : null}
              </div>
            )}
          </section>
        </div>
      </div>
    </div>
  );
}
