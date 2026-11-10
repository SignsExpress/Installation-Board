import { useEffect, useMemo, useState } from "react";
import "./scouts.css";

const DELIVERY_FEE = 12.5;

function makeSign(groupName = "") {
  return {
    id: typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : `sign-${Date.now()}-${Math.random()}`,
    groupName,
    width: 600,
    height: 300,
    quantity: 1
  };
}

function getSignPrice(sign) {
  const area = (Math.max(0, Number(sign.width) || 0) * Math.max(0, Number(sign.height) || 0)) / 1000000;
  return Math.round((35 + area * 45) * 100) / 100;
}

function money(value) {
  return new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" }).format(Number(value) || 0);
}

function SignPreview({ sign }) {
  const width = Math.max(50, Number(sign.width) || 50);
  const height = Math.max(50, Number(sign.height) || 50);
  return (
    <div className="scouts-sign-stage">
      <div className="scouts-sign-preview" style={{ aspectRatio: `${width} / ${height}` }}>
        <strong>{sign.groupName || "Your Group Name"}</strong>
      </div>
      <span>{width} x {height}mm</span>
    </div>
  );
}

export function ScoutsCustomerPortal() {
  const [unlocked, setUnlocked] = useState(false);
  const [checking, setChecking] = useState(true);
  const [username, setUsername] = useState("Scouts");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState("");
  const [confirmation, setConfirmation] = useState(null);
  const [details, setDetails] = useState({ fullName: "", mobile: "", email: "", groupName: "", address: "" });
  const [collection, setCollection] = useState(false);
  const [signs, setSigns] = useState([makeSign("")]);

  useEffect(() => {
    fetch("/api/scouts/customer/session")
      .then((response) => {
        if (!response.ok) throw new Error("locked");
        setUnlocked(true);
      })
      .catch(() => setUnlocked(false))
      .finally(() => setChecking(false));
  }, []);

  const signsTotal = useMemo(() => signs.reduce((sum, sign) => sum + getSignPrice(sign) * Math.max(1, Number(sign.quantity) || 1), 0), [signs]);
  const deliveryFee = collection ? 0 : DELIVERY_FEE;
  const total = signsTotal + deliveryFee;

  function updateDetails(key, value) {
    setDetails((current) => {
      const next = { ...current, [key]: value };
      if (key === "groupName") {
        setSigns((currentSigns) => currentSigns.map((sign) => (
          !sign.groupName || sign.groupName === current.groupName ? { ...sign, groupName: value } : sign
        )));
      }
      return next;
    });
  }

  function updateSign(id, key, value) {
    setSigns((current) => current.map((sign) => sign.id === id ? { ...sign, [key]: value } : sign));
  }

  async function unlock(event) {
    event.preventDefault();
    setLoginError("");
    const response = await fetch("/api/scouts/customer/unlock", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username, password })
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      setLoginError(payload.error || "Could not log in.");
      return;
    }
    setUnlocked(true);
    setPassword("");
  }

  async function submitOrder(event) {
    event.preventDefault();
    setError("");
    setSubmitting(true);
    try {
      const response = await fetch("/api/scouts/customer/orders", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...details, collection, signs })
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(payload.error || "Could not place the order.");
      setConfirmation(payload);
    } catch (submitError) {
      setError(submitError.message || "Could not place the order.");
    } finally {
      setSubmitting(false);
    }
  }

  if (checking) return <div className="scouts-shell"><div className="scouts-loading">Loading Scouts portal...</div></div>;

  if (!unlocked) {
    return (
      <div className="scouts-shell scouts-login-shell">
        <section className="scouts-login-card">
          <div className="scouts-mark">SCOUTS</div>
          <p>Group sign ordering portal</p>
          <h1>Customer login</h1>
          <form onSubmit={unlock}>
            <label>Username<input value={username} onChange={(event) => setUsername(event.target.value)} autoComplete="username" /></label>
            <label>Password<input type="password" value={password} onChange={(event) => setPassword(event.target.value)} autoComplete="current-password" autoFocus /></label>
            {loginError ? <div className="scouts-alert error">{loginError}</div> : null}
            <button type="submit">Log in</button>
          </form>
        </section>
      </div>
    );
  }

  if (confirmation) {
    return (
      <div className="scouts-shell scouts-confirmation-shell">
        <section className="scouts-confirmation-card">
          <div className="scouts-confirmation-icon">✓</div>
          <span>Order received</span>
          <h1>Thank you, {confirmation.order.fullName}</h1>
          <p>Your reference is <strong>{confirmation.order.id}</strong>. A copy has been emailed to {confirmation.order.email}.</p>
          <p className="scouts-payment-message">A payment link will be sent shortly.</p>
          {confirmation.emailWarning ? <div className="scouts-alert warning">{confirmation.emailWarning}</div> : null}
        </section>
      </div>
    );
  }

  return (
    <div className="scouts-shell">
      <header className="scouts-header">
        <div><div className="scouts-mark">SCOUTS</div><span>Group sign ordering</span></div>
        <div className="scouts-header-note">Prototype pricing</div>
      </header>
      <main className="scouts-main">
        <form onSubmit={submitOrder}>
          <section className="scouts-section">
            <div className="scouts-section-head"><span>1</span><div><h2>Your details</h2><p>Tell us who the signs are for.</p></div></div>
            <div className="scouts-details-grid">
              <label>Full name<input required value={details.fullName} onChange={(event) => updateDetails("fullName", event.target.value)} /></label>
              <label>Mobile number<input required type="tel" value={details.mobile} onChange={(event) => updateDetails("mobile", event.target.value)} /></label>
              <label>Email address<input required type="email" value={details.email} onChange={(event) => updateDetails("email", event.target.value)} /></label>
              <label>Group name<input required value={details.groupName} onChange={(event) => updateDetails("groupName", event.target.value)} /></label>
              <label className="wide">Delivery address<textarea required rows="3" value={details.address} onChange={(event) => updateDetails("address", event.target.value)} /></label>
            </div>
          </section>

          <section className="scouts-section">
            <div className="scouts-section-head"><span>2</span><div><h2>Your signs</h2><p>Enter finished sizes in millimetres.</p></div></div>
            <div className="scouts-sign-list">
              {signs.map((sign, index) => (
                <article className="scouts-sign-card" key={sign.id}>
                  <div className="scouts-sign-card-head"><h3>Sign {index + 1}</h3>{signs.length > 1 ? <button type="button" className="scouts-remove" onClick={() => setSigns((current) => current.filter((entry) => entry.id !== sign.id))}>Remove</button> : null}</div>
                  <div className="scouts-sign-layout">
                    <div className="scouts-sign-fields">
                      <label>Group name on sign<input required value={sign.groupName} onChange={(event) => updateSign(sign.id, "groupName", event.target.value)} /></label>
                      <div className="scouts-size-fields">
                        <label>Width (mm)<input required type="number" min="50" max="5000" value={sign.width} onChange={(event) => updateSign(sign.id, "width", event.target.value)} /></label>
                        <label>Height (mm)<input required type="number" min="50" max="5000" value={sign.height} onChange={(event) => updateSign(sign.id, "height", event.target.value)} /></label>
                        <label>Quantity<input required type="number" min="1" max="100" value={sign.quantity} onChange={(event) => updateSign(sign.id, "quantity", event.target.value)} /></label>
                      </div>
                      <div className="scouts-line-price"><span>Sign price</span><strong>{money(getSignPrice(sign) * Math.max(1, Number(sign.quantity) || 1))}</strong></div>
                    </div>
                    <SignPreview sign={sign} />
                  </div>
                </article>
              ))}
            </div>
            <button type="button" className="scouts-add-sign" onClick={() => setSigns((current) => [...current, makeSign(details.groupName)])}>+ Add another sign</button>
          </section>

          <section className="scouts-section scouts-checkout">
            <div className="scouts-section-head"><span>3</span><div><h2>Delivery and total</h2><p>Choose delivery or collection.</p></div></div>
            <div className="scouts-fulfilment">
              <button type="button" className={!collection ? "active" : ""} onClick={() => setCollection(false)}>Delivery</button>
              <button type="button" className={collection ? "active" : ""} onClick={() => setCollection(true)}>Collect from store</button>
            </div>
            <div className="scouts-total-box">
              <div><span>Signs</span><strong>{money(signsTotal)}</strong></div>
              <div><span>{collection ? "Collection" : "Delivery"}</span><strong>{collection ? "Free" : money(deliveryFee)}</strong></div>
              <div className="total"><span>Total</span><strong>{money(total)}</strong></div>
              <small>Prototype prices shown for demonstration. Final pricing will be added from the approved Scouts price list.</small>
            </div>
            {error ? <div className="scouts-alert error">{error}</div> : null}
            <button className="scouts-order-button" type="submit" disabled={submitting}>{submitting ? "Sending order..." : "ORDER SIGNS"}</button>
          </section>
        </form>
      </main>
    </div>
  );
}

export function ScoutsAdminOrders() {
  const [orders, setOrders] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    fetch("/api/scouts/admin/orders")
      .then(async (response) => {
        const payload = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(payload.error || "Could not load orders.");
        setOrders(Array.isArray(payload.orders) ? payload.orders : []);
      })
      .catch((loadError) => setError(loadError.message))
      .finally(() => setLoading(false));
  }, []);

  return (
    <div className="scouts-admin-shell">
      <header className="scouts-admin-header"><div><a href="/">← Portal</a><div className="scouts-mark">SCOUTS</div><h1>Sign orders</h1></div><span>{orders.length} order{orders.length === 1 ? "" : "s"}</span></header>
      <main className="scouts-admin-main">
        {loading ? <div className="scouts-loading">Loading orders...</div> : null}
        {error ? <div className="scouts-alert error">{error}</div> : null}
        {!loading && !orders.length ? <div className="scouts-admin-empty">No Scouts orders have been placed yet.</div> : null}
        <div className="scouts-admin-orders">
          {orders.map((order) => (
            <article className="scouts-admin-order" key={order.id}>
              <div className="scouts-admin-order-head"><div><span>{order.id}</span><h2>{order.groupName}</h2><p>{order.fullName} · {order.email} · {order.mobile}</p></div><div><strong>{money(order.total)}</strong><span>{order.status}</span></div></div>
              <div className="scouts-admin-meta"><span>{order.collection ? "Collect from store" : `Deliver to: ${order.address}`}</span><span>{new Date(order.submittedAt).toLocaleString("en-GB")}</span><span>{order.emailSent ? "Email sent" : "Email needs sending"}</span></div>
              <div className="scouts-admin-signs">
                {order.signs.map((sign, index) => <div key={sign.id}><SignPreview sign={sign} /><p><strong>Sign {index + 1}</strong> · Qty {sign.quantity} · {money(sign.lineTotal)}</p></div>)}
              </div>
            </article>
          ))}
        </div>
      </main>
    </div>
  );
}
