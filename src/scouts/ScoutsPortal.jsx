import { useEffect, useMemo, useState } from "react";
import "./scouts.css";

const DELIVERY_FEE = 12.5;
const SIGN_PRODUCTS = [
  { id: "landscape-600x400", title: "Landscape Sign", width: 600, height: 400, price: 25, code: "GE-46STACK-SIGN-C-P" },
  { id: "landscape-800x600", title: "Landscape Sign", width: 800, height: 600, price: 30, code: "GE-86STACK-SIGN-C-P" },
  { id: "landscape-1200x800", title: "Landscape Sign", width: 1200, height: 800, price: 50, code: "GE-128STACK-SIGN-C-P" },
  { id: "portrait-600x800", title: "Portrait Sign", width: 600, height: 800, price: 30, code: "GE-68STACK-SIGN-C-P" }
];
const CUSTOM_PRODUCT_ID = "custom";

function getSignProduct(sign) {
  return SIGN_PRODUCTS.find((product) => product.id === sign.productId) || SIGN_PRODUCTS[0];
}

function getSignDimensions(sign) {
  if (sign.productId === CUSTOM_PRODUCT_ID) {
    return {
      width: Math.min(5000, Math.max(50, Number(sign.width) || 50)),
      height: Math.min(5000, Math.max(50, Number(sign.height) || 50))
    };
  }
  const product = getSignProduct(sign);
  return { width: product.width, height: product.height };
}

function makeSign(groupName = "") {
  const product = SIGN_PRODUCTS[0];
  return {
    id: typeof crypto !== "undefined" && crypto.randomUUID ? crypto.randomUUID() : `sign-${Date.now()}-${Math.random()}`,
    groupName,
    productId: product.id,
    width: product.width,
    height: product.height,
    quantity: 1
  };
}

function getSignPrice(sign) {
  if (sign.productId !== CUSTOM_PRODUCT_ID) return getSignProduct(sign).price;
  const { width, height } = getSignDimensions(sign);
  return Math.round(Math.max(25, 10 + (width * height / 1000000) * (125 / 3)) * 100) / 100;
}

function money(value) {
  return new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" }).format(Number(value) || 0);
}

function SignPreview({ sign }) {
  const product = getSignProduct(sign);
  const { width, height } = getSignDimensions(sign);
  const portrait = height > width;
  const sections = [["Beavers", "6-8"], ["Cubs", "8-10½"], ["Scouts", "10½-14"], ["Explorers", "14-18"], ["Network", "18-25"]];
  return (
    <div className="scouts-sign-stage">
      <div className={`scouts-sign-preview ${portrait ? "is-portrait" : "is-landscape"}`} style={{ aspectRatio: `${width} / ${height}` }}>
        <div className="scouts-sign-identity">
          <img src="/scouts/scouts-logo-stack-white.png" alt="Scouts" />
          <strong>{sign.groupName || "Your Group Name"}</strong>
        </div>
        <div className="scouts-sign-sections">
          {sections.map(([name, ages]) => <span key={name}><b>{name}</b><small>{ages}</small></span>)}
        </div>
        <div className="scouts-sign-footer"><b>scouts.org.uk/join</b><b>0345 300 1818</b></div>
      </div>
      <span>{sign.productId === CUSTOM_PRODUCT_ID ? "Custom Sign" : product.title} · {width} x {height}mm</span>
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
    setSigns((current) => current.map((sign) => {
      if (sign.id !== id) return sign;
      if (key !== "productId") return { ...sign, [key]: value };
      if (value === CUSTOM_PRODUCT_ID) return { ...sign, productId: CUSTOM_PRODUCT_ID };
      const product = SIGN_PRODUCTS.find((entry) => entry.id === value) || SIGN_PRODUCTS[0];
      return { ...sign, productId: product.id, width: product.width, height: product.height };
    }));
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
          <img className="scouts-brand-logo" src="/scouts/scouts-logo-stack.png" alt="Scouts" />
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
        <div><img className="scouts-brand-logo" src="/scouts/scouts-logo-linear.png" alt="Scouts" /><span>Group sign ordering</span></div>
        <div className="scouts-header-note">Meeting place signs</div>
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
            <div className="scouts-section-head"><span>2</span><div><h2>Your signs</h2><p>Choose a finished sign size and see your artwork update instantly.</p></div></div>
            <div className="scouts-product-spec"><span><b>Material</b> 3mm Dibond</span><span><b>Finish</b> Full-colour outdoor print</span><span><b>Lead time</b> Dispatch in 5 working days</span><span><b>Prices</b> Include VAT at 20%</span></div>
            <div className="scouts-sign-list">
              {signs.map((sign, index) => {
                const product = getSignProduct(sign);
                return <article className="scouts-sign-card" key={sign.id}>
                  <div className="scouts-sign-card-head"><h3>Sign {index + 1}</h3>{signs.length > 1 ? <button type="button" className="scouts-remove" onClick={() => setSigns((current) => current.filter((entry) => entry.id !== sign.id))}>Remove</button> : null}</div>
                  <div className="scouts-sign-layout">
                    <div className="scouts-sign-fields">
                      <label>Group name on sign<input required value={sign.groupName} onChange={(event) => updateSign(sign.id, "groupName", event.target.value)} /></label>
                      <div className="scouts-size-fields">
                        <label className="scouts-product-select">Sign size<select value={sign.productId} onChange={(event) => updateSign(sign.id, "productId", event.target.value)}>{SIGN_PRODUCTS.map((entry) => <option key={entry.id} value={entry.id}>{entry.title} · {entry.width} x {entry.height}mm · {money(entry.price)}</option>)}<option value={CUSTOM_PRODUCT_ID}>Custom size</option></select></label>
                        <label>Quantity<input required type="number" min="1" max="100" value={sign.quantity} onChange={(event) => updateSign(sign.id, "quantity", event.target.value)} /></label>
                      </div>
                      {sign.productId === CUSTOM_PRODUCT_ID ? <div className="scouts-custom-size-fields"><label>Width (mm)<input required type="number" min="50" max="5000" value={sign.width} onChange={(event) => updateSign(sign.id, "width", event.target.value)} /></label><label>Height (mm)<input required type="number" min="50" max="5000" value={sign.height} onChange={(event) => updateSign(sign.id, "height", event.target.value)} /></label></div> : null}
                      <div className="scouts-product-detail"><span>{sign.productId === CUSTOM_PRODUCT_ID ? "CUSTOM" : product.code}</span><span>Printed Materials</span></div>
                      <div className="scouts-line-price"><span>{sign.productId === CUSTOM_PRODUCT_ID ? "Estimated " : ""}{money(getSignPrice(sign))} each · VAT included</span><strong>{money(getSignPrice(sign) * Math.max(1, Number(sign.quantity) || 1))}</strong></div>
                    </div>
                    <SignPreview sign={sign} />
                  </div>
                </article>;
              })}
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
              <small>Sign prices include VAT at 20%. Delivery is shown separately and collection is free.</small>
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
      <header className="scouts-admin-header"><div><a href="/">← Portal</a><img className="scouts-brand-logo" src="/scouts/scouts-logo-linear.png" alt="Scouts" /><h1>Sign orders</h1></div><span>{orders.length} order{orders.length === 1 ? "" : "s"}</span></header>
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
                {order.signs.map((sign, index) => <div key={sign.id}><SignPreview sign={sign} /><p><strong>Sign {index + 1}</strong> · {sign.code} · Qty {sign.quantity} · {money(sign.lineTotal)} inc VAT</p></div>)}
              </div>
            </article>
          ))}
        </div>
      </main>
    </div>
  );
}
