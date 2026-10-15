import React, { lazy, Suspense } from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import { isRetiredModulePath } from "./retired-modules.mjs";
import "./index.css";

// Retire early access even for browsers that previously selected the new skin.
document.body.classList.remove("portal-rebuild");
try { window.localStorage.removeItem("sx-portal-skin"); } catch {}
if (isRetiredModulePath(window.location.pathname)) {
  window.history.replaceState(null, "", "/");
}
const GosPage = lazy(() => import("./gos/GosPage"));
const isGos = /^\/gos(?:\/|$)/i.test(window.location.pathname);
const root = document.getElementById("root");
if (root) root.innerHTML = "";
ReactDOM.createRoot(root).render(
  <React.StrictMode>
    {isGos ? <Suspense fallback={<p>Loading GOS Vehicle Liveries…</p>}><GosPage /></Suspense> : <App />}
  </React.StrictMode>
);
