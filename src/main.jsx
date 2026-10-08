import React, { Suspense, lazy } from "react";
import ReactDOM from "react-dom/client";
import { SKIN_KEY, resolveSkin, skinLink } from "./portal-skin.mjs";
import "./index.css";
import "./portal-skin.css";

let stored;
try { stored = window.localStorage.getItem(SKIN_KEY); } catch {}
const skin = resolveSkin(window.location.search, stored);
try { window.localStorage.setItem(SKIN_KEY, skin); } catch {}
document.body.classList.toggle("portal-rebuild", skin === "new");
const App = lazy(() => skin === "new" ? import("./modern-entry") : import("./App"));

const root = document.getElementById("root");
if (root) root.innerHTML = "";
ReactDOM.createRoot(root).render(
  <React.StrictMode>
    <Suspense fallback={<div role="status" className="skin-loading">Loading portal…</div>}><App /></Suspense>
    <aside className="portal-skin-switch" aria-label="Portal appearance">
      <span>{skin === "new" ? "New look · live data" : "Classic portal"}</span>
      <a href={skinLink(window.location.href, skin === "new" ? "classic" : "new")}>
        {skin === "new" ? "Switch to Classic" : "Try the new look"}
      </a>
    </aside>
  </React.StrictMode>
);
