import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "./index.css";

// Retire early access even for browsers that previously selected the new skin.
document.body.classList.remove("portal-rebuild");
try { window.localStorage.removeItem("sx-portal-skin"); } catch {}
const root = document.getElementById("root");
if (root) root.innerHTML = "";
ReactDOM.createRoot(root).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
