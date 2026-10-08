import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import { installPortalFeedback } from "./portal-ui";
import "./index.css";
import "./rebuild.css";
import "./portal-polish.css";

document.body.classList.add("portal-rebuild");
installPortalFeedback();

const root = document.getElementById("root");
if (root) {
  root.innerHTML = "";
}

ReactDOM.createRoot(root).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
