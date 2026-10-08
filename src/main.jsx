import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "./index.css";
import "./rebuild.css";

document.body.classList.add("portal-rebuild");

const root = document.getElementById("root");
if (root) {
  root.innerHTML = "";
}

ReactDOM.createRoot(root).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
