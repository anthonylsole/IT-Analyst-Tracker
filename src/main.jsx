import "./storagePolyfill.js";
import React from "react";
import ReactDOM from "react-dom/client";
import CapacityTracker from "./CapacityTracker.jsx";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <CapacityTracker />
  </React.StrictMode>
);
