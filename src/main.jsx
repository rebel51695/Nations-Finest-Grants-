import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "./storageShim";
import "./index.css";
import GrantFlow from "./GrantFlow.jsx";

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <GrantFlow />
  </StrictMode>
);
