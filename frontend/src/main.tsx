import React from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";
import WordToExcel from "../WordToExcel";

const container = document.getElementById("root");
if (!container) throw new Error("Root container not found");
const root = createRoot(container);
root.render(
  <React.StrictMode>
    <WordToExcel />
  </React.StrictMode>
);


