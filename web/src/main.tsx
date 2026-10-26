import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import { emit, restore } from "./store";
import "./app.css";

if ((window as any).pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = "./lib/vendor/pdf.worker.min.js";
restore();
AI.init();
createRoot(document.getElementById("root")!).render(<StrictMode><App /></StrictMode>);
emit();
