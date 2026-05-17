"use client";

import { useState, useCallback, useRef } from "react";
import { useToast } from "@/app/components/ToastProvider";
import { exportToExcel } from "@/lib/exportExcel";
import {
  canUseFileSystemAccess, pickExcelFile, getSheetNames,
  buildAppendedBuffer, saveToHandle, downloadBuffer,
} from "@/lib/appendExcel";

// ── Types ─────────────────────────────────────────────────────────
type ParsedReport = Record<string, number | string>;

interface ApiResponse {
  data: ParsedReport;
  csv: string;
  error?: string;
}

// ── Section definitions ───────────────────────────────────────────
const SECTIONS: Record<string, { label: string; color: string; keys: (k: string) => boolean }> = {
  all: { label: "All Fields", color: "#f5a623", keys: () => true },
  crushing: {
    label: "A. Crushing",
    color: "#f5a623",
    keys: (k) =>
      ["Maize Crushing", "SO2", "CSL", "Effluent", "Dryer", "LG-", "Slurry", "HMCS", "MD",
       "Total Output", "Recovery", "Average Recovery", "Germ", "Gluten", "Husk"].some(t => k.includes(t)),
  },
  boiler: {
    label: "B. Boiler",
    color: "#3b8bff",
    keys: (k) =>
      ["Boiler", "Steam", "Turbine", "Exhaust", "Bearing", "MSEB", "Gas Engine", "Power"].some(t => k.includes(t)),
  },
  fuel: {
    label: "C. Fuel",
    color: "#ff6b35",
    keys: (k) =>
      ["Coal", "Bagasse", "Firewood", "Wood Dust", "Gas Fuel"].some(t => k.includes(t)),
  },
  utilities: {
    label: "D. Utilities",
    color: "#00d4aa",
    keys: (k) =>
      ["Water Used", "Waste Water", "Level Tank", "Big Tank", "Hermes Tank", "Process Water Tank", "TDS", "SiO2", "RO TDS"].some(t => k.includes(t)),
  },
  maize: {
    label: "E. Maize",
    color: "#9b6dff",
    keys: (k) => ["Maize From", "Maize Filled", "OB Maize", "Closing Silo", "Total Maize", "Filling Station"].some(t => k.includes(t)),
  },
  vats: {
    label: "F. Vats",
    color: "#22c55e",
    keys: (k) => k.startsWith("Vat "),
  },
};

const KEY_HIGHLIGHTS = ["Maize Crushing (MT)", "Total Output (MT)", "Recovery (%)", "Total Power (KWH)", "Total Dryer Production (MT)"];

// ── Demo report text ──────────────────────────────────────────────
const DEMO_TEXT = `Daily Production Report.
                11/04/2026

---
A. Crushing & Production

Maize crushing = 610.000mt


Stepped Maize So2 Summary: So2 water usage:666m3/ 610 MT GRIND =1.11%

:So2 PPM of Steeped Maize -Avg 7826 PPM


Total CSL Generated: 290m³
Total Effluent = 1340 m3 /2.19m3  per MT of Grind.
CSL to ETP: 155m³
CSL PLANT:135 m³
HCSL to Gluten: 4.81 m³
HCSL to Fiber: 27.55 m³

Dryer Production:

Dryer 3:00.00  MT

Dryer 4: 00.00 MT

Dryer 5: 00.00 MT

Dryer 6: 330.230 MT

Total Dryer Production:330.230 MT

Slurry to LGP=75.396 MT
LG-1=9.147 MT
LG-2 =18.295 MT
LG-3 = 21.954 MT
HMCS =6.000 MT
MD =20.000 MT

Total Output: 405.576 MT

Recovery %:66.48%

Average: 65.88%


Byproducts:

Germ: - 34.120 MT-(5.68%)

Gluten: 27.000 MT(4.42%,Hcsl 4.42m3)

Dry Husk:70.200 MT (11.50%,Hcsl 27.55 m3)

Wet Husk: 0.00 MT

B. Boiler & Turbine

Boiler Load: - 27.30 MT

Steam Generated:692 .00MT

Turbine Load:-  1400 KWH

Inlet Steam: - 25.30TPH

Outlet Steam: 21.50TPH

1st Exhaust Pressure: - 10.3kg/cm²

2nd Exhaust Pressure:- - - 11kg/cm2

Bearing Temp: - 79°C

Power Used:

MSEB: 47100 KWH

Gas Engine: 33200 KWH

Steam Turbine:31550 KWH

Total Power: 111850 KWH (183 units/ton of crushing)

---

C. Fuel Consumption

Coal:30.00MT(Ratio : 5.40)

Loose Bagasse: 241.270 MT  (Ratio 2.20)

Firewood: 0.00 MT

Wood Dust: 0.00 MT

Gas: 0.00 MT

D. Utilities

Water Used: 1910 m3
Waste Water:  1.00 ft

Tank Levels:

Level Tank: 234

Big Tank: 3.0ft +full

Hermes Tank: full

Process Water Tank: Full


Water Quality:

B/F TDS:-14 ppm

M/B SiO2:-  0.03 ppm

RO TDS: -- 10

B/D TDS:- 182 ppm

MIDC TDS:  -115 ppm

---

E. Maize Filling

From Silo 1:0000 bags

From Silo 2: 6200 bags.

Maize Filled in Vats: 11670bags

OB Maize in Silo 1: 58998 bags

Closing in Silo 1: 58998+000=58998-0000=58998 bags

OB Maize  in Sio no.2 : 32140 bags

Closing in Silo2:36630+1710=38340-6200=32140 bags

Total Maize Filled: 6200 bags

Maize Filling Station: 0000 bags

F. Water Damage & Small Size Maize %

Vat No   WD %     Small size%
13=16.30%       16.00%
14=15.00%       15.00%
15=14.00%.      14.80%
16=14.16%.      13.00%
17=17.66%.      16.00%
18=10.26%       14.00%
19=13.40%       15.00%
20=10.00%       15.00%
21=13.00%       13.00%
22=14.83%.      15.00%
23=14.40%.      14.00%`;

// ── Helpers ───────────────────────────────────────────────────────
function formatValue(v: number | string): string {
  if (typeof v === "number") {
    if (Number.isInteger(v)) return v.toLocaleString();
    return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 3 });
  }
  return String(v);
}

function isZero(v: number | string): boolean {
  if (typeof v === "number") return v === 0;
  return v === "";
}

function isString(v: number | string): boolean {
  return typeof v === "string" && v !== "";
}

function downloadCSV(csv: string, filename: string) {
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function downloadJSON(data: ParsedReport, filename: string) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

// ── Component ─────────────────────────────────────────────────────
export default function Home() {
  const toast = useToast();
  const [text, setText] = useState("");
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<ApiResponse | null>(null);
  const [activeSection, setActiveSection] = useState("all");
  const [search, setSearch] = useState("");
  const [showCSV, setShowCSV] = useState(false);
  const [xlsxLoading, setXlsxLoading] = useState(false);

  // ── Append-to-Excel state ──────────────────────────────────────────
  const [appendBusy, setAppendBusy]     = useState(false);
  const [sheetNames, setSheetNames]     = useState<string[]>([]);
  const [showManualModal, setShowManualModal] = useState(false);
  const [manualJson, setManualJson] = useState("");
  const fileHandleRef = useRef<FileSystemFileHandle | null>(null);
  const fileBufferRef = useRef<ArrayBuffer | null>(null);
  const pendingDataRef = useRef<ParsedReport | null>(null);

  const handleParse = useCallback(async () => {
    if (!text.trim()) {
      toast.warning("Nothing to parse", "Paste your WhatsApp report into the text area first");
      return;
    }
    setLoading(true);
    setResult(null);
    const loadId = toast.loading("Parsing report…", "Extracting production KPIs from text");
    try {
      const res = await fetch("/api/parse", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text }),
      });
      const json: ApiResponse = await res.json();
      toast.dismiss(loadId);
      if (!res.ok || json.error) {
        toast.error("Parse failed", json.error || "Unexpected server error");
      } else {
        setResult(json);
        setActiveSection("all");
        setSearch("");
        const total   = Object.keys(json.data).length;
        const matched = Object.values(json.data).filter(v => v !== 0 && v !== "").length;
        const d       = (json.data["Date"] as string) || "";
        const vats    = Object.keys(json.data).filter(k => k.startsWith("Vat ")).length / 2;
        toast.success(
          `${matched} / ${total} fields extracted`,
          `${d ? `Date: ${d}  ·  ` : ""}${vats > 0 ? `${vats} vats parsed` : ""}`
        );
      }
    } catch (e) {
      toast.dismiss(loadId);
      toast.error("Network error", String(e));
    } finally {
      setLoading(false);
    }
  }, [text, toast]);

  const loadDemo = () => {
    setText(DEMO_TEXT);
    setResult(null);
    toast.info("Demo report loaded", "Click Parse Report to extract the KPIs");
  };

  const clearAll = () => {
    setText("");
    setResult(null);
    toast.info("Cleared", "Paste a new report whenever you're ready");
  };

  const handleExportExcel = useCallback(async () => {
    if (!result) return;
    
    // Prompt user for filename
    const reportDate = (result.data["Date"] as string) || "report";
    const defaultName = `Production_${reportDate.replace(/\//g, "-")}`;
    const customName = window.prompt("Enter Excel filename (without .xlsx):", defaultName);
    
    if (customName === null) return; // User cancelled
    if (!customName.trim()) {
      toast.warning("Invalid filename", "Please enter a filename");
      return;
    }
    
    setXlsxLoading(true);
    const loadId = toast.loading("Building Excel workbook…", "Generating Microsoft Excel file");
    try {
      const filename = customName.trim().endsWith(".xlsx") 
        ? customName.trim() 
        : `${customName.trim()}.xlsx`;
      await exportToExcel(result.data, filename);
      toast.dismiss(loadId);
      toast.success("Excel downloaded", `Saved as: ${filename}`);
    } catch (e) {
      toast.dismiss(loadId);
      toast.error("Excel export failed", String(e));
    } finally {
      setXlsxLoading(false);
    }
  }, [result, toast]);

  const openManualModal = useCallback(() => {
    if (!result) return;
    setManualJson(JSON.stringify(result.data, null, 2));
    setShowManualModal(true);
  }, [result]);
  // ── Append to Excel helper (defined before callers to avoid TDZ) ─────────
  const doAppend = useCallback(async (sheetName: string, dataToAppend?: ParsedReport) => {
    const data = dataToAppend ?? result?.data;
    if (!data) return;
    setSheetNames([]); setAppendBusy(true);
    const loadId = toast.loading(`Appending to "${sheetName}"…`, "Sorting rows by date & saving");
    try {
      const { buffer, stats } = await buildAppendedBuffer(fileBufferRef.current, data, sheetName);
      if (stats.duplicate) {
        toast.dismiss(loadId);
        toast.warning("Duplicate date", `A row for ${data["Date"]} already exists in this sheet`);
        setAppendBusy(false); return;
      }
      if (fileHandleRef.current) {
        await saveToHandle(fileHandleRef.current, buffer);
        toast.dismiss(loadId);
        toast.success(
          `Row added · ${stats.totalRows} total rows`,
          `Sheet: ${sheetName}${stats.newCols.length ? `  ·  +${stats.newCols.length} new cols` : ""}`
        );
      } else {
        const date = (data["Date"] as string || "report").replace(/\//g, "-");
        downloadBuffer(buffer, `Production_${date}.xlsx`);
        toast.dismiss(loadId);
        toast.success("Excel downloaded", `${stats.totalRows} rows in sheet "${sheetName}"`);
      }
      setText(""); setResult(null); fileHandleRef.current = null; fileBufferRef.current = null;
      pendingDataRef.current = null;
    } catch (e) {
      toast.dismiss(loadId);
      toast.error("Save failed", String(e));
    } finally {
      setAppendBusy(false);
    }
  }, [result, toast]);

  const handleManualSubmit = useCallback(async () => {
    try {
      const parsed = JSON.parse(manualJson) as ParsedReport;
      pendingDataRef.current = parsed;
      setShowManualModal(false);
      setAppendBusy(true);
      if (canUseFileSystemAccess()) {
        const picked = await pickExcelFile();
        if (!picked) { setAppendBusy(false); return; }
        fileHandleRef.current = picked.handle;
        fileBufferRef.current = picked.buffer;
        const names = await getSheetNames(picked.buffer);
        if (names.length === 1) {
          await doAppend(names[0], parsed);
        } else {
          setSheetNames(names);
          setAppendBusy(false);
        }
      } else {
        const { buffer, stats } = await buildAppendedBuffer(null, parsed, "Production Report");
        const date = (parsed["Date"] as string || "report").replace(/\//g, "-");
        downloadBuffer(buffer, `Production_${date}.xlsx`);
        toast.success(`Row appended (downloaded)`, `${stats.totalRows} total rows · ${stats.newCols.length} new columns`);
        setText(""); setResult(null);
        setAppendBusy(false);
      }
    } catch (e) {
      toast.error("Invalid JSON", String(e));
      setAppendBusy(false);
    }
  }, [manualJson, doAppend, toast]);

  // ── Append to Excel handlers ───────────────────────────────────────
  const handleAddToExcel = useCallback(async () => {
    if (!result) return;
    setAppendBusy(true);
    try {
      if (canUseFileSystemAccess()) {
        const picked = await pickExcelFile();
        if (!picked) { setAppendBusy(false); return; }
        fileHandleRef.current = picked.handle;
        fileBufferRef.current = picked.buffer;
        const names = await getSheetNames(picked.buffer);
        if (names.length === 1) {
          await doAppend(names[0]);
        } else {
          setSheetNames(names); // show modal
          setAppendBusy(false);
        }
      } else {
        // Fallback: build new file and download
        const { buffer, stats } = await buildAppendedBuffer(null, result.data, "Production Report");
        const date = (result.data["Date"] as string || "report").replace(/\//g, "-");
        downloadBuffer(buffer, `Production_${date}.xlsx`);
        toast.success(`Row appended (downloaded)`, `${stats.totalRows} total rows · ${stats.newCols.length} new columns`);
        setText(""); setResult(null);
        setAppendBusy(false);
      }
    } catch (e) {
      toast.error("Append failed", String(e));
      setAppendBusy(false);
    }
  }, [result, toast]);

  // Filter entries
  const allEntries = result ? Object.entries(result.data) : [];
  const vatEntries = allEntries.filter(([k]) => k.startsWith("Vat "));
  const vatNumbers = [...new Set(vatEntries.map(([k]) => parseInt(k.replace(/[^\d]/g, "").slice(0,3), 10)))].sort((a,b)=>a-b);

  const filteredEntries = allEntries
    .filter(([k]) => SECTIONS[activeSection].keys(k))
    .filter(([k, v]) => {
      if (!search) return true;
      return k.toLowerCase().includes(search.toLowerCase()) ||
        String(v).toLowerCase().includes(search.toLowerCase());
    });

  const nonVatFiltered = filteredEntries.filter(([k]) => !k.startsWith("Vat "));
  const vatFiltered = filteredEntries.filter(([k]) => k.startsWith("Vat "));

  const totalFields = allEntries.length;
  const matchedFields = allEntries.filter(([, v]) => !isZero(v)).length;
  const date = result?.data["Date"] as string || "";
  const recovery = result?.data["Recovery (%)"] as number || 0;
  const maizeCrushing = result?.data["Maize Crushing (MT)"] as number || 0;

  return (
    <div className="app-wrapper">
      <div className="bg-grid" />
      <div className="bg-orb bg-orb-1" />
      <div className="bg-orb bg-orb-2" />
      <div className="bg-orb bg-orb-3" />

      <main>
        <div className="container">
          {/* ── HEADER ───────────────────────────────────────────── */}
          <header className="header">
            <div className="header-badge">
              <span className="dot" />
              Maize Crushing Production
            </div>
            <h1 className="header-title">
              WhatsApp Report<br />Intelligence Platform
            </h1>
            <p className="header-subtitle">
              Paste your daily WhatsApp production report and instantly extract
              80+ structured KPIs — ready for Power BI, Excel, and CSV export.
            </p>

            {result && (
              <div className="stats-bar">
                <div className="stat-item">
                  <span className="stat-value">{date || "—"}</span>
                  <span className="stat-label">Report Date</span>
                </div>
                <div className="stat-item">
                  <span className="stat-value">{matchedFields}</span>
                  <span className="stat-label">Fields Matched</span>
                </div>
                <div className="stat-item">
                  <span className="stat-value">{maizeCrushing > 0 ? `${maizeCrushing.toFixed(1)}` : "—"}</span>
                  <span className="stat-label">Crushing MT</span>
                </div>
                <div className="stat-item">
                  <span className="stat-value">{recovery > 0 ? `${recovery.toFixed(2)}%` : "—"}</span>
                  <span className="stat-label">Recovery</span>
                </div>
                <div className="stat-item">
                  <span className="stat-value">{vatNumbers.length}</span>
                  <span className="stat-label">Vats Parsed</span>
                </div>
              </div>
            )}
          </header>

          {/* ── INPUT CARD ───────────────────────────────────────── */}
          <div className="glass-card" style={{ padding: "32px", marginBottom: "24px" }}>
            <div className="input-section">
              <div className="input-label">
                <svg className="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
                </svg>
                Paste WhatsApp Report
              </div>
              <textarea
                id="report-input"
                className="report-textarea"
                value={text}
                onChange={e => setText(e.target.value)}
                placeholder={`Paste your daily report here...\n\nExample:\n  Daily Production Report.\n  11/04/2026\n  Maize crushing = 610.000mt\n  ...`}
                spellCheck={false}
              />
            </div>


            <div className="actions-row">
              <button id="parse-btn" className="btn btn-primary" onClick={handleParse} disabled={loading}>
                {loading
                  ? <><div className="spinner" />&nbsp;Parsing…</>
                  : <>
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                        <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>
                      </svg>
                      Parse Report
                    </>}
              </button>

              <button id="demo-btn" className="demo-chip" onClick={loadDemo}>
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <polygon points="5 3 19 12 5 21 5 3"/>
                </svg>
                Load Demo
              </button>

              {text && (
                <button id="clear-btn" className="btn btn-ghost" onClick={clearAll}>
                  Clear
                </button>
              )}

              <span style={{ marginLeft:"auto", color: "var(--text-muted)", fontSize: "0.8rem" }}>
                {text.length.toLocaleString()} chars
              </span>
            </div>
          </div>

          {/* ── RESULTS ──────────────────────────────────────────── */}
          {result && (
            <div className="glass-card" style={{ padding: "32px" }}>

              {/* ── Add to Excel Sheet CTA banner ── */}
              <div className="add-excel-banner">
                <div>
                  <p className="add-excel-banner-text">✅ Report parsed — ready to save</p>
                  <p className="add-excel-banner-sub">Append this row to your master Excel workbook, sorted by date</p>
                </div>
                <button
                  id="add-excel-btn"
                  className="btn-add-excel"
                  onClick={handleAddToExcel}
                  disabled={appendBusy}
                >
                  {appendBusy ? (
                    <><div className="spinner" style={{borderColor:"rgba(255,255,255,0.3)",borderTopColor:"#fff",width:14,height:14}} />&nbsp;Working…</>
                  ) : (
                    <>
                      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round">
                        <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/>
                        <path d="m12 13 0 4m-2-2 2 2 2-2" strokeWidth="2"/>
                      </svg>
                      Add to Excel Sheet
                    </>
                  )}
                </button>
                <button
                  id="manual-add-btn"
                  className="btn btn-ghost"
                  onClick={openManualModal}
                  disabled={appendBusy}
                  style={{ marginLeft: 12 }}
                >
                  Manual Add
                </button>
              </div>

              {/* Progress bar */}
              <div className="progress-bar">
                <div className="progress-fill" style={{ width: `${(matchedFields / totalFields) * 100}%` }} />
              </div>

              {/* Results header */}
              <div className="results-header">
                <div className="results-title">
                  Extracted Fields
                  <span className="results-count-badge">{totalFields} total</span>
                </div>
                <div style={{ display:"flex", gap:"10px", alignItems:"center", flexWrap:"wrap" }}>
                  <div className="search-wrap">
                    <svg className="search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/>
                    </svg>
                    <input
                      id="search-input"
                      type="text"
                      className="search-input"
                      placeholder="Search fields…"
                      value={search}
                      onChange={e => setSearch(e.target.value)}
                    />
                  </div>
                  <button id="csv-btn" className="btn btn-teal" style={{padding:"9px 18px",fontSize:"0.82rem"}}
                    onClick={() => {
                      if (result.csv) {
                        downloadCSV(result.csv, `production_${date || "report"}.csv`);
                        toast.success("CSV downloaded", "Open in Excel or import to Power BI");
                      }
                    }}>
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                    </svg>
                    CSV
                  </button>
                  <button id="json-btn" className="btn btn-secondary" style={{padding:"9px 18px",fontSize:"0.82rem"}}
                    onClick={() => {
                      downloadJSON(result.data, `production_${date || "report"}.json`);
                      toast.success("JSON downloaded", "Structured production data ready");
                    }}>
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                    </svg>
                    JSON
                  </button>
                  <button
                    id="excel-btn"
                    className="btn"
                    disabled={xlsxLoading}
                    onClick={handleExportExcel}
                    style={{
                      padding: "9px 18px",
                      fontSize: "0.82rem",
                      background: xlsxLoading
                        ? "rgba(34,197,94,0.4)"
                        : "linear-gradient(135deg,#217346,#1a5c38)",
                      color: "#fff",
                      boxShadow: "0 4px 16px rgba(33,115,70,0.35)",
                      border: "1px solid rgba(255,255,255,0.1)",
                    }}
                  >
                    {xlsxLoading ? (
                      <><div className="spinner" style={{borderColor:"rgba(255,255,255,0.3)",borderTopColor:"#fff"}} />&nbsp;Generating…</>
                    ) : (
                      <>
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                          <rect x="3" y="3" width="18" height="18" rx="2"/>
                          <path d="M3 9h18M9 21V9"/>
                          <path d="m7 13 2 2 4-4" strokeWidth="2"/>
                        </svg>
                        Excel
                      </>
                    )}
                  </button>
                  <button className="btn btn-ghost" style={{padding:"9px 14px",fontSize:"0.78rem"}}
                    onClick={() => setShowCSV(v => !v)}>
                    {showCSV ? "Hide" : "Preview CSV"}
                  </button>
                </div>
              </div>

              {/* Section tabs */}
              <div className="section-tabs">
                {Object.entries(SECTIONS).map(([key, s]) => (
                  <button
                    key={key}
                    className={`tab-btn ${activeSection === key ? "active" : ""}`}
                    style={activeSection === key ? { background: s.color, borderColor: s.color } : {}}
                    onClick={() => { setActiveSection(key); setSearch(""); }}
                  >
                    {s.label}
                  </button>
                ))}
              </div>

              {/* CSV preview */}
              {showCSV && (
                <div className="csv-preview">
                  {result.csv}
                </div>
              )}

              {/* Non-vat data */}
              {nonVatFiltered.length > 0 && (
                <>
                  {activeSection !== "vats" && (
                    <div className="section-heading">
                      {SECTIONS[activeSection].label === "F. Vats" ? "" : SECTIONS[activeSection].label} Data Fields
                    </div>
                  )}
                  <div className="data-grid">
                    {nonVatFiltered.map(([key, value], i) => (
                      <div
                        key={key}
                        className={`data-cell ${KEY_HIGHLIGHTS.includes(key) ? "highlighted" : ""}`}
                        style={{ animationDelay: `${Math.min(i * 20, 400)}ms` }}
                      >
                        <div className="cell-key" title={key}>{key}</div>
                        <div className={`cell-value ${isZero(value) ? "zero" : ""} ${isString(value) ? "string-val" : ""}`}>
                          {formatValue(value) || "—"}
                        </div>
                      </div>
                    ))}
                  </div>
                </>
              )}

              {/* Vat table */}
              {vatFiltered.length > 0 && (
                <>
                  <div className="section-heading" style={{ color: "#22c55e80" }}>F. Vat Water Damage & Small Size</div>
                  <div className="vat-table-wrap">
                    <table className="vat-table">
                      <thead>
                        <tr>
                          <th>Vat #</th>
                          <th>WD (%)</th>
                          <th>Small Size (%)</th>
                          <th>WD Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {vatNumbers
                          .filter(n => {
                            if (search) {
                              return `Vat ${n}`.toLowerCase().includes(search.toLowerCase());
                            }
                            return result.data[`Vat ${n} WD (%)`] !== undefined;
                          })
                          .map(n => {
                            const wd = result.data[`Vat ${n} WD (%)`] as number;
                            const ss = result.data[`Vat ${n} SS (%)`] as number;
                            const status = wd < 12 ? { label: "Good", color: "#22c55e" }
                              : wd < 16 ? { label: "Normal", color: "#f5a623" }
                              : { label: "High", color: "#ef4444" };
                            return (
                              <tr key={n}>
                                <td style={{ color: "var(--accent-gold)", fontWeight: 600 }}>{n}</td>
                                <td>{wd?.toFixed(2)}%</td>
                                <td>{ss?.toFixed(2)}%</td>
                                <td>
                                  <span style={{
                                    background: `${status.color}22`,
                                    color: status.color,
                                    border: `1px solid ${status.color}44`,
                                    borderRadius: "100px",
                                    padding: "2px 10px",
                                    fontSize: "0.72rem",
                                    fontWeight: 700,
                                    letterSpacing: "0.05em",
                                  }}>
                                    {status.label}
                                  </span>
                                </td>
                              </tr>
                            );
                          })}
                      </tbody>
                    </table>
                  </div>
                </>
              )}

              {filteredEntries.length === 0 && (
                <div style={{ textAlign:"center", padding:"40px 0", color:"var(--text-muted)", fontSize:"0.9rem" }}>
                  No fields match your search.
                </div>
              )}
            </div>
          )}
        </div>
      </main>

      <footer className="footer">
        <p>
          MazeCrush Production Intelligence &nbsp;·&nbsp;
          Parses <strong style={{color:"var(--text-secondary)"}}>80+</strong> production KPIs from WhatsApp reports &nbsp;·&nbsp;
          <a href="https://vercel.com" target="_blank" rel="noreferrer">Deployed on Vercel</a>
        </p>
      </footer>

      {/* Sheet selection modal (when workbook has multiple sheets) */}
      {sheetNames.length > 0 && (
        <div className="modal-backdrop">
          <div className="modal-card" style={{ width: 500, maxWidth: "95%", background: "linear-gradient(135deg, rgba(15,23,42,0.95), rgba(20,30,50,0.95))", border: "1px solid rgba(148,163,184,0.1)" }}>
            {/* Header */}
            <div style={{ display: "flex", alignItems: "flex-start", gap: 16, marginBottom: 24 }}>
              <div style={{ width: 40, height: 40, borderRadius: 8, background: "rgba(0,120,212,0.15)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" style={{ color: "#0078D4" }}>
                  <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/><path d="m12 13 0 4m-2-2 2 2 2-2" strokeWidth="2"/>
                </svg>
              </div>
              <div style={{ flex: 1 }}>
                <h3 style={{ margin: 0, fontSize: "1.1rem", fontWeight: 600, color: "#fff" }}>Choose Sheet</h3>
                <p style={{ margin: "4px 0 0 0", color: "var(--text-muted)", fontSize: "0.85rem" }}>Select which sheet to append the data to</p>
              </div>
              <button onClick={() => { setSheetNames([]); setAppendBusy(false); pendingDataRef.current = null; }} style={{ background: "none", border: "none", color: "var(--text-muted)", fontSize: "1.5rem", cursor: "pointer", padding: 0, width: 32, height: 32, display: "flex", alignItems: "center", justifyContent: "center" }}>
                ✕
              </button>
            </div>

            {/* Sheets list */}
            <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 24 }}>
              {sheetNames.map(n => (
                <button 
                  key={n} 
                  onClick={() => doAppend(n, pendingDataRef.current ?? undefined)}
                  style={{
                    padding: "14px 16px",
                    borderRadius: 8,
                    border: "1px solid rgba(148,163,184,0.2)",
                    background: "rgba(2,8,23,0.4)",
                    color: "#e2e8f0",
                    cursor: "pointer",
                    fontSize: "0.95rem",
                    fontWeight: 500,
                    transition: "all 0.2s",
                    textAlign: "left",
                  }}
                  onMouseEnter={(e) => {
                    (e.target as HTMLButtonElement).style.background = "rgba(0,120,212,0.15)";
                    (e.target as HTMLButtonElement).style.borderColor = "rgba(0,120,212,0.3)";
                  }}
                  onMouseLeave={(e) => {
                    (e.target as HTMLButtonElement).style.background = "rgba(2,8,23,0.4)";
                    (e.target as HTMLButtonElement).style.borderColor = "rgba(148,163,184,0.2)";
                  }}
                >
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ marginRight: 10, display: "inline", color: "#0078D4" }}>
                    <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/>
                  </svg>
                  {n}
                </button>
              ))}
            </div>

            {/* Actions */}
            <div style={{ display: "flex", gap: 12, justifyContent: "flex-end", paddingTop: 16, borderTop: "1px solid rgba(148,163,184,0.1)" }}>
              <button 
                className="btn btn-ghost" 
                onClick={() => { setSheetNames([]); setAppendBusy(false); pendingDataRef.current = null; }}
                style={{ padding: "10px 20px", fontSize: "0.9rem" }}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Manual-add JSON modal */}
      {showManualModal && (
        <div className="modal-backdrop">
          <div className="modal-card" style={{ width: 760, maxWidth: "95%", background: "linear-gradient(135deg, rgba(15,23,42,0.95), rgba(20,30,50,0.95))", border: "1px solid rgba(148,163,184,0.1)" }}>
            {/* Header */}
            <div style={{ display: "flex", alignItems: "flex-start", gap: 16, marginBottom: 24 }}>
              <div style={{ width: 40, height: 40, borderRadius: 8, background: "rgba(59,139,255,0.1)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ color: "#3b8bff" }}>
                  <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
                </svg>
              </div>
              <div style={{ flex: 1 }}>
                <h3 style={{ margin: 0, fontSize: "1.2rem", fontWeight: 600, color: "#fff" }}>Edit Production Data</h3>
                <p style={{ margin: "4px 0 0 0", color: "var(--text-muted)", fontSize: "0.9rem" }}>Customize the extracted fields before appending to your workbook</p>
              </div>
              <button onClick={() => setShowManualModal(false)} style={{ background: "none", border: "none", color: "var(--text-muted)", fontSize: "1.5rem", cursor: "pointer", padding: 0, width: 32, height: 32, display: "flex", alignItems: "center", justifyContent: "center" }}>
                ✕
              </button>
            </div>

            {/* Editor Section */}
            <div style={{ marginBottom: 24 }}>
              <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 600, color: "var(--text-secondary)", marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.5px" }}>JSON Content</label>
              <div style={{ position: "relative", borderRadius: 8, overflow: "hidden", border: "1px solid rgba(148,163,184,0.2)", background: "rgba(2,8,23,0.6)" }}>
                <textarea 
                  value={manualJson} 
                  onChange={e => setManualJson(e.target.value)} 
                  style={{ 
                    width: "100%", 
                    height: 300, 
                    fontFamily: "ui-monospace, 'Courier New', monospace", 
                    fontSize: 13,
                    padding: 16,
                    background: "transparent",
                    color: "#e2e8f0",
                    border: "none",
                    outline: "none",
                    resize: "none",
                    lineHeight: "1.5",
                  }} 
                  spellCheck="false"
                />
                <div style={{ position: "absolute", bottom: 12, right: 12, fontSize: "0.75rem", color: "rgba(148,163,184,0.6)" }}>
                  {manualJson.split('\n').length} lines
                </div>
              </div>
              <p style={{ margin: "8px 0 0 0", fontSize: "0.8rem", color: "rgba(148,163,184,0.7)" }}>💡 Modify field values directly. Invalid JSON will show an error when you submit.</p>
            </div>

            {/* Actions */}
            <div style={{ display: "flex", gap: 12, justifyContent: "flex-end", paddingTop: 16, borderTop: "1px solid rgba(148,163,184,0.1)" }}>
              <button 
                className="btn btn-ghost" 
                onClick={() => setShowManualModal(false)}
                style={{ padding: "10px 20px", fontSize: "0.9rem" }}
              >
                Cancel
              </button>
              <button 
                className="btn btn-primary" 
                onClick={handleManualSubmit} 
                disabled={appendBusy}
                style={{ padding: "10px 24px", fontSize: "0.9rem", fontWeight: 600 }}
              >
                {appendBusy ? (
                  <><div className="spinner" style={{borderColor:"rgba(255,255,255,0.3)",borderTopColor:"#fff",width:14,height:14,marginRight:8}} />Processing…</>
                ) : (
                  <>
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" style={{ marginRight: 8, display: "inline" }}>
                      <rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/><path d="m12 13 0 4m-2-2 2 2 2-2" strokeWidth="2"/>
                    </svg>
                    Pick Workbook & Add
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
