/**
 * exportExcel.ts — Browser-side Excel export (XLSX/SheetJS)
 * Industry-standard XLSX library for professional Excel generation
 * Modern Microsoft Office 365 Design
 */

import { utils, write } from "xlsx";
import { saveAs } from "file-saver";
import type { ParsedReport } from "./parser";

// ── Modern Microsoft Office 365 Design Tokens ────────────────────────────────
const HEADER_BG  = "0078D4";   // Modern Office Blue
const HEADER_FG  = "FFFFFF";   // White
const ROW_ODD    = "FFFFFF";   // Clean White
const ROW_EVEN   = "F5F5F5";   // Subtle Light Grey

// ── Parse DD/MM/YYYY → Excel Date number ────────────────────────────────────
function parseDateStr(s: string): number | string {
  const patterns = [
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,
    /^(\d{1,2})-(\d{1,2})-(\d{4})$/,
    /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/,
  ];
  for (const re of patterns) {
    const m = s.match(re);
    if (m) {
      const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
      if (!isNaN(d.getTime())) return Math.floor((d.getTime() - new Date(1900, 0, 1).getTime()) / 86400000) + 2;
    }
  }
  return s;
}

// ── String keys ──────────────────────────────────────────────────────────────
const STRING_KEYS = new Set(["Date", "Big Tank", "Hermes Tank", "Process Water Tank"]);

// ── Main export ───────────────────────────────────────────────────────────────
export async function exportToExcel(
  data: ParsedReport,
  filename = "Production_Report.xlsx",
  sheetName = "Production Report",
): Promise<void> {
  const headers = Object.keys(data);
  
  // Prepare data row
  const rowValues = headers.map((h, i) => {
    const v = Object.values(data)[i];
    if (h === "Date" && typeof v === "string" && v) {
      return parseDateStr(v);
    }
    return v;
  });

  // Create worksheet with headers and data
  const ws = utils.aoa_to_sheet([headers, rowValues]);

  // ── Style header row ─────────────────────────────────────────────────────────
  for (let c = 0; c < headers.length; c++) {
    const cellRef = utils.encode_cell({ r: 0, c });
    if (!ws[cellRef]) ws[cellRef] = {};
    // Preserve value and add styling
    ws[cellRef].v = ws[cellRef].v || headers[c];
    ws[cellRef].t = "s";
    ws[cellRef].s = {
      font: { bold: true, color: { rgb: HEADER_FG }, name: "Segoe UI", sz: 11 },
      fill: { patternType: "solid", fgColor: { rgb: HEADER_BG }, theme: undefined, tint: undefined },
      alignment: { horizontal: "center", vertical: "middle", wrapText: true },
      border: {
        left: { style: "thin", color: { rgb: "D0D0D0" } },
        right: { style: "thin", color: { rgb: "D0D0D0" } },
        top: { style: "thin", color: { rgb: "D0D0D0" } },
        bottom: { style: "thin", color: { rgb: "D0D0D0" } },
      },
    };
  }

  // ── Style data row ───────────────────────────────────────────────────────────
  for (let c = 0; c < headers.length; c++) {
    const cellRef = utils.encode_cell({ r: 1, c });
    const h = headers[c];
    const v = rowValues[c];
    const isDateCol = h === "Date";
    const isPercentCol = h.includes("%");
    const isZero = typeof v === "number" && v === 0;

    if (!ws[cellRef]) ws[cellRef] = {};
    ws[cellRef].s = {
      font: {
        name: "Segoe UI",
        sz: 10,
        bold: isDateCol,
        color: {
          rgb: isDateCol ? "0078D4" : isZero ? "A8AAAD" : typeof v === "number" ? "242424" : "606060",
        },
        italic: typeof v === "string" && v !== "",
      },
      fill: { patternType: "solid", fgColor: { rgb: ROW_ODD }, theme: undefined, tint: undefined },
      alignment: { horizontal: "center", vertical: "middle", wrapText: true },
      border: {
        left: { style: "thin", color: { rgb: "D0D0D0" } },
        right: { style: "thin", color: { rgb: "D0D0D0" } },
        top: { style: "thin", color: { rgb: "D0D0D0" } },
        bottom: { style: "thin", color: { rgb: "D0D0D0" } },
      },
      numFmt: isDateCol ? "DD/MM/YYYY" : isPercentCol ? "0.00%" : typeof v === "number" ? "#,##0.000" : "@",
    };
  }

  // ── Set column widths ────────────────────────────────────────────────────────
  ws["!cols"] = headers.map((h, i) => {
    const val = rowValues[i] as any;
    const valLen = String(val instanceof Date ? val.toLocaleString() : val).length;
    return { wch: Math.max(h.length + 3, valLen + 2, 14) };
  });

  // ── Set row heights ─────────────────────────────────────────────────────────
  ws["!rows"] = [
    { hpx: 40 },  // header row
    { hpx: 24 },  // data row
  ];

  // ── Freeze panes ─────────────────────────────────────────────────────────────
  ws["!freeze"] = { xSplit: 1, ySplit: 1 };

  // ── Create workbook ─────────────────────────────────────────────────────────
  const wb = utils.book_new();
  utils.book_append_sheet(wb, ws, sheetName);

  // ── Set sheet tab color (Modern Office Blue) ─────────────────────────────────
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  if ((wb as any).sheets && (wb as any).sheets[0]) {
    (wb as any).sheets[0].tabColor = { rgb: "0078D4" };
  }

  // ── Generate XLSX buffer and download ────────────────────────────────────────
  const buffer = write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  saveAs(blob, filename);
}
