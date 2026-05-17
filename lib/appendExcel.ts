/**
 * appendExcel.ts — Append production data to existing Excel workbooks
 * Uses industry-standard XLSX (SheetJS) library
 * Modern Microsoft Office 365 Design
 */

import { read, utils, write } from "xlsx";
import { saveAs } from "file-saver";
import type { ParsedReport } from "./parser";

// ── Modern Microsoft Office 365 Design Tokens ────────────────────────────────
const HEADER_BG = "0078D4";   // Modern Office Blue
const HEADER_FG = "FFFFFF";   // White
const ROW_ODD   = "FFFFFF";   // Clean White
const ROW_EVEN  = "F5F5F5";   // Subtle Light Grey

const STRING_KEYS = new Set(["Date", "Big Tank", "Hermes Tank", "Process Water Tank"]);

function parseDateStr(s: unknown): Date | null {
  if (typeof s !== "string" || !s) return null;
  for (const re of [/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, /^(\d{1,2})-(\d{1,2})-(\d{4})$/]) {
    const m = s.match(re);
    if (m) {
      const d = new Date(+m[3], +m[2] - 1, +m[1]);
      if (!isNaN(d.getTime())) return d;
    }
  }
  return null;
}

function dateToStr(v: unknown): string {
  if (v instanceof Date) return `${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;
  return String(v ?? "").trim();
}

function styleCell(s: any, bg: string, h: string, v: any) {
  const isDateCol = h === "Date";
  const isPercentCol = h.includes("%");
  const isZero = typeof v === "number" && v === 0;

  s.fill = { patternType: "solid", fgColor: { rgb: bg }, theme: undefined, tint: undefined };
  s.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  s.border = {
    left: { style: "thin", color: { rgb: "D0D0D0" } },
    right: { style: "thin", color: { rgb: "D0D0D0" } },
    top: { style: "thin", color: { rgb: "D0D0D0" } },
    bottom: { style: "thin", color: { rgb: "D0D0D0" } },
  };
  s.font = {
    name: "Segoe UI",
    sz: 10,
    bold: isDateCol,
    color: {
      rgb: isDateCol ? "0078D4" : isZero ? "A8AAAD" : typeof v === "number" ? "242424" : "606060",
    },
    italic: typeof v === "string" && v !== "",
  };
  s.numFmt = isDateCol ? "DD/MM/YYYY" : isPercentCol ? "0.00%" : typeof v === "number" ? "#,##0.000" : "@";
  return s;
}

// ── Public API ────────────────────────────────────────────────────────────────

export function canUseFileSystemAccess(): boolean {
  return typeof window !== "undefined" && "showOpenFilePicker" in window;
}

export async function pickExcelFile(): Promise<{ handle: FileSystemFileHandle; buffer: ArrayBuffer } | null> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const [handle] = await (window as any).showOpenFilePicker({
      types: [{ description: "Excel Workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }],
      multiple: false,
    });
    return { handle, buffer: await (await handle.getFile()).arrayBuffer() };
  } catch (e: unknown) {
    if ((e as { name?: string })?.name === "AbortError") return null;
    throw e;
  }
}

export async function getSheetNames(buffer: ArrayBuffer): Promise<string[]> {
  const wb = read(buffer, { type: "array" });
  return wb.SheetNames;
}

export interface AppendStats {
  totalRows: number;
  newCols: string[];
  duplicate: boolean;
}

export async function buildAppendedBuffer(
  existingBuffer: ArrayBuffer | null,
  data: ParsedReport,
  sheetName: string,
): Promise<{ buffer: ArrayBuffer; stats: AppendStats }> {
  let wb: any;
  let ws: any;
  let headers: string[] = [];
  const newCols: string[] = [];
  let duplicate = false;

  if (existingBuffer) {
    wb = read(existingBuffer, { type: "array" });
    if (wb.SheetNames.includes(sheetName)) {
      ws = wb.Sheets[sheetName];
      const aoa = utils.sheet_to_aoa(ws);
      if (aoa && aoa.length > 0) {
        headers = aoa[0] as string[];
      }
    } else {
      ws = utils.aoa_to_sheet([]);
    }
  } else {
    wb = utils.book_new();
    ws = utils.aoa_to_sheet([]);
  }

  const incoming = Object.keys(data);

  // Check for new columns
  if (headers.length > 0) {
    for (const k of incoming) {
      if (!headers.includes(k)) {
        newCols.push(k);
        headers.push(k);
      }
    }
  } else {
    headers = incoming;
  }

  // Add header if new sheet
  if (!existingBuffer || !wb.SheetNames.includes(sheetName)) {
    const headerRow = headers.map(h => h);
    const headerAoa = [headerRow];
    ws = utils.aoa_to_sheet(headerAoa);
    for (let c = 0; c < headers.length; c++) {
      const cellRef = utils.encode_cell({ r: 0, c });
      if (!ws[cellRef]) ws[cellRef] = {};
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
  }

  // Duplicate check
  const todayStr = String(data["Date"] ?? "").trim();
  let aoa = utils.sheet_to_aoa(ws);
  const dateIdx = headers.indexOf("Date");

  if (todayStr && dateIdx >= 0 && aoa && aoa.length > 1) {
    for (let r = 1; r < aoa.length; r++) {
      if (dateToStr(aoa[r]?.[dateIdx]) === todayStr) {
        duplicate = true;
        break;
      }
    }
  }

  // Append new row if not duplicate
  if (!duplicate) {
    const rowData = headers.map(h => {
      const v = data[h];
      if (v === undefined) return STRING_KEYS.has(h) ? "" : 0;
      if (h === "Date" && typeof v === "string") return parseDateStr(v) ?? v;
      return v;
    });

    if (!aoa || aoa.length === 0) {
      aoa = [headers];
    }
    aoa.push(rowData);

    // Sort by date if present
    if (dateIdx >= 0 && aoa.length > 2) {
      const header = aoa[0];
      const rows = aoa.slice(1);
      rows.sort((a, b) => {
        const da = a[dateIdx] instanceof Date ? (a[dateIdx] as Date) : parseDateStr(String(a[dateIdx] ?? ""));
        const db = b[dateIdx] instanceof Date ? (b[dateIdx] as Date) : parseDateStr(String(b[dateIdx] ?? ""));
        if (!da && !db) return 0;
        if (!da) return 1;
        if (!db) return -1;
        return da.getTime() - db.getTime();
      });
      aoa = [header, ...rows];
    }

    ws = utils.aoa_to_sheet(aoa);

    // Style all rows
    for (let r = 0; r < aoa.length; r++) {
      for (let c = 0; c < headers.length; c++) {
        const cellRef = utils.encode_cell({ r, c });
        const v = aoa[r]?.[c];
        const h = headers[c];
        const bg = r === 0 ? HEADER_BG : r % 2 === 0 ? ROW_EVEN : ROW_ODD;

        if (!ws[cellRef]) ws[cellRef] = {};
        if (r === 0) {
          ws[cellRef].v = ws[cellRef].v || (aoa[r]?.[c] ?? "");
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
        } else {
          ws[cellRef].s = styleCell({}, bg, h, v);
        }
      }
    }
  }

  // Set column widths
  ws["!cols"] = headers.map((h, i) => {
    const aoa = utils.sheet_to_aoa(ws);
    let w = h.length + 3;
    if (aoa) {
      for (let r = 1; r < aoa.length; r++) {
        const v = aoa[r]?.[i];
        const l = v instanceof Date ? 12 : String(v ?? "").length + 2;
        if (l > w) w = l;
      }
    }
    return { wch: Math.max(w, 14) };
  });

  // Set row heights
  aoa = utils.sheet_to_aoa(ws);
  ws["!rows"] = [];
  if (aoa) {
    ws["!rows"][0] = { hpx: 40 };
    for (let i = 1; i < aoa.length; i++) {
      ws["!rows"][i] = { hpx: 24 };
    }
  }

  // Freeze panes
  ws["!freeze"] = { xSplit: 1, ySplit: 1 };

  // Add or update sheet in workbook
  if (!wb.SheetNames.includes(sheetName)) {
    utils.book_append_sheet(wb, ws, sheetName);
  } else {
    wb.Sheets[sheetName] = ws;
  }

  // Set sheet tab color
  const idx = wb.SheetNames.indexOf(sheetName);
  if (idx >= 0 && wb.sheets) {
    wb.sheets[idx].tabColor = { rgb: "0078D4" };
  }

  const buffer = write(wb, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
  const totalRows = Math.max(0, (aoa?.length ?? 1) - 1);

  return {
    buffer,
    stats: { totalRows, newCols, duplicate },
  };
}

export async function saveToHandle(handle: FileSystemFileHandle, buffer: ArrayBuffer) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const writable = await (handle as any).createWritable();
  await writable.write(buffer);
  await writable.close();
}

export function downloadBuffer(buffer: ArrayBuffer, filename: string) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  saveAs(blob, filename);
}
