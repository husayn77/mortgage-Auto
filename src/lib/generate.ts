import ExcelJS from "exceljs";
import path from "path";
import { promises as fs } from "fs";
import { amountToWords } from "./numberToWords";
import {
  RAW_COLUMN,
  HEADER_ALIASES,
  FORMULA_COLUMN,
  WORD_COLUMNS,
  ageFormula,
  SN_COLUMN,
  AGE_COLUMN,
  SHEET_NAME,
  DATA_START_ROW,
  RawField,
} from "./mapping";

const COLS = [..."ABCDEFGHIJKLMNOPQRSTUVWXYZ".split(""), "AA"];

import type { CleanRecord } from "./fields";

export interface ParseResult {
  records: CleanRecord[];
  detectedFields: RawField[];
  sheetName: string;
  headerRow: number;
}

// Build a lookup of alias -> field for fast header detection.
const ALIAS_LOOKUP: Record<string, RawField> = (() => {
  const map: Record<string, RawField> = {};
  (Object.keys(HEADER_ALIASES) as RawField[]).forEach((field) => {
    HEADER_ALIASES[field].forEach((alias) => {
      map[normalize(alias)] = field;
    });
  });
  return map;
})();

function normalize(value: unknown): string {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function cellText(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "object") {
    // ExcelJS rich text / formula / hyperlink objects
    const anyVal = value as unknown as Record<string, unknown>;
    if ("text" in anyVal) return String(anyVal.text ?? "");
    if ("result" in anyVal) return String(anyVal.result ?? "");
    if ("richText" in anyVal && Array.isArray(anyVal.richText)) {
      return (anyVal.richText as { text: string }[]).map((r) => r.text).join("");
    }
  }
  return String(value);
}

/**
 * Parse an uploaded workbook buffer, auto-detecting the header row and
 * mapping columns to known raw fields.
 */
export async function parseInput(buffer: ArrayBuffer): Promise<ParseResult> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  let best: {
    ws: ExcelJS.Worksheet;
    headerRow: number;
    colMap: Map<number, RawField>;
    matches: number;
  } | null = null;

  for (const ws of wb.worksheets) {
    const maxScan = Math.min(15, ws.rowCount || 15);
    for (let r = 1; r <= maxScan; r++) {
      const row = ws.getRow(r);
      const colMap = new Map<number, RawField>();
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const key = normalize(cellText(cell.value));
        const field = ALIAS_LOOKUP[key];
        if (field && ![...colMap.values()].includes(field)) {
          colMap.set(colNumber, field);
        }
      });
      if (!best || colMap.size > best.matches) {
        best = { ws, headerRow: r, colMap, matches: colMap.size };
      }
    }
  }

  if (!best || best.matches === 0) {
    throw new Error(
      "Could not detect any known columns in the uploaded file. Expected headers like Account Number, Account Name, PIN, Balance, DOB, etc."
    );
  }

  const { ws, headerRow, colMap } = best;
  const detectedFields = [...colMap.values()];
  const records: CleanRecord[] = [];

  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const record: CleanRecord = {};
    let hasData = false;
    colMap.forEach((field, colNumber) => {
      const value = row.getCell(colNumber).value;
      if (value === null || value === undefined || cellText(value).trim() === "") {
        return;
      }
      hasData = true;
      if (field === "balance") {
        record.balance = toNumber(value);
      } else if (field === "dob") {
        record.dob = normalizeDob(value);
      } else {
        record[field] = cellText(value).trim();
      }
    });
    if (hasData) records.push(record);
  }

  return { records, detectedFields, sheetName: ws.name, headerRow };
}

function toNumber(value: ExcelJS.CellValue): number | null {
  if (value === null || value === undefined) return null;
  if (typeof value === "number") return value;
  if (typeof value === "object" && "result" in (value as object)) {
    const r = (value as { result?: unknown }).result;
    if (typeof r === "number") return r;
  }
  const cleaned = cellText(value).replace(/[^0-9.\-]/g, "");
  if (cleaned === "" || cleaned === "-" || cleaned === ".") return null;
  const n = Number(cleaned);
  return isNaN(n) ? null : n;
}

// Coerce an edited value (number | string | null) into a number.
function toNumberLoose(value: number | string | null | undefined): number | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return isNaN(value) ? null : value;
  const cleaned = String(value).replace(/[^0-9.\-]/g, "");
  if (cleaned === "" || cleaned === "-" || cleaned === ".") return null;
  const n = Number(cleaned);
  return isNaN(n) ? null : n;
}

function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

// Normalise a date-of-birth cell into an ISO yyyy-mm-dd string for editing.
function normalizeDob(value: ExcelJS.CellValue): string {
  const d = toDate(value);
  if (d) {
    return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(
      d.getUTCDate()
    )}`;
  }
  return cellText(value).trim();
}

// Parse a date of birth into a real Date so Excel's YEAR() works.
function toDate(value: ExcelJS.CellValue | string | null): Date | null {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    // Excel serial date (days since 1899-12-30)
    const ms = Math.round((value - 25569) * 86400 * 1000);
    const d = new Date(ms);
    return isNaN(d.getTime()) ? null : d;
  }
  const text = cellText(value as ExcelJS.CellValue).trim();
  if (!text) return null;

  // ISO yyyy-mm-dd
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    const [, y, mm, dd] = iso;
    const d = new Date(Date.UTC(Number(y), Number(mm) - 1, Number(dd)));
    return isNaN(d.getTime()) ? null : d;
  }

  // Try dd/mm/yyyy or dd-mm-yyyy (common Nigerian format)
  const m = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) {
    const [, dd, mm, yyyy] = m;
    let year = Number(yyyy);
    if (year < 100) year += year < 30 ? 2000 : 1900;
    const d = new Date(Date.UTC(year, Number(mm) - 1, Number(dd)));
    return isNaN(d.getTime()) ? null : d;
  }

  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function cloneStyle(style: Partial<ExcelJS.Style>): Partial<ExcelJS.Style> {
  return JSON.parse(JSON.stringify(style ?? {}));
}

export function getTemplatePath(): string {
  return path.join(process.cwd(), "templates", "EXCEL TEMPLATE.xlsx");
}

/**
 * Fill the template's RUN sheet with the given records, applying live
 * formulas per row and computed number-to-words text.
 */
export async function generateWorkbook(records: CleanRecord[]): Promise<Buffer> {
  const templatePath = getTemplatePath();
  const templateBuffer = await fs.readFile(templatePath);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(
    templateBuffer.buffer.slice(
      templateBuffer.byteOffset,
      templateBuffer.byteOffset + templateBuffer.byteLength
    )
  );

  const ws = wb.getWorksheet(SHEET_NAME);
  if (!ws) throw new Error(`Template is missing the "${SHEET_NAME}" sheet.`);

  // Capture the style of the first data row so new rows look identical.
  const styleTemplate: Record<string, Partial<ExcelJS.Style>> = {};
  COLS.forEach((col) => {
    styleTemplate[col] = cloneStyle(ws.getCell(`${col}${DATA_START_ROW}`).style);
  });

  const originalLastRow = ws.rowCount;

  const currentYear = new Date().getFullYear();

  records.forEach((record, index) => {
    const r = DATA_START_ROW + index;

    // --- Raw fields -------------------------------------------------
    const balance = toNumberLoose(record.balance ?? null);
    const dob = toDate(record.dob ?? null);

    (Object.keys(RAW_COLUMN) as RawField[]).forEach((field) => {
      if (field === "balance" || field === "dob") return; // handled below
      const value = record[field];
      if (value === undefined || value === null || value === "") return;
      ws.getCell(`${RAW_COLUMN[field]}${r}`).value = String(value);
    });

    if (balance !== null) ws.getCell(`F${r}`).value = balance;
    if (dob) ws.getCell(`V${r}`).value = dob;

    // --- Sequential number -----------------------------------------
    ws.getCell(`${SN_COLUMN}${r}`).value = index + 1;

    // --- Computed numeric results (for immediate display) ----------
    const bal = balance ?? 0;
    const G = bal / 4 + 200000;
    const I = bal / 4;
    const K = G - 300000;
    const M = G;
    const N = M * 0.0021;
    const O = 2000;
    const P = I * 0.02;

    const results: Record<string, number> = { G, I, K, M, N, O, P };

    // --- Formula columns -------------------------------------------
    Object.entries(FORMULA_COLUMN).forEach(([col, fn]) => {
      const cell = ws.getCell(`${col}${r}`);
      if (col === "Z") {
        cell.value = { formula: fn(r), result: record.address ?? "" };
      } else if (col === "W") {
        cell.value = {
          formula: fn(r),
          result: dob ? dob.getUTCFullYear() : undefined,
        };
      } else {
        cell.value = { formula: fn(r), result: results[col] };
      }
    });

    // Age column (uses current year)
    ws.getCell(`${AGE_COLUMN}${r}`).value = {
      formula: ageFormula(r, currentYear),
      result: dob ? currentYear - dob.getUTCFullYear() : undefined,
    };

    // --- Number-to-words columns (text) ----------------------------
    WORD_COLUMNS.forEach(({ target, source }) => {
      const numeric = results[source as string];
      ws.getCell(`${target}${r}`).value =
        balance !== null ? amountToWords(numeric) : "";
    });

    // --- Apply captured styling ------------------------------------
    COLS.forEach((col) => {
      const cell = ws.getCell(`${col}${r}`);
      cell.style = cloneStyle(styleTemplate[col]);
    });
  });

  // Clear any leftover sample rows beyond the data we just wrote.
  const clearFrom = DATA_START_ROW + records.length;
  for (let r = clearFrom; r <= originalLastRow; r++) {
    COLS.forEach((col) => {
      ws.getCell(`${col}${r}`).value = null;
    });
  }

  // Clear sample data from the secondary source sheet, if present.
  const sheet1 = wb.getWorksheet("Sheet1");
  if (sheet1) {
    const last1 = sheet1.rowCount;
    for (let r = 2; r <= last1; r++) {
      const row = sheet1.getRow(r);
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.value = null;
      });
    }
  }

  const out = await wb.xlsx.writeBuffer();
  return Buffer.from(out);
}
