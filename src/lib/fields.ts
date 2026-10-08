import { amountToWords } from "./numberToWords";

// A cleaned, editable record: plain primitives suitable for JSON transport,
// display in the UI table, and round-tripping back for export.
export interface CleanRecord {
  bank?: string;
  accountNumber?: string;
  accountName?: string;
  pin?: string;
  balance?: number | null;
  address?: string;
  nextOfKin?: string;
  nin?: string;
  bvn?: string;
  phone?: string;
  dob?: string;
  employer?: string;
  signatory?: string;
}

export type RawFieldKey = keyof CleanRecord;

// Editable columns shown in the UI table, in display order.
export const RAW_FIELDS: {
  key: RawFieldKey;
  label: string;
  type: "text" | "number" | "date";
  width?: number;
}[] = [
  { key: "bank", label: "Bank", type: "text", width: 130 },
  { key: "accountNumber", label: "Account No", type: "text", width: 130 },
  { key: "accountName", label: "Account Name", type: "text", width: 200 },
  { key: "pin", label: "PIN / PEN", type: "text", width: 160 },
  { key: "balance", label: "Balance", type: "number", width: 130 },
  { key: "address", label: "Address", type: "text", width: 220 },
  { key: "nextOfKin", label: "Next of Kin", type: "text", width: 180 },
  { key: "nin", label: "NIN", type: "text", width: 130 },
  { key: "bvn", label: "BVN", type: "text", width: 130 },
  { key: "phone", label: "Phone", type: "text", width: 130 },
  { key: "dob", label: "DOB", type: "date", width: 140 },
  { key: "employer", label: "Employer", type: "text", width: 160 },
  { key: "signatory", label: "Signatory", type: "text", width: 160 },
];

export interface Derived {
  purchasePrice: number;
  equity: number;
  forsale: number;
  sumInsured: number;
  fireInsurance: number;
  creditInsurance: number;
  processingFee: number;
  year: number | null;
  age: number | null;
  purchaseWords: string;
  equityWords: string;
  forsaleWords: string;
}

function parseBalance(value: number | string | null | undefined): number {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return isNaN(value) ? 0 : value;
  const cleaned = String(value).replace(/[^0-9.\-]/g, "");
  const n = Number(cleaned);
  return isNaN(n) ? 0 : n;
}

function yearFromDob(dob: string | undefined): number | null {
  if (!dob) return null;
  const iso = dob.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) return Number(iso[1]);
  const dmy = dob.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (dmy) {
    let y = Number(dmy[3]);
    if (y < 100) y += y < 30 ? 2000 : 1900;
    return y;
  }
  const d = new Date(dob);
  return isNaN(d.getTime()) ? null : d.getUTCFullYear();
}

// Compute the derived columns exactly as the Excel formulas do.
export function computeDerived(
  rec: CleanRecord,
  currentYear: number
): Derived {
  const bal = parseBalance(rec.balance);
  const purchasePrice = bal / 4 + 200000;
  const equity = bal / 4;
  const forsale = purchasePrice - 300000;
  const sumInsured = purchasePrice;
  const fireInsurance = sumInsured * 0.0021;
  const creditInsurance = 2000;
  const processingFee = equity * 0.02;
  const year = yearFromDob(rec.dob);
  const age = year !== null ? currentYear - year : null;

  return {
    purchasePrice,
    equity,
    forsale,
    sumInsured,
    fireInsurance,
    creditInsurance,
    processingFee,
    year,
    age,
    purchaseWords: amountToWords(purchasePrice),
    equityWords: amountToWords(equity),
    forsaleWords: amountToWords(forsale),
  };
}
