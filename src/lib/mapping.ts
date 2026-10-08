// Column mapping configuration for the "RUN" output sheet.
// Data rows start at row 2 (row 1 = headers). Columns are referenced by
// their Excel letters exactly as they appear in the template.

export type RawField =
  | "bank" // A  - SN (bank / scheme e.g. ACCESS)
  | "accountNumber" // C
  | "accountName" // D
  | "pin" // E  - PIN / pension number
  | "balance" // F  - numeric balance
  | "address" // Q
  | "nextOfKin" // R
  | "nin" // S
  | "bvn" // T
  | "phone" // U
  | "dob" // V  - date of birth
  | "employer" // Y
  | "signatory"; // AA - Sgns

// Which Excel column each raw field is written to.
export const RAW_COLUMN: Record<RawField, string> = {
  bank: "A",
  accountNumber: "C",
  accountName: "D",
  pin: "E",
  balance: "F",
  address: "Q",
  nextOfKin: "R",
  nin: "S",
  bvn: "T",
  phone: "U",
  dob: "V",
  employer: "Y",
  signatory: "AA",
};

// Header aliases (lower-cased, trimmed) used to detect raw fields in the
// uploaded input file. Add new aliases here if a different header is used.
export const HEADER_ALIASES: Record<RawField, string[]> = {
  bank: ["bank", "scheme", "pfa bank"],
  accountNumber: ["account number", "accountnumber", "productaccountnumber", "acct number", "acct no", "account no"],
  accountName: ["account name", "accountname", "name", "full name", "customer name"],
  pin: ["pin number", "pin", "pension number", "pension no", "rsa pin", "pen number", "pen", "pen number", "pension pin", "rsa pin number"],
  balance: ["balance", "rsa bal", "rsa balance", "bal", "rsa"],
  address: ["address"],
  nextOfKin: ["next of kin", "next of kins", "next of kin's", "nok", "nextofkin", "next of kin name", "nok name"],
  nin: ["nin", "national id", "national identification number"],
  bvn: ["bvn", "bank verification number"],
  phone: ["phone", "phone number", "phone no", "mobile", "gsm", "telephone"],
  dob: ["dob", "date of birth", "birth date", "d.o.b"],
  employer: ["employer", "employer name", "company", "pfa"],
  signatory: ["sgns", "signatory", "signature", "sign", "agent"],
};

// Formula columns. `{r}` is replaced with the current data row number.
// Word columns (H, J, L) are computed in JS from the referenced value and
// written as text, because Excel has no native number-to-words function.
export const FORMULA_COLUMN: Record<string, (r: number) => string> = {
  G: (r) => `F${r}/4+200000`, // Purchase price
  I: (r) => `F${r}/4`, // Equity Contribution
  K: (r) => `G${r}-300000`, // Forsale value
  M: (r) => `G${r}`, // Sum Insured
  N: (r) => `M${r}*0.21%`, // Fire Insurance
  O: () => `200000*1%`, // Credit Insurance
  P: (r) => `I${r}*2%`, // Processing fee
  W: (r) => `YEAR(V${r})`, // Year of birth
  Z: (r) => `Q${r}`, // Residential address = Address
};

// Age formula uses the current year, matching the template's `=2026-W{r}`.
export const ageFormula = (r: number, year: number) => `${year}-W${r}`;

// The number-to-words columns and the source column whose numeric value
// they describe.
export const WORD_COLUMNS: { target: string; source: RawField | "G" | "I" | "K" }[] = [
  { target: "H", source: "G" }, // Purchase Price Word
  { target: "J", source: "I" }, // Equity Contribution Word
  { target: "L", source: "K" }, // Forsale value word
];

export const SN_COLUMN = "B"; // sequential number
export const AGE_COLUMN = "X";
export const SHEET_NAME = "RUN";
export const DATA_START_ROW = 2;
export const LAST_COLUMN = "AA";
