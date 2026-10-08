import JSZip from "jszip";
import path from "path";
import { computeDerived } from "./fields";
import type { CleanRecord } from "./fields";
import { getCompressedTemplate, type TemplateCompression } from "./compressDocx";
import { getLetterTemplate, templatesDir } from "./letterTemplates";
import { docxBuffersToPdf } from "./pdfConvert";

function money(n: number): string {
  return (isNaN(n) ? 0 : n).toLocaleString("en-NG", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function parseNum(v: number | string | null | undefined): number {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  const n = Number(String(v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? 0 : n;
}

function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

// Convert an ISO (yyyy-mm-dd) or other date string to dd/mm/yyyy.
function formatDob(dob: string | undefined): string {
  if (!dob) return "";
  const iso = dob.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) return `${pad2(Number(iso[3]))}/${pad2(Number(iso[2]))}/${iso[1]}`;
  return dob;
}

function today(): string {
  const d = new Date();
  return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
}

// Map each Word MERGEFIELD name to its value for a member record.
export function buildMergeValues(
  record: CleanRecord,
  currentYear: number
): Record<string, string> {
  const v = computeDerived(record, currentYear);
  return {
    ACCOUNT_NAME: record.accountName ?? "",
    Address: record.address ?? "",
    RESIDENTIAL_ADDRESS: record.address ?? "",
    ACCOUNT_NUMBER: record.accountNumber ?? "",
    PIN_NUMBER: record.pin ?? "",
    BVN: record.bvn ?? "",
    NIN: record.nin ?? "",
    Next_of_KIN: record.nextOfKin ?? "",
    Phone: record.phone ?? "",
    DOB: formatDob(record.dob),
    BALANCE: money(parseNum(record.balance)),
    Purchase_price: money(v.purchasePrice),
    Purchace_Price_Word: v.purchaseWords,
    Equity_Contribution: money(v.equity),
    Forsale_value: money(v.forsale),
    Forsale_value_word: v.forsaleWords,
    Sum_Insured: money(v.sumInsured),
    Fire_Insurance: money(v.fireInsurance),
    Processing_fee_: money(v.processingFee),
    Sgns: record.signatory ?? "",
    date: today(),
  };
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Normalise a field name so different spellings/casings match, e.g.
// "Purchace_Price_Word" and "Purchase_price_word" both resolve, as do
// "Equity_Contribution" and "Equity_contribution".
function normName(k: string): string {
  return k
    .toLowerCase()
    .replace(/purchace/g, "purchase")
    .replace(/[^a-z0-9]/g, "");
}

// Build a resolver that maps any merge-field name to its value tolerantly.
function makeResolver(
  values: Record<string, string>
): (name: string) => string | undefined {
  const map: Record<string, string> = {};
  for (const [k, v] of Object.entries(values)) map[normName(k)] = v;
  return (name: string) => {
    if (name in values) return values[name];
    const n = normName(name);
    return n in map ? map[n] : undefined;
  };
}

function runXml(rPr: string, value: string): string {
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(value)}</w:t></w:r>`;
}

function firstRpr(fragment: string): string {
  const m = fragment.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
  return m ? m[0] : "";
}

// Extract the MERGEFIELD name from concatenated instrText runs.
function fieldNameFromInstr(instr: string): string | null {
  const m = instr.match(/MERGEFIELD\s+"?([A-Za-z0-9_]+)"?/);
  return m ? m[1] : null;
}

function fieldNameFromMiddle(middle: string): string | null {
  const instr = [...middle.matchAll(/<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/g)]
    .map((m) => m[1])
    .join("");
  return fieldNameFromInstr(instr);
}

// The result run's rPr (after the "separate" field char) preserves formatting.
function resultRpr(middle: string): string {
  const sep = middle.indexOf('w:fldCharType="separate"');
  const after = sep >= 0 ? middle.slice(sep) : middle;
  return firstRpr(after);
}

/**
 * Replace Word MERGEFIELDs (both fldSimple and complex field forms) with the
 * provided values, keyed by field name.
 */
export function replaceMergeFields(
  xml: string,
  values: Record<string, string>
): string {
  const resolve = makeResolver(values);

  // 1. Simple fields: <w:fldSimple w:instr=" MERGEFIELD NAME ">...</w:fldSimple>
  xml = xml.replace(
    /<w:fldSimple\b[^>]*\bw:instr="([^"]*)"[^>]*>([\s\S]*?)<\/w:fldSimple>/g,
    (whole, instr: string, inner: string) => {
      const name = fieldNameFromInstr(instr);
      const val = name ? resolve(name) : undefined;
      if (val !== undefined) {
        return runXml(firstRpr(inner), val);
      }
      return whole;
    }
  );

  // 2. Complex fields: begin ... instrText MERGEFIELD NAME ... separate ...
  //    [cached result runs] ... end
  const complexRe =
    /<w:r\b[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?\s*<w:fldChar w:fldCharType="begin"\/>\s*<\/w:r>([\s\S]*?)<w:r\b[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?\s*<w:fldChar w:fldCharType="end"\/>\s*<\/w:r>/g;
  xml = xml.replace(complexRe, (whole, middle: string) => {
    if (!/MERGEFIELD/.test(middle)) return whole;
    const name = fieldNameFromMiddle(middle);
    const val = name ? resolve(name) : undefined;
    if (val !== undefined) {
      return runXml(resultRpr(middle), val);
    }
    return whole;
  });

  // 3. Cross-run fallback for placeholders split across multiple runs:
  //    both «FieldName» (leftover merge-field cached results) and <<KEY>>.
  xml = replaceCrossRunPlaceholders(xml, values);

  return xml;
}

// Resolve a typed placeholder key (e.g. "Forsale value", "NAME") to a value.
function typedLookup(values: Record<string, string>): Record<string, string> {
  const norm = (k: string) => k.toLowerCase().replace(/[^a-z0-9]/g, "");
  const lookup: Record<string, string> = {};
  for (const [k, v] of Object.entries(values)) lookup[norm(k)] = v;
  lookup["name"] = values.ACCOUNT_NAME ?? "";
  lookup["forsalevalue"] = values.Forsale_value ?? "";
  lookup["forsalevalueword"] = values.Forsale_value_word ?? "";
  return lookup;
}

/**
 * Replace placeholders that Word has split across multiple <w:t> runs:
 *  - «FieldName»  (leftover merge-field cached results), resolved via `values`
 *  - <<KEY>>      (typed placeholders, stored as &lt;&lt;KEY&gt;&gt;)
 */
function replaceCrossRunPlaceholders(
  xml: string,
  values: Record<string, string>
): string {
  const lookup = typedLookup(values);
  const nodeRe = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;

  interface Node {
    openEnd: number;
    closeStart: number;
    text: string;
    newText: string;
  }
  const nodes: Node[] = [];
  let m: RegExpExecArray | null;
  while ((m = nodeRe.exec(xml))) {
    const openEnd = m.index + m[0].indexOf(">") + 1;
    const closeStart = m.index + m[0].length - "</w:t>".length;
    nodes.push({ openEnd, closeStart, text: m[1], newText: m[1] });
  }
  if (nodes.length === 0) return xml;

  let concat = "";
  const owner: number[] = [];
  const offsetInNode: number[] = [];
  for (let ni = 0; ni < nodes.length; ni++) {
    const t = nodes[ni].text;
    for (let ci = 0; ci < t.length; ci++) {
      owner.push(ni);
      offsetInNode.push(ci);
    }
    concat += t;
  }

  const edits: { ni: number; from: number; to: number; insert: string }[] = [];
  const schedule = (start: number, end: number, value: string) => {
    const firstNi = owner[start];
    const lastNi = owner[end - 1];
    for (let ni = firstNi; ni <= lastNi; ni++) {
      const from = ni === firstNi ? offsetInNode[start] : 0;
      const to =
        ni === lastNi ? offsetInNode[end - 1] + 1 : nodes[ni].text.length;
      edits.push({ ni, from, to, insert: ni === firstNi ? value : "" });
    }
  };

  // «FieldName» tokens
  const resolve = makeResolver(values);
  const guillemetRe = /\u00AB([^\u00AB\u00BB]{1,60})\u00BB/g;
  let g: RegExpExecArray | null;
  while ((g = guillemetRe.exec(concat))) {
    const name = g[1].trim();
    const val = resolve(name);
    if (val !== undefined) {
      schedule(g.index, g.index + g[0].length, escapeXml(val));
    }
  }

  // <<KEY>> tokens (escaped)
  const angleRe = /&lt;&lt;([\s\S]*?)&gt;&gt;/g;
  let a: RegExpExecArray | null;
  while ((a = angleRe.exec(concat))) {
    const key = a[1].replace(/\s+/g, " ").trim();
    const normKey = key.toLowerCase().replace(/[^a-z0-9]/g, "");
    if (normKey in lookup) {
      schedule(a.index, a.index + a[0].length, escapeXml(lookup[normKey]));
    }
  }

  if (edits.length === 0) return xml;

  const byNode = new Map<number, { from: number; to: number; insert: string }[]>();
  for (const e of edits) {
    if (!byNode.has(e.ni)) byNode.set(e.ni, []);
    byNode.get(e.ni)!.push(e);
  }
  byNode.forEach((list, ni) => {
    list.sort((x, y) => y.from - x.from);
    let t = nodes[ni].text;
    for (const e of list) {
      t = t.slice(0, e.from) + e.insert + t.slice(e.to);
    }
    nodes[ni].newText = t;
  });

  let result = "";
  let cursor = 0;
  for (const node of nodes) {
    result += xml.slice(cursor, node.openEnd);
    result += node.newText;
    cursor = node.closeStart;
  }
  result += xml.slice(cursor);
  return result;
}

// Remove the Office 2010 "artistic effect" image layers. These store a
// duplicate of each picture as an HD-Photo/JPEG-XR (.wdp) blip inside an
// <a:extLst> on the <a:blip>. LibreOffice cannot decode JPEG-XR, so when such a
// layer is present it renders nothing at all - which is why some signatures /
// stamps vanish from the PDF. Dropping just the <a:extLst> leaves the normal
// PNG/JPEG base image (referenced by a:blip r:embed) so it renders correctly.
// Other blip children (alphaModFix, duotone, etc.) are preserved.
function stripImageArtisticEffects(xml: string): string {
  return xml.replace(
    /(<a:blip\b[^>]*>)([\s\S]*?)(<\/a:blip>)/g,
    (_whole, open: string, inner: string, close: string) => {
      const cleaned = inner.replace(/<a:extLst>[\s\S]*?<\/a:extLst>/g, "");
      return open + cleaned + close;
    }
  );
}

// Remap font names that are NOT installed on the conversion host to the closest
// installed equivalent, so LibreOffice lays text out with the same metrics as
// Word instead of substituting a wider/taller fallback (which pushes paragraphs
// onto the next page). "Aptos"/"Aptos Display" are the new Office defaults that
// ship only with recent Office and are usually absent; "Arial MT" (and its
// PostScript bold/italic variants) is not a real Windows family name - only
// "Arial" is installed. Applied to every part that can carry a font name.
const FONT_REPLACEMENTS: { re: RegExp; to: string }[] = [
  { re: /Aptos Display/g, to: "Calibri" },
  { re: /Aptos/g, to: "Calibri" },
  { re: /Arial(?:-[A-Za-z]+)?MT|Arial MT/g, to: "Arial" },
];

function normalizeFontsForPdf(xml: string): string {
  let out = xml;
  for (const { re, to } of FONT_REPLACEMENTS) out = out.replace(re, to);
  return out;
}

// Fill one copy of the template with a member's values, returning docx bytes.
export async function populateDocx(
  templateData: Buffer,
  values: Record<string, string>,
  prepareForPdf = false
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(templateData);
  const targets = Object.keys(zip.files).filter((f) =>
    /^word\/(document|header\d+|footer\d+)\.xml$/.test(f)
  );
  for (const t of targets) {
    const file = zip.file(t);
    if (!file) continue;
    const xml = await file.async("string");
    let out = replaceMergeFields(xml, values);
    if (prepareForPdf) {
      out = stripImageArtisticEffects(out);
      out = normalizeFontsForPdf(out);
    }
    zip.file(t, out);
  }

  // For PDF, also remap fonts in the parts that declare or default them, so
  // paragraphs that inherit the theme/default font (e.g. Aptos) don't reflow.
  if (prepareForPdf) {
    const fontParts = Object.keys(zip.files).filter((f) =>
      /^word\/(styles|fontTable)\.xml$/.test(f) ||
      /^word\/theme\/theme\d+\.xml$/.test(f)
    );
    for (const p of fontParts) {
      const file = zip.file(p);
      if (!file) continue;
      const xml = await file.async("string");
      zip.file(p, normalizeFontsForPdf(xml));
    }
  }

  // Detach the document from its mail-merge data source so Word opens it as a
  // normal document (otherwise it prompts to select a data source / run SQL).
  // Also drop the <w:attachedTemplate> reference: its relationship is removed
  // below, and leaving the element behind creates a dangling r:id that makes
  // Word treat the file as corrupted when opened via automation (PDF export).
  const settings = zip.file("word/settings.xml");
  if (settings) {
    let s = await settings.async("string");
    s = s
      .replace(/<w:mailMerge>[\s\S]*?<\/w:mailMerge>/g, "")
      .replace(/<w:mailMerge\b[^>]*\/>/g, "")
      .replace(/<w:attachedTemplate\b[^>]*\/>/g, "");
    zip.file("word/settings.xml", s);
  }

  // Remove the external data-source relationships as well.
  const settingsRels = zip.file("word/_rels/settings.xml.rels");
  if (settingsRels) {
    let r = await settingsRels.async("string");
    r = r
      .replace(/<Relationship\b[^>]*Type="[^"]*mailMergeSource"[^>]*\/>/g, "")
      .replace(/<Relationship\b[^>]*Type="[^"]*attachedTemplate"[^>]*\/>/g, "");
    zip.file("word/_rels/settings.xml.rels", r);
  }

  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

// Build the output file base name: MEMBER NAME_PIN in upper case,
// e.g. "AHMED TANI SANI_PEN100045257373".
function displayName(record: CleanRecord, index: number): string {
  const name = (record.accountName || `MEMBER ${index + 1}`).trim().toUpperCase();
  const pin = (record.pin || "").trim().toUpperCase();
  const base = pin ? `${name}_${pin}` : name;
  return base
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
}

export type LetterFormat = "docx" | "pdf";

/**
 * Generate one populated letter per record using the chosen template, returned
 * as a single zip buffer. Each file is named `MEMBER NAME_PIN.<ext>`.
 */
export async function generateLetters(
  records: CleanRecord[],
  templateId?: string,
  format: LetterFormat = "docx",
  compression?: TemplateCompression
): Promise<Buffer> {
  const tpl = await getLetterTemplate(templateId);
  if (!tpl) throw new Error("No letter template is available.");

  const filePath = path.join(templatesDir(), tpl.file);
  // For PDF use full-resolution images (no downscale) so nothing is distorted,
  // and strip the JPEG-XR artistic-effect layers below so every signature
  // renders. DOCX keeps the size-optimised images.
  const mode = compression ?? (format === "pdf" ? "none" : "small");
  const templateData = await getCompressedTemplate(filePath, mode);
  const currentYear = new Date().getFullYear();

  const usedNames = new Map<string, number>();
  const items: { name: string; docx: Uint8Array }[] = [];

  for (let i = 0; i < records.length; i++) {
    const values = buildMergeValues(records[i], currentYear);
    const docx = await populateDocx(templateData, values, format === "pdf");

    let base = displayName(records[i], i);
    const n = (usedNames.get(base) ?? 0) + 1;
    usedNames.set(base, n);
    if (n > 1) base = `${base}_${n}`;
    items.push({ name: base, docx });
  }

  const out = new JSZip();

  if (format === "pdf") {
    const pdfs = await docxBuffersToPdf(items.map((it) => it.docx));
    items.forEach((it, idx) => out.file(`${it.name}.pdf`, pdfs[idx]));
  } else {
    items.forEach((it) => out.file(`${it.name}.docx`, it.docx));
  }

  return out.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
}
