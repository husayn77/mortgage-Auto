"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  CleanRecord,
  RAW_FIELDS,
  RawFieldKey,
  computeDerived,
} from "@/lib/fields";

interface Meta {
  detectedFields: string[];
  sheetName: string;
  headerRow: number;
}

interface TemplateOption {
  id: string;
  label: string;
  suffix: string;
}

const FIELD_LABELS: Record<string, string> = {
  bank: "Bank",
  accountNumber: "Account Number",
  accountName: "Account Name",
  pin: "PIN / PEN",
  balance: "Balance",
  address: "Address",
  nextOfKin: "Next of Kin",
  nin: "NIN",
  bvn: "BVN",
  phone: "Phone",
  dob: "Date of Birth",
  employer: "Employer",
  signatory: "Signatory",
};

const CALC_COLUMNS: {
  label: string;
  get: (d: ReturnType<typeof computeDerived>) => string;
}[] = [
  { label: "Purchase Price", get: (d) => money(d.purchasePrice) },
  { label: "Equity", get: (d) => money(d.equity) },
  { label: "Forsale", get: (d) => money(d.forsale) },
  { label: "Sum Insured", get: (d) => money(d.sumInsured) },
  { label: "Fire Ins.", get: (d) => money(d.fireInsurance) },
  { label: "Credit Ins.", get: (d) => money(d.creditInsurance) },
  { label: "Proc. Fee", get: (d) => money(d.processingFee) },
  { label: "Year", get: (d) => (d.year ?? "").toString() },
  { label: "Age", get: (d) => (d.age ?? "").toString() },
];

function money(n: number): string {
  return n.toLocaleString("en-NG", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [exportingLetters, setExportingLetters] = useState(false);
  const [activeTab, setActiveTab] = useState<"excel" | "letters">("excel");
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [records, setRecords] = useState<CleanRecord[] | null>(null);
  const [meta, setMeta] = useState<Meta | null>(null);
  const [templates, setTemplates] = useState<TemplateOption[]>([]);
  const [templateId, setTemplateId] = useState<string>("");
  const [letterFormat, setLetterFormat] = useState<"docx" | "pdf">("docx");
  const [compressionMode, setCompressionMode] = useState<
    "none" | "balanced" | "small"
  >("small");
  const inputRef = useRef<HTMLInputElement>(null);

  const currentYear = useMemo(() => new Date().getFullYear(), []);

  useEffect(() => {
    fetch("/api/templates")
      .then((r) => r.json())
      .then((d) => {
        const list: TemplateOption[] = d.templates ?? [];
        setTemplates(list);
        if (list.length > 0) setTemplateId((prev) => prev || list[0].id);
      })
      .catch(() => setTemplates([]));
  }, []);

  const reset = () => {
    setFile(null);
    setError(null);
    setSuccess(null);
    setRecords(null);
    setMeta(null);
    if (inputRef.current) inputRef.current.value = "";
  };

  const loadPreview = useCallback(async (f: File) => {
    setLoading(true);
    setError(null);
    setSuccess(null);
    setRecords(null);
    setMeta(null);
    try {
      const fd = new FormData();
      fd.append("file", f);
      const res = await fetch("/api/preview", { method: "POST", body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Failed to read file.");
      setRecords(data.records ?? []);
      setMeta({
        detectedFields: data.detectedFields ?? [],
        sheetName: data.sheetName,
        headerRow: data.headerRow,
      });
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to read file.");
    } finally {
      setLoading(false);
    }
  }, []);

  const onSelect = (f: File | null) => {
    setError(null);
    setSuccess(null);
    if (!f) return;
    if (!/\.(xlsx|xlsm|xls)$/i.test(f.name)) {
      setError("Please upload an Excel file (.xlsx, .xlsm or .xls).");
      return;
    }
    setFile(f);
    loadPreview(f);
  };

  const updateCell = (rowIndex: number, key: RawFieldKey, value: string) => {
    setRecords((prev) => {
      if (!prev) return prev;
      const next = [...prev];
      next[rowIndex] = {
        ...next[rowIndex],
        [key]:
          key === "balance" ? (value === "" ? null : Number(value)) : value,
      };
      return next;
    });
  };

  const addRow = () => {
    setRecords((prev) => [...(prev ?? []), {}]);
  };

  const deleteRow = (index: number) => {
    setRecords((prev) => (prev ? prev.filter((_, i) => i !== index) : prev));
  };

  const onExport = async () => {
    if (!records || records.length === 0) return;
    setExporting(true);
    setError(null);
    setSuccess(null);
    try {
      const res = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ records, fileName: file?.name ?? "output" }),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || "Export failed.");
      }
      const count = res.headers.get("X-Record-Count") ?? "";
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${(file?.name ?? "output").replace(
        /\.[^.]+$/,
        ""
      )} - MGT.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setSuccess(
        `Exported the MGT formula sheet for ${count} record${
          count === "1" ? "" : "s"
        }. Your download has started.`
      );
    } catch (e) {
      setError(e instanceof Error ? e.message : "Export failed.");
    } finally {
      setExporting(false);
    }
  };

  const onGenerateLetters = async () => {
    if (!records || records.length === 0) return;
    setExportingLetters(true);
    setError(null);
    setSuccess(null);
    try {
      const res = await fetch("/api/letters", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          records,
          fileName: file?.name ?? "output",
          templateId,
          format: letterFormat,
          compression: compressionMode,
        }),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || "Letter generation failed.");
      }
      const count = res.headers.get("X-Record-Count") ?? "";
      const tplLabel = decodeURIComponent(
        res.headers.get("X-Template-Label") ?? ""
      );
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${(file?.name ?? "output").replace(/\.[^.]+$/, "")} - ${
        tplLabel || "Letters"
      } Letters (${letterFormat.toUpperCase()}).zip`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setSuccess(
        `Generated ${count} ${tplLabel} welcome letter${
          count === "1" ? "" : "s"
        } as ${letterFormat.toUpperCase()} (one file per member) in a ZIP. Your download has started.`
      );
    } catch (e) {
      setError(e instanceof Error ? e.message : "Letter generation failed.");
    } finally {
      setExportingLetters(false);
    }
  };

  return (
    <main className="container">
      <div className="header">
        <div className="logo">M</div>
        <div>
          <h1>MGT Mapper</h1>
        </div>
      </div>
      <p className="subtitle">
        Upload a spreadsheet of member details, review &amp; edit the mapped
        data, then export the ready-to-use MGT Excel template with live formulas.
      </p>

      <div className="card">
        <div
          className={`dropzone${dragging ? " drag" : ""}`}
          onClick={() => inputRef.current?.click()}
          onDragOver={(e) => {
            e.preventDefault();
            setDragging(true);
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragging(false);
            onSelect(e.dataTransfer.files?.[0] ?? null);
          }}
        >
          <div className="icon">📊</div>
          <div>
            <strong>Click to upload</strong> or drag &amp; drop your input Excel
          </div>
          <div className="hint">
            .xlsx, .xlsm or .xls &mdash; headers auto-detected
          </div>
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xlsm,.xls"
            hidden
            onChange={(e) => onSelect(e.target.files?.[0] ?? null)}
          />
        </div>

        {file && (
          <div className="filepill">
            <div className="meta">
              <span>📄</span>
              <div style={{ minWidth: 0 }}>
                <div className="name">{file.name}</div>
                <div className="size">{formatBytes(file.size)}</div>
              </div>
            </div>
            <button className="iconbtn" onClick={reset} title="Remove">
              ✕ Remove
            </button>
          </div>
        )}

        {loading && (
          <div className="status info">
            <span className="spinner" />
            Reading file and detecting columns…
          </div>
        )}

        {meta && (
          <div className="status info">
            Loaded from sheet <strong>{meta.sheetName}</strong> (header row{" "}
            {meta.headerRow}). Detected columns:
            <div className="chips">
              {meta.detectedFields.map((f) => (
                <span className="chip" key={f}>
                  {FIELD_LABELS[f] ?? f}
                </span>
              ))}
            </div>
          </div>
        )}

        {error && <div className="status error">⚠️ {error}</div>}
        {success && <div className="status success">✅ {success}</div>}
      </div>

      {records && (
        <>
          <div className="tabs">
            <button
              className={`tab${activeTab === "excel" ? " active" : ""}`}
              onClick={() => setActiveTab("excel")}
            >
              📗 Excel Template
            </button>
            <button
              className={`tab${activeTab === "letters" ? " active" : ""}`}
              onClick={() => setActiveTab("letters")}
            >
              📄 Welcome Letters
            </button>
          </div>

          <div className="tabhint">
            {activeTab === "excel"
              ? "Fills the MGT workbook (RUN sheet) with live formulas and amount-in-words."
              : "Creates one populated Word document per member from the chosen letter template, delivered as a ZIP."}
          </div>

          {activeTab === "letters" && (
            <div className="templatebar">
              <label htmlFor="tpl">Letter template</label>
              <select
                id="tpl"
                className="tplselect"
                value={templateId}
                onChange={(e) => setTemplateId(e.target.value)}
              >
                {templates.length === 0 && <option value="">No templates found</option>}
                {templates.map((t) => (
                  <option key={t.id} value={t.id}>
                    {t.label}
                  </option>
                ))}
              </select>

              <span className="fmtlabel">Format</span>
              <div className="fmttoggle">
                <button
                  type="button"
                  className={`fmtopt${letterFormat === "docx" ? " active" : ""}`}
                  onClick={() => {
                    setLetterFormat("docx");
                    setCompressionMode("small");
                  }}
                >
                  Word
                </button>
                <button
                  type="button"
                  className={`fmtopt${letterFormat === "pdf" ? " active" : ""}`}
                  onClick={() => {
                    setLetterFormat("pdf");
                    setCompressionMode("none");
                  }}
                >
                  PDF
                </button>
              </div>

              <label htmlFor="compress" style={{ marginLeft: 8 }}>
                Quality
              </label>
              <select
                id="compress"
                className="tplselect"
                value={compressionMode}
                onChange={(e) =>
                  setCompressionMode(
                    e.target.value as "none" | "balanced" | "small"
                  )
                }
              >
                <option value="none">High quality (largest file)</option>
                <option value="balanced">Balanced</option>
                <option value="small">Small file (strong compression)</option>
              </select>

              <span className="tplnote">
                Files named{" "}
                <code>MEMBER NAME_PIN.{letterFormat}</code>
                {letterFormat === "pdf"
                  ? " · PDF uses Word/LibreOffice, may take longer"
                  : ""}
              </span>
            </div>
          )}

          <div className="toolbar">
            <span className="count-badge">
              <strong>{records.length}</strong> record
              {records.length === 1 ? "" : "s"} &mdash; edit any cell before
              export
            </span>
            <div className="spacer" />
            <button className="btn-ghost btn-sm" onClick={addRow}>
              + Add row
            </button>
            {activeTab === "excel" ? (
              <button
                className="btn-primary btn-sm"
                onClick={onExport}
                disabled={exporting || records.length === 0}
              >
                {exporting ? (
                  <>
                    <span className="spinner" />
                    Exporting…
                  </>
                ) : (
                  "Export MGT Excel"
                )}
              </button>
            ) : (
              <button
                className="btn-primary btn-sm"
                onClick={onGenerateLetters}
                disabled={
                  exportingLetters || records.length === 0 || !templateId
                }
              >
                {exportingLetters ? (
                  <>
                    <span className="spinner" />
                    Generating…
                  </>
                ) : (
                  `Generate Letters (${letterFormat.toUpperCase()} ZIP)`
                )}
              </button>
            )}
          </div>

          <div className="tablewrap">
            <table className="grid">
              <thead>
                <tr>
                  <th className="sn">S/N</th>
                  {RAW_FIELDS.map((col) => (
                    <th key={col.key} style={{ minWidth: col.width }}>
                      {col.label}
                    </th>
                  ))}
                  {CALC_COLUMNS.map((c) => (
                    <th key={c.label} className="calc">
                      {c.label}
                    </th>
                  ))}
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {records.map((rec, i) => {
                  const d = computeDerived(rec, currentYear);
                  return (
                    <tr key={i}>
                      <td className="sn">{i + 1}</td>
                      {RAW_FIELDS.map((col) => (
                        <td key={col.key} style={{ minWidth: col.width }}>
                          <input
                            className="cellinput"
                            type={
                              col.type === "date"
                                ? "date"
                                : col.type === "number"
                                ? "number"
                                : "text"
                            }
                            value={
                              rec[col.key] === null || rec[col.key] === undefined
                                ? ""
                                : String(rec[col.key])
                            }
                            onChange={(e) =>
                              updateCell(i, col.key, e.target.value)
                            }
                          />
                        </td>
                      ))}
                      {CALC_COLUMNS.map((c) => (
                        <td key={c.label} className="calc">
                          {c.get(d)}
                        </td>
                      ))}
                      <td>
                        <button
                          className="rowdel"
                          title="Delete row"
                          onClick={() => deleteRow(i)}
                        >
                          🗑
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </>
      )}

      {!records && (
        <div className="features">
          <div className="feature">
            <div className="t">Live formulas</div>
            <div className="d">
              Purchase price, equity, forsale value, insurance and fees are
              written as real Excel formulas that recalculate.
            </div>
          </div>
          <div className="feature">
            <div className="t">Review &amp; edit</div>
            <div className="d">
              Every mapped value is shown in a table you can correct before
              exporting. Computed columns update live.
            </div>
          </div>
          <div className="feature">
            <div className="t">Amount in words</div>
            <div className="d">
              Naira &amp; Kobo word columns are generated automatically in the
              exact template style.
            </div>
          </div>
        </div>
      )}

      <div className="footnote">
        One input file → one filled MGT formula workbook.
      </div>
    </main>
  );
}
