import { spawn } from "child_process";
import { promises as fs } from "fs";
import os from "os";
import path from "path";
import { randomBytes } from "crypto";

// Common install locations for LibreOffice's headless converter. On Windows the
// console launcher `soffice.com` is listed first: unlike `soffice.exe`, it blocks
// until the conversion finishes (soffice.exe detaches and returns immediately,
// so the output PDFs may not exist yet when it exits).
const SOFFICE_CANDIDATES = [
  "C:/Program Files/LibreOffice/program/soffice.com",
  "C:/Program Files (x86)/LibreOffice/program/soffice.com",
  "C:/Program Files/LibreOffice/program/soffice.exe",
  "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
  "/usr/bin/soffice",
  "/usr/local/bin/soffice",
  "/Applications/LibreOffice.app/Contents/MacOS/soffice",
];

async function findSoffice(): Promise<string | null> {
  for (const p of SOFFICE_CANDIDATES) {
    try {
      await fs.access(p);
      return p;
    } catch {
      // continue
    }
  }
  return null;
}

function run(
  cmd: string,
  args: string[],
  timeoutMs: number
): Promise<{ code: number | null; stderr: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, { windowsHide: true });
    let stderr = "";
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error("PDF conversion timed out."));
    }, timeoutMs);
    // Drain both streams. soffice is verbose on stdout; if it is left unread the
    // OS pipe buffer fills up and the child blocks forever (a silent hang).
    child.stdout.on("data", () => {});
    child.stderr.on("data", (d) => (stderr += d.toString()));
    child.on("error", (err) => {
      clearTimeout(timer);
      reject(err);
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      resolve({ code, stderr });
    });
  });
}

// Convert docx files (0.docx..N-1.docx) in `dir` to PDFs using LibreOffice.
async function convertWithSoffice(
  soffice: string,
  dir: string,
  count: number
): Promise<void> {
  const files = Array.from({ length: count }, (_, i) =>
    path.join(dir, `${i}.docx`)
  );
  // Use a private, single-use profile directory. LibreOffice is single-instance
  // per profile: if another instance is already running with the default
  // profile, our command is forwarded to it and returns without converting.
  const profileDir = path.join(dir, "louser");
  const profileUri = "file:///" + profileDir.replace(/\\/g, "/");
  try {
    await run(
      soffice,
      [
        `-env:UserInstallation=${profileUri}`,
        "--headless",
        "--norestore",
        "--nologo",
        "--nolockcheck",
        "--convert-to",
        "pdf",
        "--outdir",
        dir,
        ...files,
      ],
      // Allow generous headroom: a cold LibreOffice start initialises its
      // profile (~15-25s) before converting the first document, and large
      // high-fidelity images make each conversion slower.
      Math.max(240000, 60000 + count * 30000)
    );
  } finally {
    await fs.rm(profileDir, { recursive: true, force: true }).catch(() => {});
  }
}

// Convert docx files in `dir` to PDFs using Microsoft Word via COM automation.
async function convertWithWord(dir: string, count: number): Promise<void> {
  const script = `
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try { $word.DisplayAlerts = 0 } catch {}
try {
  for ($i = 0; $i -lt ${count}; $i++) {
    $src = Join-Path '${dir.replace(/\\/g, "\\\\")}' ("" + $i + ".docx")
    $dst = Join-Path '${dir.replace(/\\/g, "\\\\")}' ("" + $i + ".pdf")
    $doc = $word.Documents.Open($src, $false, $true)
    $doc.ExportAsFixedFormat($dst, 17)
    $doc.Close($false)
  }
} finally {
  $word.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
`;
  const scriptPath = path.join(dir, "_convert.ps1");
  await fs.writeFile(scriptPath, script, "utf8");
  const { code, stderr } = await run(
    "powershell.exe",
    ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", scriptPath],
    Math.max(90000, count * 20000)
  );
  if (code !== 0) {
    throw new Error(
      `Microsoft Word could not convert the documents to PDF.${
        stderr ? " " + stderr.trim() : ""
      }`
    );
  }
}

/**
 * Convert an array of .docx buffers to .pdf buffers. Uses LibreOffice if
 * available, otherwise falls back to Microsoft Word (Windows COM automation).
 */
export async function docxBuffersToPdf(
  docs: Uint8Array[]
): Promise<Buffer[]> {
  if (docs.length === 0) return [];

  const dir = path.join(os.tmpdir(), `mgt-pdf-${randomBytes(6).toString("hex")}`);
  await fs.mkdir(dir, { recursive: true });

  try {
    for (let i = 0; i < docs.length; i++) {
      await fs.writeFile(path.join(dir, `${i}.docx`), Buffer.from(docs[i]));
    }

    const soffice = await findSoffice();
    if (soffice) {
      await convertWithSoffice(soffice, dir, docs.length);
    } else if (process.platform === "win32") {
      await convertWithWord(dir, docs.length);
    } else {
      throw new Error(
        "No PDF converter found. Please install LibreOffice to enable PDF export."
      );
    }

    const pdfs: Buffer[] = [];
    for (let i = 0; i < docs.length; i++) {
      const pdfPath = path.join(dir, `${i}.pdf`);
      try {
        pdfs.push(await fs.readFile(pdfPath));
      } catch {
        throw new Error(
          "PDF conversion failed for one or more documents. Ensure Microsoft Word or LibreOffice is installed."
        );
      }
    }
    return pdfs;
  } finally {
    await fs.rm(dir, { recursive: true, force: true }).catch(() => {});
  }
}
