import path from "path";
import { promises as fs } from "fs";

export interface LetterTemplate {
  id: string;
  label: string;
  suffix: string;
  file: string;
}

export function templatesDir(): string {
  return path.join(process.cwd(), "templates");
}

function slug(s: string): string {
  return s
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

interface ConfigEntry {
  file?: string;
  id?: string;
  label?: string;
  suffix?: string;
}

/**
 * List the available Word letter templates. Templates are `.docx` files in the
 * `templates/` folder. Friendly labels and filename suffixes come from
 * `templates/letter-templates.json`; any other `.docx` is auto-discovered.
 *
 * To add a new template: drop a `.docx` into `templates/` and (optionally) add
 * an entry to `letter-templates.json` for a nice label and filename suffix.
 */
export async function listLetterTemplates(): Promise<LetterTemplate[]> {
  const dir = templatesDir();

  let files: string[] = [];
  try {
    files = (await fs.readdir(dir)).filter((f) => /\.docx$/i.test(f));
  } catch {
    files = [];
  }

  let config: ConfigEntry[] = [];
  try {
    const raw = await fs.readFile(
      path.join(dir, "letter-templates.json"),
      "utf8"
    );
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) config = parsed;
  } catch {
    config = [];
  }

  const list: LetterTemplate[] = [];
  const seen = new Set<string>();
  const usedIds = new Set<string>();

  const addUnique = (id: string): string => {
    let candidate = id || "template";
    let n = 2;
    while (usedIds.has(candidate)) candidate = `${id}_${n++}`;
    usedIds.add(candidate);
    return candidate;
  };

  // Configured templates first (preserves ordering).
  for (const c of config) {
    if (!c || !c.file || !files.includes(c.file)) continue;
    const base = c.file.replace(/\.docx$/i, "");
    const id = addUnique(c.id || slug(c.suffix || base));
    list.push({
      id,
      label: c.label || base,
      suffix: slug(c.suffix || c.label || base),
      file: c.file,
    });
    seen.add(c.file);
  }

  // Auto-discovered templates not in the config.
  for (const f of files) {
    if (seen.has(f)) continue;
    const base = f.replace(/\.docx$/i, "");
    const id = addUnique(slug(base));
    list.push({ id, label: base, suffix: slug(base), file: f });
  }

  return list;
}

export async function getLetterTemplate(
  id?: string
): Promise<LetterTemplate | null> {
  const list = await listLetterTemplates();
  if (id) {
    const found = list.find((t) => t.id === id);
    if (found) return found;
  }
  return list[0] ?? null;
}
