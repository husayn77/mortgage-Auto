import JSZip from "jszip";
import sharp from "sharp";
import { promises as fs } from "fs";

export type TemplateCompression = "none" | "balanced" | "small";

const cache = new Map<string, { mtimeMs: number; data: Buffer }>();

// Recompress the images inside a .docx to dramatically reduce file size.
export async function compressDocxImages(
  buffer: Buffer,
  mode: TemplateCompression = "small"
): Promise<Buffer> {
  if (mode === "none") return buffer;

  const zip = await JSZip.loadAsync(buffer);
  const media = Object.keys(zip.files).filter((f) =>
    /^word\/media\/.+\.(png|jpe?g)$/i.test(f)
  );

  const maxWidth = mode === "balanced" ? 2200 : 1400;
  const pngQuality = mode === "balanced" ? 92 : 80;
  const jpegQuality = mode === "balanced" ? 85 : 62;
  const pngPalette = mode === "balanced" ? false : true;

  for (const name of media) {
    const file = zip.file(name);
    if (!file) continue;
    const data = await file.async("nodebuffer");
    const ext = name.split(".").pop()!.toLowerCase();
    try {
      let img = sharp(data, { failOn: "none" }).rotate();
      const meta = await img.metadata();
      if (meta.width && meta.width > maxWidth) img = img.resize({ width: maxWidth });

      const out =
        ext === "png"
          ? await img
              .png({
                compressionLevel: 9,
                quality: pngQuality,
                palette: pngPalette,
                effort: 7,
              })
              .toBuffer()
          : await img.jpeg({ quality: jpegQuality, mozjpeg: true }).toBuffer();

      if (out.length < data.length) {
        zip.file(name, out);
      }
    } catch {
      // Leave the original image if it can't be decoded/re-encoded.
    }
  }

  return zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
}

/**
 * Return a size-optimised copy of a template file, cached in memory keyed by
 * file path + modified time so compression only runs once per template.
 */
export async function getCompressedTemplate(
  filePath: string,
  mode: TemplateCompression = "small"
): Promise<Buffer> {
  if (mode === "none") {
    return fs.readFile(filePath);
  }

  const cacheKey = `${filePath}::${mode}`;
  const stat = await fs.stat(filePath);
  const hit = cache.get(cacheKey);
  if (hit && hit.mtimeMs === stat.mtimeMs) return hit.data;

  const raw = await fs.readFile(filePath);
  const compressed = await compressDocxImages(raw, mode);
  cache.set(cacheKey, { mtimeMs: stat.mtimeMs, data: compressed });
  return compressed;
}
