import { NextRequest, NextResponse } from "next/server";
import { generateLetters } from "@/lib/docxMerge";
import { getLetterTemplate } from "@/lib/letterTemplates";
import type { CleanRecord } from "@/lib/fields";
import type { TemplateCompression } from "@/lib/compressDocx";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: NextRequest) {
  try {
    const body = (await request.json()) as {
      records?: CleanRecord[];
      fileName?: string;
      templateId?: string;
      format?: "docx" | "pdf";
      compression?: TemplateCompression;
    };

    const records = body.records;
    if (!Array.isArray(records) || records.length === 0) {
      return NextResponse.json(
        { error: "No records provided to generate letters." },
        { status: 400 }
      );
    }

    const tpl = await getLetterTemplate(body.templateId);
    if (!tpl) {
      return NextResponse.json(
        { error: "No letter template is available in the templates folder." },
        { status: 400 }
      );
    }

    const format = body.format === "pdf" ? "pdf" : "docx";
    const compression =
      body.compression === "none" ||
      body.compression === "balanced" ||
      body.compression === "small"
        ? body.compression
        : undefined;

    const zipBuffer = await generateLetters(records, tpl.id, format, compression);

    const base = (body.fileName || "output").replace(/\.[^.]+$/, "");
    const outName = `${base} - ${tpl.label} Letters (${format.toUpperCase()}).zip`;

    return new NextResponse(new Uint8Array(zipBuffer), {
      status: 200,
      headers: {
        "Content-Type": "application/zip",
        "Content-Disposition": `attachment; filename="${encodeURIComponent(
          outName
        )}"`,
        "X-Record-Count": String(records.length),
        "X-Template-Label": encodeURIComponent(tpl.label),
        "X-Format": format,
      },
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unexpected error.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
