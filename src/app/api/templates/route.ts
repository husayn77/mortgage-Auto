import { NextResponse } from "next/server";
import { listLetterTemplates } from "@/lib/letterTemplates";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET() {
  try {
    const templates = await listLetterTemplates();
    return NextResponse.json({
      templates: templates.map((t) => ({
        id: t.id,
        label: t.label,
        suffix: t.suffix,
      })),
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unexpected error.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
