import { NextRequest, NextResponse } from "next/server";
import { parseInput } from "@/lib/generate";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof File)) {
      return NextResponse.json({ error: "No file uploaded." }, { status: 400 });
    }

    const buffer = await file.arrayBuffer();
    const { records, detectedFields, sheetName, headerRow } = await parseInput(
      buffer
    );

    return NextResponse.json({
      recordCount: records.length,
      detectedFields,
      sheetName,
      headerRow,
      records,
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unexpected error.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
