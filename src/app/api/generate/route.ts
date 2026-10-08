import { NextRequest, NextResponse } from "next/server";
import { generateWorkbook } from "@/lib/generate";
import type { CleanRecord } from "@/lib/fields";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: NextRequest) {
  try {
    const body = (await request.json()) as {
      records?: CleanRecord[];
      fileName?: string;
    };

    const records = body.records;
    if (!Array.isArray(records) || records.length === 0) {
      return NextResponse.json(
        { error: "No records provided to export." },
        { status: 400 }
      );
    }

    const outBuffer = await generateWorkbook(records);

    const base = (body.fileName || "output").replace(/\.[^.]+$/, "");
    const outName = `${base} - MGT.xlsx`;

    return new NextResponse(new Uint8Array(outBuffer), {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${encodeURIComponent(
          outName
        )}"`,
        "X-Record-Count": String(records.length),
      },
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unexpected error.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
