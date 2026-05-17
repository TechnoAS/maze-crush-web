import { NextRequest, NextResponse } from "next/server";
import { parseReport, toCSVRow } from "@/lib/parser";

export async function POST(req: NextRequest) {
  try {
    const { text } = await req.json();
    if (!text || typeof text !== "string" || text.trim().length === 0) {
      return NextResponse.json({ error: "No report text provided." }, { status: 400 });
    }
    const data = parseReport(text);
    const csv = toCSVRow(data);
    return NextResponse.json({ data, csv });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
