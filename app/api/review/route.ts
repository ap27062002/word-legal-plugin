import { NextRequest, NextResponse } from "next/server";
import { reviewDocument } from "@/lib/agent";

export async function POST(request: NextRequest) {
  const body = await request.json().catch(() => null);
  const documentText = body?.documentText;

  if (typeof documentText !== "string" || documentText.trim().length === 0) {
    return NextResponse.json({ error: "documentText is required" }, { status: 400 });
  }

  const findings = await reviewDocument(documentText);
  return NextResponse.json({ findings });
}
