import { NextRequest, NextResponse } from "next/server";
import { decryptTokens, hasRequiredGoogleScopes } from "@/app/lib/gmail";

export async function GET(request: NextRequest) {
  const cookie = request.cookies.get("gmail_tokens");
  if (!cookie) {
    return NextResponse.json({ authenticated: false });
  }

  try {
    const tokens = decryptTokens(cookie.value);
    return NextResponse.json({ authenticated: hasRequiredGoogleScopes(tokens) });
  } catch {
    return NextResponse.json({ authenticated: false });
  }
}
