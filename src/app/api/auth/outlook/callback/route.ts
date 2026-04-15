// Next.js route handler for the Microsoft identity platform redirect.
//
// Phase 8 Plan 08-03 (D-07). Lives as a sibling to the Google callback at
// src/app/api/auth/callback/route.ts. Mirrors that route's shape:
//   1. Read ?code (and ?error) from the query.
//   2. Hand the code to OutlookProvider.exchangeCode() → MSAL returns
//      access + refresh + id tokens.
//   3. Fetch the signed-in user's email via Graph /me (OutlookProvider holds
//      the token; we set it on the provider for this single Graph call).
//   4. Persist via the same db.ts + session.ts helpers the Google callback
//      uses, writing into the existing users row. Plan 08-04 extends the
//      schema with an outlook_accounts table analogous to google_accounts;
//      until then we use the existing user row and session cookie so this
//      route compiles and end-to-end tests can run without schema changes.
//   5. Set the encrypted session cookie and redirect to /app, exactly like
//      the Google callback.
//
// Drafts-only posture (D-01): this route never requests Mail.Send — the
// scopes are set in OutlookProvider.getAuthUrl(). This file never calls
// /me/sendMail or any Graph send endpoint.

import { NextRequest, NextResponse } from "next/server";
import { Client } from "@microsoft/microsoft-graph-client";
import { OutlookProvider } from "@/app/lib/providers/OutlookProvider";
import {
  SESSION_COOKIE_NAME,
  SESSION_MAX_AGE,
  createSessionCookieValue,
} from "@/app/lib/session";
import { initDb, upsertUser } from "@/app/lib/db";

type GraphMe = {
  id?: string;
  mail?: string | null;
  userPrincipalName?: string | null;
  displayName?: string | null;
};

/**
 * Fetch the signed-in user's primary email via Graph /me. Graph sometimes
 * returns `mail` as null (when the account has no primary SMTP address
 * provisioned) — fall back to `userPrincipalName`, which is the account's
 * login and is always present.
 */
async function fetchGraphEmail(accessToken: string): Promise<string> {
  const client = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => accessToken,
    },
  });
  const me = (await client.api("/me").select("id,mail,userPrincipalName,displayName").get()) as GraphMe;
  const email = me.mail ?? me.userPrincipalName ?? null;
  if (!email) {
    throw new Error("Graph /me returned no mail or userPrincipalName");
  }
  return email;
}

export async function GET(request: NextRequest) {
  const code = request.nextUrl.searchParams.get("code");
  const error = request.nextUrl.searchParams.get("error");
  const errorDescription = request.nextUrl.searchParams.get("error_description");

  const host =
    request.headers.get("x-forwarded-host") ||
    request.headers.get("host") ||
    "";
  const proto = request.headers.get("x-forwarded-proto") || "https";
  const origin = host ? `${proto}://${host}` : request.url;

  if (error) {
    console.error(
      `outlook_auth_callback: provider error=${error} description=${errorDescription}`
    );
    const msg = errorDescription || error;
    return NextResponse.redirect(
      new URL(`/?auth_error=${encodeURIComponent(msg)}`, origin)
    );
  }
  if (!code) {
    return NextResponse.json({ error: "No code provided" }, { status: 400 });
  }

  try {
    const provider = new OutlookProvider();
    const tokens = await provider.exchangeCode(code);

    // Fetch signed-in user's email via Graph /me so we can key the app's
    // user record. Mirror the google-auth flow which calls getUserEmail().
    const email = await fetchGraphEmail(tokens.accessToken);

    await initDb();
    const user = await upsertUser(email);
    if (!user) {
      return NextResponse.json({ error: "Database error" }, { status: 500 });
    }

    // TODO (Plan 08-04): persist tokens in a dedicated outlook_accounts
    // table analogous to google_accounts (addOutlookAccount helper). For
    // Plan 08-03 the acceptance criteria are "route exists, type-checks,
    // completes MSAL code exchange, sets session" — which this does. The
    // tokens live in memory for this single request; Plan 08-04 wires
    // persistence + a refresh-on-demand path.
    void tokens;

    const response = NextResponse.redirect(new URL("/app", origin));
    response.cookies.set(
      SESSION_COOKIE_NAME,
      createSessionCookieValue(user.id),
      {
        httpOnly: true,
        secure: process.env.NODE_ENV === "production",
        sameSite: "lax",
        maxAge: SESSION_MAX_AGE,
        path: "/",
      }
    );

    console.log(`outlook_auth_callback: login email=${email}`);
    return response;
  } catch (err) {
    console.error("Outlook OAuth callback error:", err);
    const msg = err instanceof Error ? err.message : String(err);
    return NextResponse.redirect(
      new URL(`/?auth_error=${encodeURIComponent(msg)}`, origin)
    );
  }
}
