// Microsoft Graph implementation of the EmailProvider interface.
//
// Phase 8 Plan 08-03 — ships the Outlook path promised by the provider
// abstraction introduced in Plan 08-02. Uses @azure/msal-node for OAuth
// (authorization-code flow, confidential-client variant — we have a client
// secret from the GW2 Azure app registration) and @microsoft/microsoft-graph-
// client for mail/calendar calls.
//
// DESIGN NOTES (do NOT weaken — these are locked by the phase context):
//   D-01: Drafts-only. No Mail.Send scope anywhere. No /sendMail call. No
//         /me/messages/{id}/send. Draft creation uses POST /me/messages
//         (which Graph treats as a draft until the user hits send in
//         Outlook). This mirrors the post-Softeria-incident posture — even
//         if the Azure app ever gained Mail.Send, this code must not use it.
//   D-05: ConfidentialClientApplication (not PublicClient) because the GW2
//         app registration has a client secret configured; Next.js server
//         routes can safely hold the secret via env vars.
//   D-06: Env var names are the VOICEMAIL_OUTLOOK_* set provisioned by
//         Plan 08-01.
//   D-08: Unsubscribe link extraction reads `internetMessageHeaders` on the
//         Graph message resource (Gmail uses payload.headers[]; Graph's
//         shape is different).

import { ConfidentialClientApplication, type Configuration } from "@azure/msal-node";
import { Client, type AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import type {
  EmailProvider,
  EmailMessageRef,
  EmailMessageFull,
  CalendarEventRef,
  OAuthTokens,
} from "./EmailProvider";

// LOCKED per D-01. NEVER add 'Mail.Send' to this list — drafts-only posture.
const GRAPH_SCOPES = [
  "Mail.ReadWrite",
  "Mail.Read",
  "Calendars.ReadWrite",
  "User.Read",
  "offline_access",
];

function env(name: string): string {
  const v = process.env[name];
  if (!v) {
    throw new Error(`OutlookProvider: missing env var ${name}`);
  }
  return v;
}

// Shapes we rely on from Graph responses. Kept local to avoid pulling in the
// full Graph types package (which would pull in a lot of surface for little
// benefit).
type GraphEmailAddress = { address?: string; name?: string };
type GraphRecipient = { emailAddress?: GraphEmailAddress };
type GraphMessage = {
  id: string;
  conversationId?: string;
  subject?: string;
  from?: GraphRecipient;
  bodyPreview?: string;
  receivedDateTime?: string;
  body?: { contentType?: string; content?: string };
  internetMessageHeaders?: Array<{ name: string; value: string }>;
};
type GraphEvent = {
  id: string;
  subject?: string;
  start?: { dateTime: string; timeZone?: string };
  end?: { dateTime: string; timeZone?: string };
  attendees?: Array<{ emailAddress?: GraphEmailAddress; type?: string }>;
};
type GraphList<T> = { value: T[] };

export class OutlookProvider implements EmailProvider {
  private msal: ConfidentialClientApplication;
  private redirectUri: string;
  // Token-per-request state. The OAuth callback route calls exchangeCode and
  // persists tokens via session.ts; downstream callers obtain this provider
  // via getProvider() and call setAccessTokenForRequest(token) before using
  // mail/calendar methods (Plan 08-04 wires that layer end-to-end).
  private accessTokenForCurrentRequest: string | null = null;

  constructor() {
    const config: Configuration = {
      auth: {
        clientId: env("VOICEMAIL_OUTLOOK_CLIENT_ID"),
        clientSecret: env("VOICEMAIL_OUTLOOK_CLIENT_SECRET"),
        authority: `https://login.microsoftonline.com/${env("VOICEMAIL_OUTLOOK_TENANT_ID")}`,
      },
    };
    this.msal = new ConfidentialClientApplication(config);
    this.redirectUri = env("VOICEMAIL_OUTLOOK_REDIRECT_URI");
  }

  // ----- Auth -----

  /**
   * Build the Microsoft identity platform authorization URL. The
   * EmailProvider interface requires a synchronous return, and MSAL's
   * `getAuthCodeUrl` is async, so we construct the URL manually. This
   * mirrors the shape `getAuthCodeUrl` produces — verified against the
   * v2.0 /authorize endpoint documentation.
   */
  getAuthUrl(state?: string): string {
    const params = new URLSearchParams({
      client_id: env("VOICEMAIL_OUTLOOK_CLIENT_ID"),
      response_type: "code",
      redirect_uri: this.redirectUri,
      response_mode: "query",
      scope: GRAPH_SCOPES.join(" "),
      state: state ?? "",
    });
    return `https://login.microsoftonline.com/${env("VOICEMAIL_OUTLOOK_TENANT_ID")}/oauth2/v2.0/authorize?${params.toString()}`;
  }

  async exchangeCode(code: string): Promise<OAuthTokens> {
    const result = await this.msal.acquireTokenByCode({
      code,
      scopes: GRAPH_SCOPES,
      redirectUri: this.redirectUri,
    });
    if (!result) {
      throw new Error("OutlookProvider.exchangeCode: MSAL returned null");
    }
    return {
      accessToken: result.accessToken,
      // MSAL's typed result doesn't expose the refresh_token publicly, but
      // it's present on the raw response. offline_access scope is required
      // to get it (we request it in GRAPH_SCOPES).
      refreshToken: (result as unknown as { refreshToken?: string }).refreshToken,
      expiresAt: result.expiresOn ?? undefined,
      idToken: result.idToken ?? undefined,
      scope: result.scopes?.join(" "),
    };
  }

  async refreshToken(refreshTokenValue: string): Promise<OAuthTokens> {
    const result = await this.msal.acquireTokenByRefreshToken({
      refreshToken: refreshTokenValue,
      scopes: GRAPH_SCOPES,
    });
    if (!result) {
      throw new Error("OutlookProvider.refreshToken: MSAL returned null");
    }
    return {
      accessToken: result.accessToken,
      refreshToken: (result as unknown as { refreshToken?: string }).refreshToken,
      expiresAt: result.expiresOn ?? undefined,
      idToken: result.idToken ?? undefined,
      scope: result.scopes?.join(" "),
    };
  }

  // ----- Token state seam -----

  /**
   * Called by the callback route (after exchangeCode persists tokens) and
   * by downstream handlers that have already looked up stored tokens for
   * the current user. Plan 08-04 wires this into the handler layer.
   */
  setAccessTokenForRequest(token: string): void {
    this.accessTokenForCurrentRequest = token;
  }

  private requireToken(): string {
    if (!this.accessTokenForCurrentRequest) {
      throw new Error(
        "OutlookProvider: no access token set. Call setAccessTokenForRequest(token) before invoking provider methods."
      );
    }
    return this.accessTokenForCurrentRequest;
  }

  private graph(accessToken: string): Client {
    const authProvider: AuthenticationProvider = {
      getAccessToken: async () => accessToken,
    };
    return Client.initWithMiddleware({ authProvider });
  }

  // ----- Mail -----

  async listMessages(query?: { maxResults?: number; q?: string }): Promise<EmailMessageRef[]> {
    const client = this.graph(this.requireToken());
    const top = query?.maxResults ?? 25;
    let req = client
      .api("/me/messages")
      .top(top)
      .select("id,subject,from,receivedDateTime,bodyPreview,conversationId");
    if (query?.q) {
      req = req.search(`"${query.q}"`);
    }
    const res = (await req.get()) as GraphList<GraphMessage>;
    return (res.value ?? []).map((m) => ({
      id: m.id,
      threadId: m.conversationId,
      subject: m.subject,
      from: m.from?.emailAddress?.address,
      snippet: m.bodyPreview,
      receivedAt: m.receivedDateTime ? new Date(m.receivedDateTime) : undefined,
    }));
  }

  async getMessage(id: string): Promise<EmailMessageFull> {
    const client = this.graph(this.requireToken());
    const msg = (await client
      .api(`/me/messages/${encodeURIComponent(id)}`)
      .select("id,subject,from,receivedDateTime,body,internetMessageHeaders,conversationId")
      .get()) as GraphMessage;

    const headers: Record<string, string> = {};
    for (const h of msg.internetMessageHeaders ?? []) {
      // Graph returns header names in their original case (e.g. "List-Unsubscribe").
      // We normalise to lowercase so callers can look up headers consistently.
      headers[h.name.toLowerCase()] = h.value;
    }

    const bodyText = msg.body?.contentType?.toLowerCase() === "text" ? msg.body?.content : undefined;
    const bodyHtml = msg.body?.contentType?.toLowerCase() === "html" ? msg.body?.content : undefined;

    return {
      id: msg.id,
      threadId: msg.conversationId,
      subject: msg.subject,
      from: msg.from?.emailAddress?.address,
      receivedAt: msg.receivedDateTime ? new Date(msg.receivedDateTime) : undefined,
      body: bodyText,
      bodyHtml,
      headers,
    };
  }

  async createDraft(args: {
    to: string;
    subject: string;
    body: string;
    inReplyTo?: string;
  }): Promise<{ id: string }> {
    // D-01: POST /me/messages creates a DRAFT in the user's Drafts folder.
    // It does NOT send. We never call /me/sendMail and never POST to
    // /me/messages/{id}/send. The user sends manually from Outlook.
    const client = this.graph(this.requireToken());
    const draft = (await client.api("/me/messages").post({
      subject: args.subject,
      body: { contentType: "Text", content: args.body },
      toRecipients: [{ emailAddress: { address: args.to } }],
    })) as GraphMessage;
    return { id: draft.id };
  }

  async replyDraft(args: { messageId: string; body: string }): Promise<{ id: string }> {
    // createReply produces a draft reply (already populated with recipients
    // and threading headers). We PATCH the body onto it. Still a draft —
    // never sent by this code.
    const client = this.graph(this.requireToken());
    const replyDraft = (await client
      .api(`/me/messages/${encodeURIComponent(args.messageId)}/createReply`)
      .post({})) as GraphMessage;
    await client.api(`/me/messages/${encodeURIComponent(replyDraft.id)}`).patch({
      body: { contentType: "Text", content: args.body },
    });
    return { id: replyDraft.id };
  }

  // ----- Calendar -----

  async getCalendarEvents(args?: {
    from?: Date;
    to?: Date;
    maxResults?: number;
  }): Promise<CalendarEventRef[]> {
    const client = this.graph(this.requireToken());
    const top = args?.maxResults ?? 25;
    let req;
    if (args?.from && args?.to) {
      // Use /me/calendarView for range queries — it correctly expands
      // recurring events into their instances.
      req = client
        .api("/me/calendarView")
        .query({
          startDateTime: args.from.toISOString(),
          endDateTime: args.to.toISOString(),
        })
        .top(top)
        .select("id,subject,start,end,attendees");
    } else {
      req = client.api("/me/events").top(top).select("id,subject,start,end,attendees");
    }
    const res = (await req.get()) as GraphList<GraphEvent>;
    return (res.value ?? []).map((e) => ({
      id: e.id,
      summary: e.subject,
      start: e.start?.dateTime,
      end: e.end?.dateTime,
      attendees: e.attendees
        ?.map((a) => a.emailAddress?.address ?? "")
        .filter((a) => a.length > 0),
    }));
  }

  async createCalendarEvent(args: {
    summary: string;
    start: Date;
    end: Date;
    attendees?: string[];
    description?: string;
  }): Promise<{ id: string }> {
    const client = this.graph(this.requireToken());
    const ev = (await client.api("/me/events").post({
      subject: args.summary,
      body: args.description
        ? { contentType: "Text", content: args.description }
        : undefined,
      start: { dateTime: args.start.toISOString(), timeZone: "UTC" },
      end: { dateTime: args.end.toISOString(), timeZone: "UTC" },
      attendees: (args.attendees ?? []).map((addr) => ({
        emailAddress: { address: addr },
        type: "required",
      })),
    })) as GraphEvent;
    return { id: ev.id };
  }

  // ----- Unsubscribe (D-08) -----

  /**
   * Parse the RFC 2369 List-Unsubscribe header from the message's
   * `internetMessageHeaders` array. Graph returns headers as [{name, value}]
   * rather than Gmail's payload.headers[] shape, so this implementation is
   * per-provider.
   *
   * We prefer HTTPS URLs over mailto (and never return http). Returns null
   * when the header is absent or contains no usable value.
   */
  async getUnsubscribeLink(messageId: string): Promise<string | null> {
    const full = await this.getMessage(messageId);
    const raw = full.headers?.["list-unsubscribe"];
    if (!raw) return null;
    // RFC 2369: comma-separated angle-bracketed values, e.g.
    //   <https://example.com/unsub?x=1>, <mailto:unsub@example.com>
    const matches = raw.match(/<([^>]+)>/g) ?? [];
    const urls = matches.map((m) => m.slice(1, -1));
    const https = urls.find((u) => u.toLowerCase().startsWith("https://"));
    if (https) return https;
    const mailto = urls.find((u) => u.toLowerCase().startsWith("mailto:"));
    if (mailto) return mailto;
    return urls[0] ?? null;
  }
}
