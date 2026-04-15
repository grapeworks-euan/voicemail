// Thin wrapper that adapts the existing Google implementation
// (gmail.ts / calendar-api.ts / google-auth.ts / unsubscribe.ts) to the
// shared EmailProvider interface.
//
// Zero behavioural change — every method delegates to the existing helper
// that Phase 6 already shipped. The only wrinkle is that the original Google
// code is token-parameterised (tokens passed to every call), whereas the
// provider abstraction is stateful: callers obtain a provider instance via
// getProvider() and feed tokens once via setTokens(). Plan 08-04 wires
// callers up to this new shape; until then GmailProvider exists only to
// prove the abstraction compiles and so Plan 08-03 has an interface to
// implement against.
//
// If a caller invokes a method before setTokens(), we throw rather than
// silently calling Google with undefined credentials.

import * as gmail from "../gmail";
import { getAuthUrl as googleGetAuthUrl, exchangeCode as googleExchangeCode } from "../gmail";
import { getOAuth2Client } from "../google-auth";
import {
  listCalendarEvents,
  createCalendarInvite,
  type CalendarEventSummary,
} from "../calendar-api";
import { getUnsubscribeInfo } from "../unsubscribe";
import type {
  EmailProvider,
  EmailMessageRef,
  EmailMessageFull,
  CalendarEventRef,
  OAuthTokens,
} from "./EmailProvider";

export class GmailProvider implements EmailProvider {
  private tokens: unknown | null = null;

  /**
   * Gmail's existing helpers take tokens as an argument on every call, but
   * the EmailProvider interface is stateful. Callers should set tokens once
   * (typically right after reading them from the session) and then use the
   * provider normally.
   */
  setTokens(tokens: unknown): void {
    this.tokens = tokens;
  }

  private requireTokens(): unknown {
    if (!this.tokens) {
      throw new Error(
        "GmailProvider: no tokens set. Call setTokens(tokens) before invoking provider methods."
      );
    }
    return this.tokens;
  }

  // ---------- Mail ----------

  async listMessages(query?: { maxResults?: number; q?: string }): Promise<EmailMessageRef[]> {
    const tokens = this.requireTokens();
    const maxResults = query?.maxResults ?? 10;

    if (query?.q) {
      const emails = await gmail.searchEmails(tokens, query.q, maxResults);
      return emails.map((e) => ({
        id: e.id,
        threadId: e.threadId,
        subject: e.subject,
        from: e.from,
        snippet: e.snippet,
        receivedAt: e.date ? new Date(e.date) : undefined,
      }));
    }

    const { emails } = await gmail.getUnreadEmails(tokens, maxResults);
    return emails.map((e) => ({
      id: e.id,
      threadId: e.threadId,
      subject: e.subject,
      from: e.from,
      snippet: e.snippet,
      receivedAt: e.date ? new Date(e.date) : undefined,
    }));
  }

  async getMessage(id: string): Promise<EmailMessageFull> {
    const tokens = this.requireTokens();
    const body = await gmail.getEmailBody(tokens, id);
    // The existing helper returns body text only; metadata (subject/from)
    // isn't surfaced by a single call. For the MVP abstraction we return
    // the body and the id — Plan 08-04 (or OutlookProvider parity work)
    // can extend this to fetch headers if needed.
    return {
      id,
      body,
    };
  }

  async createDraft(args: {
    to: string;
    subject: string;
    body: string;
    inReplyTo?: string;
  }): Promise<{ id: string }> {
    // Gmail's existing surface (sendNewEmail) sends immediately. The
    // EmailProvider contract is drafts-only (D-01). A Gmail-side drafts
    // API implementation lives in Plan 08-04; until then the method exists
    // to satisfy the interface so the build is green.
    this.requireTokens();
    void args;
    throw new Error(
      "GmailProvider.createDraft: not yet implemented. Plan 08-04 adds the Gmail drafts path; use the direct gmail.ts helpers in the meantime."
    );
  }

  async replyDraft(args: { messageId: string; body: string }): Promise<{ id: string }> {
    // Same reasoning as createDraft. Existing gmail.sendReply sends
    // immediately; the drafts-equivalent lives in Plan 08-04.
    this.requireTokens();
    void args;
    throw new Error(
      "GmailProvider.replyDraft: not yet implemented. Plan 08-04 adds the Gmail drafts path; use gmail.sendReply directly in the meantime."
    );
  }

  // ---------- Calendar ----------

  async getCalendarEvents(args?: {
    from?: Date;
    to?: Date;
    maxResults?: number;
  }): Promise<CalendarEventRef[]> {
    const tokens = this.requireTokens();
    const events: CalendarEventSummary[] = await listCalendarEvents(tokens, {
      startTime: args?.from?.toISOString(),
      endTime: args?.to?.toISOString(),
      maxResults: args?.maxResults,
    });
    return events.map((e) => ({
      id: e.id,
      summary: e.summary,
      start: e.start,
      end: e.end,
      attendees: e.attendees,
    }));
  }

  async createCalendarEvent(args: {
    summary: string;
    start: Date;
    end: Date;
    attendees?: string[];
    description?: string;
  }): Promise<{ id: string }> {
    const tokens = this.requireTokens();
    const { event } = await createCalendarInvite(tokens, {
      title: args.summary,
      startTime: args.start.toISOString(),
      endTime: args.end.toISOString(),
      attendeeEmails: args.attendees,
      notes: args.description,
      locationPreference: "none",
    });
    return { id: event.id };
  }

  // ---------- Auth ----------

  getAuthUrl(state?: string): string {
    return googleGetAuthUrl(undefined, state ? { state } : undefined);
  }

  async exchangeCode(code: string): Promise<OAuthTokens> {
    const raw = await googleExchangeCode(code);
    return this.normalizeGoogleTokens(raw);
  }

  async refreshToken(refreshTokenValue: string): Promise<OAuthTokens> {
    // google-auth.ts doesn't expose a direct refresh helper — the googleapis
    // OAuth2 client handles refresh transparently when credentials are set.
    // Here we do an explicit refresh so callers can rotate stored tokens.
    const client = getOAuth2Client();
    client.setCredentials({ refresh_token: refreshTokenValue });
    const { credentials } = await client.refreshAccessToken();
    return this.normalizeGoogleTokens(credentials);
  }

  private normalizeGoogleTokens(raw: {
    access_token?: string | null;
    refresh_token?: string | null;
    expiry_date?: number | null;
    id_token?: string | null;
    scope?: string | null;
  }): OAuthTokens {
    if (!raw.access_token) {
      throw new Error("Google OAuth response did not include an access_token");
    }
    return {
      accessToken: raw.access_token,
      refreshToken: raw.refresh_token ?? undefined,
      expiresAt: raw.expiry_date ? new Date(raw.expiry_date) : undefined,
      idToken: raw.id_token ?? undefined,
      scope: raw.scope ?? undefined,
    };
  }

  // ---------- Unsubscribe ----------

  async getUnsubscribeLink(messageId: string): Promise<string | null> {
    const tokens = this.requireTokens();
    const info = await getUnsubscribeInfo(tokens, messageId);
    // Prefer an HTTPS list-unsubscribe URL, fall back to mailto, then body.
    return (
      info.httpsUrls[0] ??
      info.mailtoUrls[0] ??
      info.bodyLinks[0] ??
      null
    );
  }
}
