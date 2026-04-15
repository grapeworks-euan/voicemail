// Common provider interface for email + calendar providers (Gmail, Outlook).
// Phase 8 (plan 08-02) — introduced as the abstraction point so Plan 08-03
// can slot in an Outlook implementation without touching Gmail code.
//
// Design note (D-01, 2026-04-14): NO Mail.Send capability is exposed on this
// interface. createDraft / replyDraft return a draft id; users send manually.
// This mirrors the Azure GW2 app scope posture (no Mail.Send permission) and
// honours the post-Softeria-incident no-autonomous-send rule.

export interface EmailMessageRef {
  id: string;
  threadId?: string;
  subject?: string;
  from?: string;
  snippet?: string;
  receivedAt?: Date;
  headers?: Record<string, string>;
}

export interface EmailMessageFull extends EmailMessageRef {
  body?: string;
  bodyHtml?: string;
}

export interface CalendarEventRef {
  id: string;
  summary?: string;
  start?: string; // ISO 8601
  end?: string;
  attendees?: string[];
}

export interface OAuthTokens {
  accessToken: string;
  refreshToken?: string;
  expiresAt?: Date;
  idToken?: string;
  scope?: string;
}

export interface EmailProvider {
  // Mail
  listMessages(query?: { maxResults?: number; q?: string }): Promise<EmailMessageRef[]>;
  getMessage(id: string): Promise<EmailMessageFull>;
  createDraft(args: { to: string; subject: string; body: string; inReplyTo?: string }): Promise<{ id: string }>;
  replyDraft(args: { messageId: string; body: string }): Promise<{ id: string }>;

  // Calendar
  getCalendarEvents(args?: { from?: Date; to?: Date; maxResults?: number }): Promise<CalendarEventRef[]>;
  createCalendarEvent(args: { summary: string; start: Date; end: Date; attendees?: string[]; description?: string }): Promise<{ id: string }>;

  // Auth
  getAuthUrl(state?: string): string;
  exchangeCode(code: string): Promise<OAuthTokens>;
  refreshToken(refreshToken: string): Promise<OAuthTokens>;

  // Unsubscribe
  getUnsubscribeLink(messageId: string): Promise<string | null>;
}
