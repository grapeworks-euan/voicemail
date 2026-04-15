// Factory entry point for the provider abstraction.
//
// VOICEMAIL_EMAIL_PROVIDER selects the active provider:
//   - unset or "gmail"  → GmailProvider (default, preserves Phase 6 behaviour)
//   - "outlook"         → OutlookProvider (Plan 08-03 ships the implementation)
//
// The OutlookProvider file does not exist until Plan 08-03, so the import is
// deferred via require() and only happens when outlook mode is active. This
// keeps the default gmail path fully green while 08-03 is pending.

import type { EmailProvider } from "./EmailProvider";
import { GmailProvider } from "./GmailProvider";

export type {
  EmailProvider,
  EmailMessageRef,
  EmailMessageFull,
  CalendarEventRef,
  OAuthTokens,
} from "./EmailProvider";
export { GmailProvider } from "./GmailProvider";

let cached: EmailProvider | null = null;

export function getProvider(): EmailProvider {
  if (cached) return cached;

  const which = (process.env.VOICEMAIL_EMAIL_PROVIDER || "gmail").toLowerCase();

  if (which === "outlook") {
    // Lazy require — OutlookProvider.ts lands in Plan 08-03. Until then,
    // setting VOICEMAIL_EMAIL_PROVIDER=outlook will throw a clear module-
    // not-found error at runtime, which is the desired behaviour: fail fast
    // rather than silently fall back to Gmail.
    //
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const mod = require("./OutlookProvider") as {
      OutlookProvider: new () => EmailProvider;
    };
    cached = new mod.OutlookProvider();
  } else {
    cached = new GmailProvider();
  }

  return cached;
}

/**
 * Test seam — lets tests inject a specific provider instance (or null to
 * force the next getProvider() call to re-read the env var). Not intended
 * for production use.
 */
export function __setProviderForTest(p: EmailProvider | null): void {
  cached = p;
}
