import { google } from "googleapis";

export function getOAuth2Client(redirectUri?: string) {
  return new google.auth.OAuth2(
    process.env.VOICEMAIL_GOOGLE_CLIENT_ID,
    process.env.VOICEMAIL_GOOGLE_CLIENT_SECRET,
    redirectUri || process.env.VOICEMAIL_GOOGLE_REDIRECT_URI
  );
}

export function getAuthedClient(tokens: any) {
  const client = getOAuth2Client();
  client.setCredentials(tokens);
  return client;
}

export function getGmailClient(tokens: any) {
  return google.gmail({ version: "v1", auth: getAuthedClient(tokens) });
}

export function getCalendarClient(tokens: any) {
  return google.calendar({ version: "v3", auth: getAuthedClient(tokens) });
}
