import { NextResponse } from "next/server";

export const runtime = "nodejs";

const GMAIL_OAUTH_VERIFIER_COOKIE = "event_based_reminders_app_gmail_oauth_verifier";
const GMAIL_SCOPES = [
  "openid",
  "email",
  "profile",
  "https://www.googleapis.com/auth/gmail.compose",
  "https://www.googleapis.com/auth/calendar.events",
] as const;
const GMAIL_LOCAL_REDIRECT_URI = "http://localhost:8664/api/auth/google/callback";
const GMAIL_HOSTED_REDIRECT_URI = "https://event-based-reminders-app.vercel.app/api/auth/google/callback";

function getGoogleClientId() {
  return String(process.env.GOOGLE_CLIENT_ID ?? process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID ?? "").trim();
}

function getGoogleClientSecret() {
  return String(process.env.GOOGLE_CLIENT_SECRET ?? "").trim();
}

function getGoogleRedirectUri(url: URL) {
  return url.hostname === "localhost" ? GMAIL_LOCAL_REDIRECT_URI : GMAIL_HOSTED_REDIRECT_URI;
}

function getCookieValue(request: Request, name: string) {
  const cookieHeader = request.headers.get("cookie");
  if (!cookieHeader) return "";

  for (const part of cookieHeader.split(";")) {
    const [rawName, ...rest] = part.trim().split("=");
    if (rawName === name) {
      return decodeURIComponent(rest.join("="));
    }
  }

  return "";
}

export async function GET(request: Request) {
  const url = new URL(request.url);
  const code = url.searchParams.get("code");
  const state = url.searchParams.get("state");
  const error = url.searchParams.get("error");
  const errorDescription = url.searchParams.get("error_description");
  const origin = `${url.protocol}//${url.host}`;
  const redirectUri = getGoogleRedirectUri(url);
  const clientId = getGoogleClientId();
  const clientSecret = getGoogleClientSecret();
  const codeVerifier = getCookieValue(request, GMAIL_OAUTH_VERIFIER_COOKIE);

  let accessToken: string | null = null;
  let refreshToken: string | null = null;
  let expiresIn: number | null = null;
  let scope: string | null = null;
  let resolvedError = error;
  let resolvedErrorDescription = errorDescription;

  if (!resolvedError && code) {
    if (!clientId || !clientSecret) {
      resolvedError = "server_misconfigured";
      resolvedErrorDescription = "Gmail OAuth is missing GOOGLE_CLIENT_ID or GOOGLE_CLIENT_SECRET on the server.";
    } else if (!codeVerifier) {
      resolvedError = "missing_code_verifier";
      resolvedErrorDescription = "Gmail sign-in could not be verified. Please try again.";
    } else {
      const body = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: "authorization_code",
        code,
        redirect_uri: redirectUri,
        code_verifier: codeVerifier,
      });

      const tokenResponse = await fetch("https://oauth2.googleapis.com/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
        cache: "no-store",
      });

      const tokenPayload = (await tokenResponse.json().catch(() => null)) as
        | {
            access_token?: string;
            refresh_token?: string;
            expires_in?: number;
            scope?: string;
            error?: string;
            error_description?: string;
          }
        | null;

      if (!tokenResponse.ok || !tokenPayload?.access_token) {
        resolvedError = tokenPayload?.error || "token_exchange_failed";
        resolvedErrorDescription = tokenPayload?.error_description || "Failed to connect Gmail.";
      } else {
        accessToken = tokenPayload.access_token;
        refreshToken = tokenPayload.refresh_token ?? null;
        expiresIn = tokenPayload.expires_in ?? null;
        scope = tokenPayload.scope ?? GMAIL_SCOPES.join(" ");
      }
    }
  }

  const html = `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <title>Connecting Gmail…</title>
  </head>
  <body>
    <script>
      (function () {
        const payload = {
          type: "event_based_reminders_app_gmail_oauth_result",
          accessToken: ${JSON.stringify(accessToken)},
          refreshToken: ${JSON.stringify(refreshToken)},
          expiresIn: ${JSON.stringify(expiresIn)},
          scope: ${JSON.stringify(scope)},
          state: ${JSON.stringify(state)},
          error: ${JSON.stringify(resolvedError)},
          errorDescription: ${JSON.stringify(resolvedErrorDescription)},
        };
        if (window.opener) {
          window.opener.postMessage(payload, ${JSON.stringify(origin)});
          window.close();
        }
      })();
    </script>
    <p>Connecting Gmail… You can close this window if it does not close automatically.</p>
  </body>
</html>`;

  return new NextResponse(html, {
    headers: {
      "Content-Type": "text/html; charset=utf-8",
      "Cache-Control": "no-store",
      "Set-Cookie": `${GMAIL_OAUTH_VERIFIER_COOKIE}=; Path=/; Max-Age=0; SameSite=Lax; Secure`,
    },
  });
}
