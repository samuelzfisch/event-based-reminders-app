import { NextResponse } from "next/server";

export const runtime = "nodejs";

const OUTLOOK_OAUTH_VERIFIER_COOKIE = "event_based_reminders_app_outlook_oauth_verifier";
const OUTLOOK_SCOPES = ["User.Read", "offline_access", "Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite"] as const;
const OUTLOOK_LOCAL_REDIRECT_URI = "http://localhost:8664/api/auth/microsoft/callback";
const OUTLOOK_HOSTED_REDIRECT_URI = "https://event-based-reminders-app.vercel.app/api/auth/microsoft/callback";

function getMicrosoftClientId() {
  return String(process.env.MICROSOFT_CLIENT_ID ?? process.env.NEXT_PUBLIC_MICROSOFT_CLIENT_ID ?? "").trim();
}

function getMicrosoftClientSecret() {
  return String(process.env.MICROSOFT_CLIENT_SECRET ?? "").trim();
}

function getMicrosoftTenantId() {
  return "common";
}

function getMicrosoftRedirectUri(url: URL) {
  return url.hostname === "localhost" ? OUTLOOK_LOCAL_REDIRECT_URI : OUTLOOK_HOSTED_REDIRECT_URI;
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
  const redirectUri = getMicrosoftRedirectUri(url);
  const clientId = getMicrosoftClientId();
  const clientSecret = getMicrosoftClientSecret();
  const codeVerifier = getCookieValue(request, OUTLOOK_OAUTH_VERIFIER_COOKIE);
  const clientSecretDiagnostic = {
    hasClientSecret: Boolean(clientSecret),
    clientSecretLength: clientSecret.length,
  };

  console.info("[Microsoft OAuth callback] server env diagnostic", clientSecretDiagnostic);

  let accessToken: string | null = null;
  let refreshToken: string | null = null;
  let expiresIn: number | null = null;
  let scope: string | null = null;
  let resolvedError = error;
  let resolvedErrorDescription = errorDescription;

  if (!resolvedError && code) {
    if (!clientId || !clientSecret) {
      resolvedError = "server_misconfigured";
      resolvedErrorDescription =
        `Microsoft OAuth is missing MICROSOFT_CLIENT_ID or MICROSOFT_CLIENT_SECRET on the server. ${JSON.stringify(clientSecretDiagnostic)}`;
    } else if (!codeVerifier) {
      resolvedError = "missing_code_verifier";
      resolvedErrorDescription = "Outlook sign-in could not be verified. Please try again.";
    } else {
      const tokenUrl = `https://login.microsoftonline.com/${getMicrosoftTenantId()}/oauth2/v2.0/token`;
      const body = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: "authorization_code",
        code,
        redirect_uri: redirectUri,
        code_verifier: codeVerifier,
      });

      console.info("[Microsoft OAuth callback] token request", {
        tokenUrl,
        redirectUri,
        hasClientId: Boolean(clientId),
        hasClientSecret: Boolean(clientSecret),
        clientIdLength: clientId.length,
        clientSecretLength: clientSecret.length,
        contentType: "application/x-www-form-urlencoded",
        bodyType: typeof body.toString(),
        bodyKeys: Array.from(body.keys()),
        currentHost: url.host,
        requestUrl: request.url,
      });

      const tokenResponse = await fetch(
        tokenUrl,
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: body.toString(),
          cache: "no-store",
        }
      );

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
        console.error("[Microsoft OAuth callback] token response failure", {
          status: tokenResponse.status,
          statusText: tokenResponse.statusText,
          azureErrorBody: tokenPayload
            ? {
                error: tokenPayload.error ?? null,
                error_description: tokenPayload.error_description ?? null,
              }
            : null,
        });
        resolvedError = tokenPayload?.error || "token_exchange_failed";
        resolvedErrorDescription = tokenPayload?.error_description || "Failed to connect Outlook.";
      } else {
        accessToken = tokenPayload.access_token;
        refreshToken = tokenPayload.refresh_token ?? null;
        expiresIn = tokenPayload.expires_in ?? null;
        scope = tokenPayload.scope ?? null;
      }
    }
  }

  const html = `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <title>Connecting Outlook…</title>
  </head>
  <body>
    <script>
      (function () {
        const payload = {
          type: "event_based_reminders_app_outlook_oauth_result",
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
    <p>Connecting Outlook… You can close this window if it does not close automatically.</p>
  </body>
</html>`;

  return new NextResponse(html, {
    headers: {
      "Content-Type": "text/html; charset=utf-8",
      "Cache-Control": "no-store",
      "Set-Cookie": `${OUTLOOK_OAUTH_VERIFIER_COOKIE}=; Path=/; Max-Age=0; SameSite=Lax; Secure`,
    },
  });
}
