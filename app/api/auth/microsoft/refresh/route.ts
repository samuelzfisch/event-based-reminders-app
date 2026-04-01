import { NextResponse } from "next/server";

const OUTLOOK_SCOPES = ["User.Read", "offline_access", "Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite"] as const;

function getMicrosoftClientId() {
  return String(process.env.MICROSOFT_CLIENT_ID ?? process.env.NEXT_PUBLIC_MICROSOFT_CLIENT_ID ?? "").trim();
}

function getMicrosoftClientSecret() {
  return String(process.env.MICROSOFT_CLIENT_SECRET ?? "").trim();
}

function getMicrosoftTenantId() {
  return "common";
}

export async function POST(request: Request) {
  const { refreshToken, redirectUri } = (await request.json().catch(() => ({}))) as {
    refreshToken?: string;
    redirectUri?: string;
  };

  const clientId = getMicrosoftClientId();
  const clientSecret = getMicrosoftClientSecret();
  const safeRefreshToken = String(refreshToken ?? "").trim();
  const safeRedirectUri = String(redirectUri ?? "").trim();

  if (!clientId || !clientSecret || !safeRefreshToken || !safeRedirectUri) {
    return NextResponse.json({ error: "server_misconfigured" }, { status: 400 });
  }

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "refresh_token",
    refresh_token: safeRefreshToken,
    redirect_uri: safeRedirectUri,
    scope: OUTLOOK_SCOPES.join(" "),
  });

  const response = await fetch(`https://login.microsoftonline.com/${getMicrosoftTenantId()}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
    cache: "no-store",
  });

  const payload = (await response.json().catch(() => null)) as Record<string, unknown> | null;
  return NextResponse.json(payload ?? { error: "invalid_token_response" }, { status: response.status });
}
