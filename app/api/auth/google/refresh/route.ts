import { NextResponse } from "next/server";

function getGoogleClientId() {
  return String(process.env.GOOGLE_CLIENT_ID ?? process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID ?? "").trim();
}

function getGoogleClientSecret() {
  return String(process.env.GOOGLE_CLIENT_SECRET ?? "").trim();
}

export async function POST(request: Request) {
  const { refreshToken } = (await request.json().catch(() => ({}))) as {
    refreshToken?: string;
  };

  const clientId = getGoogleClientId();
  const clientSecret = getGoogleClientSecret();
  const safeRefreshToken = String(refreshToken ?? "").trim();

  if (!clientId || !clientSecret || !safeRefreshToken) {
    return NextResponse.json({ error: "server_misconfigured" }, { status: 400 });
  }

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "refresh_token",
    refresh_token: safeRefreshToken,
  });

  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
    cache: "no-store",
  });

  const payload = (await response.json().catch(() => null)) as Record<string, unknown> | null;
  return NextResponse.json(payload ?? { error: "invalid_token_response" }, { status: response.status });
}
