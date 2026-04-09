import { NextResponse } from "next/server";

export const runtime = "nodejs";

export async function POST(request: Request) {
  const body = (await request.json().catch(() => null)) as
    | {
        accessToken?: string;
        raw?: string;
      }
    | null;

  const accessToken = String(body?.accessToken ?? "").trim();
  const raw = String(body?.raw ?? "").trim();

  if (!accessToken || !raw) {
    return NextResponse.json({ error: { message: "missing_google_send_payload" } }, { status: 400 });
  }

  const response = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/messages/send", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ raw }),
    cache: "no-store",
  });

  const payload = (await response.json().catch(() => null)) as Record<string, unknown> | null;
  return NextResponse.json(payload ?? { error: { message: "invalid_google_send_response" } }, { status: response.status });
}
