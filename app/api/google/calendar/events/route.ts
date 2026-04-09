import { NextResponse } from "next/server";

export const runtime = "nodejs";

export async function POST(request: Request) {
  const body = (await request.json().catch(() => null)) as
    | {
        accessToken?: string;
        payload?: Record<string, unknown>;
        conferenceDataVersion?: number;
      }
    | null;

  const accessToken = String(body?.accessToken ?? "").trim();
  const payload = body?.payload;
  const conferenceDataVersion =
    typeof body?.conferenceDataVersion === "number" && Number.isFinite(body.conferenceDataVersion)
      ? body.conferenceDataVersion
      : 0;

  if (!accessToken || !payload || typeof payload !== "object" || Array.isArray(payload)) {
    return NextResponse.json({ error: { message: "missing_google_calendar_payload" } }, { status: 400 });
  }

  const endpoint = new URL("https://www.googleapis.com/calendar/v3/calendars/primary/events");
  if (conferenceDataVersion > 0) {
    endpoint.searchParams.set("conferenceDataVersion", String(conferenceDataVersion));
  }

  const response = await fetch(endpoint.toString(), {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
    cache: "no-store",
  });

  const responsePayload = (await response.json().catch(() => null)) as Record<string, unknown> | null;
  return NextResponse.json(responsePayload ?? { error: { message: "invalid_google_calendar_response" } }, { status: response.status });
}
