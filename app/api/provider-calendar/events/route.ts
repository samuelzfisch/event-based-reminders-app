import { NextResponse } from "next/server";

export const runtime = "nodejs";

export async function POST(request: Request) {
  console.info("[api/provider-calendar/events] route entered");
  try {
    const body = (await request.json().catch(() => null)) as
      | {
          accessToken?: string;
          payload?: Record<string, unknown>;
          conferenceDataVersion?: number;
        }
      | null;

    console.info("[api/provider-calendar/events] request body parsed", {
      hasBody: Boolean(body),
      tokenPresent: Boolean(String(body?.accessToken ?? "").trim()),
      hasPayload: Boolean(body?.payload && typeof body.payload === "object" && !Array.isArray(body.payload)),
      conferenceDataVersion: body?.conferenceDataVersion ?? null,
      requestPayload: body?.payload ?? null,
    });

    const accessToken = String(body?.accessToken ?? "").trim();
    const payload = body?.payload;
    const conferenceDataVersion =
      typeof body?.conferenceDataVersion === "number" && Number.isFinite(body.conferenceDataVersion)
        ? body.conferenceDataVersion
        : 0;

    if (!accessToken || !payload || typeof payload !== "object" || Array.isArray(payload)) {
      console.error("[api/provider-calendar/events] missing payload", {
        tokenPresent: Boolean(accessToken),
        hasPayload: Boolean(payload),
        conferenceDataVersion,
      });
      return NextResponse.json({ error: { message: "missing_google_calendar_payload" } }, { status: 400 });
    }

    const endpoint = new URL("https://www.googleapis.com/calendar/v3/calendars/primary/events");
    if (conferenceDataVersion > 0) {
      endpoint.searchParams.set("conferenceDataVersion", String(conferenceDataVersion));
    }

    console.info("[api/provider-calendar/events] google api call start", {
      endpoint: endpoint.toString(),
    });
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
    console.info("[api/provider-calendar/events] google api call complete", {
      ok: response.ok,
      status: response.status,
      hasPayload: Boolean(responsePayload),
      responsePayload,
    });
    return NextResponse.json(
      responsePayload ?? { error: { message: "invalid_google_calendar_response" } },
      { status: response.status }
    );
  } catch (error) {
    console.error("[api/provider-calendar/events] route failed", {
      error: error instanceof Error ? error.message : String(error),
    });
    return NextResponse.json(
      { error: { message: error instanceof Error ? error.message : "provider_calendar_event_failed" } },
      { status: 500 }
    );
  }
}
