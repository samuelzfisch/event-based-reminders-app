import { NextResponse } from "next/server";

export const runtime = "nodejs";

function parseGoogleCalendarResponsePayload(rawText: string) {
  if (!rawText.trim()) return null;
  try {
    return JSON.parse(rawText) as Record<string, unknown>;
  } catch {
    return {
      error: {
        message: rawText.trim(),
      },
    } satisfies Record<string, unknown>;
  }
}

function buildGoogleCalendarEventEndpoint(eventId?: string, conferenceDataVersion?: number) {
  const endpoint = eventId
    ? new URL(`https://www.googleapis.com/calendar/v3/calendars/primary/events/${encodeURIComponent(eventId)}`)
    : new URL("https://www.googleapis.com/calendar/v3/calendars/primary/events");
  if (conferenceDataVersion && conferenceDataVersion > 0) {
    endpoint.searchParams.set("conferenceDataVersion", String(conferenceDataVersion));
  }
  return endpoint;
}

async function parseRequestBody(request: Request) {
  return (await request.json().catch(() => null)) as
    | {
        accessToken?: string;
        payload?: Record<string, unknown>;
        conferenceDataVersion?: number;
        eventId?: string;
      }
    | null;
}

function readConferenceDataVersion(value: unknown) {
  return typeof value === "number" && Number.isFinite(value) ? value : 0;
}

export async function POST(request: Request) {
  try {
    const body = await parseRequestBody(request);

    const accessToken = String(body?.accessToken ?? "").trim();
    const payload = body?.payload;
    const conferenceDataVersion = readConferenceDataVersion(body?.conferenceDataVersion);

    if (!accessToken || !payload || typeof payload !== "object" || Array.isArray(payload)) {
      console.error("[api/provider-calendar/events] missing payload", {
        tokenPresent: Boolean(accessToken),
        hasPayload: Boolean(payload),
        conferenceDataVersion,
      });
      return NextResponse.json({ error: { message: "missing_google_calendar_payload" } }, { status: 400 });
    }

    const response = await fetch(buildGoogleCalendarEventEndpoint(undefined, conferenceDataVersion).toString(), {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
      cache: "no-store",
    });

    const responseText = await response.text().catch(() => "");
    const responsePayload = parseGoogleCalendarResponsePayload(responseText);
    return NextResponse.json(
      responsePayload ?? { error: { message: "invalid_google_calendar_response" } },
      { status: response.status }
    );
  } catch (error) {
    console.error("[api/provider-calendar/events] request failed", {
      error: error instanceof Error ? error.message : String(error),
    });
    return NextResponse.json(
      { error: { message: error instanceof Error ? error.message : "provider_calendar_event_failed" } },
      { status: 500 }
    );
  }
}

export async function PATCH(request: Request) {
  try {
    const body = await parseRequestBody(request);

    const accessToken = String(body?.accessToken ?? "").trim();
    const payload = body?.payload;
    const eventId = String(body?.eventId ?? "").trim();
    const conferenceDataVersion = readConferenceDataVersion(body?.conferenceDataVersion);

    if (!accessToken || !eventId || !payload || typeof payload !== "object" || Array.isArray(payload)) {
      console.error("[api/provider-calendar/events] missing patch payload", {
        tokenPresent: Boolean(accessToken),
        eventIdPresent: Boolean(eventId),
        hasPayload: Boolean(payload),
        conferenceDataVersion,
      });
      return NextResponse.json({ error: { message: "missing_google_calendar_patch_payload" } }, { status: 400 });
    }

    const response = await fetch(buildGoogleCalendarEventEndpoint(eventId, conferenceDataVersion).toString(), {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
      cache: "no-store",
    });

    const responseText = await response.text().catch(() => "");
    const responsePayload = parseGoogleCalendarResponsePayload(responseText);
    return NextResponse.json(
      responsePayload ?? { error: { message: "invalid_google_calendar_response" } },
      { status: response.status }
    );
  } catch (error) {
    console.error("[api/provider-calendar/events] patch failed", {
      error: error instanceof Error ? error.message : String(error),
    });
    return NextResponse.json(
      { error: { message: error instanceof Error ? error.message : "provider_calendar_event_patch_failed" } },
      { status: 500 }
    );
  }
}

export async function DELETE(request: Request) {
  try {
    const body = await parseRequestBody(request);

    const accessToken = String(body?.accessToken ?? "").trim();
    const eventId = String(body?.eventId ?? "").trim();

    if (!accessToken || !eventId) {
      console.error("[api/provider-calendar/events] missing delete payload", {
        tokenPresent: Boolean(accessToken),
        eventIdPresent: Boolean(eventId),
      });
      return NextResponse.json({ error: { message: "missing_google_calendar_delete_payload" } }, { status: 400 });
    }

    const response = await fetch(buildGoogleCalendarEventEndpoint(eventId).toString(), {
      method: "DELETE",
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
      cache: "no-store",
    });

    if (response.status === 204) {
      return new Response(null, { status: 204 });
    }

    const responseText = await response.text().catch(() => "");
    const responsePayload = parseGoogleCalendarResponsePayload(responseText);
    return NextResponse.json(
      responsePayload ?? { error: { message: "invalid_google_calendar_response" } },
      { status: response.status }
    );
  } catch (error) {
    console.error("[api/provider-calendar/events] delete failed", {
      error: error instanceof Error ? error.message : String(error),
    });
    return NextResponse.json(
      { error: { message: error instanceof Error ? error.message : "provider_calendar_event_delete_failed" } },
      { status: 500 }
    );
  }
}
