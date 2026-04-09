import { NextResponse } from "next/server";

export const runtime = "nodejs";

export async function POST(request: Request) {
  console.info("[api/provider-mail/send] route entered");
  try {
    const body = (await request.json().catch(() => null)) as
      | {
          accessToken?: string;
          raw?: string;
        }
      | null;

    console.info("[api/provider-mail/send] request body parsed", {
      hasBody: Boolean(body),
      tokenPresent: Boolean(String(body?.accessToken ?? "").trim()),
      rawLength: String(body?.raw ?? "").trim().length,
    });

    const accessToken = String(body?.accessToken ?? "").trim();
    const raw = String(body?.raw ?? "").trim();

    if (!accessToken || !raw) {
      console.error("[api/provider-mail/send] missing payload", {
        tokenPresent: Boolean(accessToken),
        rawLength: raw.length,
      });
      return NextResponse.json({ error: { message: "missing_google_send_payload" } }, { status: 400 });
    }

    console.info("[api/provider-mail/send] google api call start");
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
    console.info("[api/provider-mail/send] google api call complete", {
      ok: response.ok,
      status: response.status,
      hasPayload: Boolean(payload),
    });
    return NextResponse.json(payload ?? { error: { message: "invalid_google_send_response" } }, { status: response.status });
  } catch (error) {
    console.error("[api/provider-mail/send] route failed", {
      error: error instanceof Error ? error.message : String(error),
    });
    return NextResponse.json(
      { error: { message: error instanceof Error ? error.message : "provider_mail_send_failed" } },
      { status: 500 }
    );
  }
}
