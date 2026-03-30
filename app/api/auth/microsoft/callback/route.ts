import { NextResponse } from "next/server";

export async function GET(request: Request) {
  const url = new URL(request.url);
  const code = url.searchParams.get("code");
  const state = url.searchParams.get("state");
  const error = url.searchParams.get("error");
  const errorDescription = url.searchParams.get("error_description");
  const origin = `${url.protocol}//${url.host}`;

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
          code: ${JSON.stringify(code)},
          state: ${JSON.stringify(state)},
          error: ${JSON.stringify(error)},
          errorDescription: ${JSON.stringify(errorDescription)},
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
    },
  });
}
