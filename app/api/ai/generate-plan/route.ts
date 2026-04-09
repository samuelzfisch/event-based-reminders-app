import OpenAI from "openai";
import { NextResponse } from "next/server";

import {
  aiPlanChatSchema,
  diagnoseAIPlanChatTurnResult,
  parseAIPlanChatRequest,
  parseAIPlanChatTurnResult,
} from "../../../../lib/aiPlanGeneration";

export const runtime = "nodejs";

const MAX_MESSAGE_CHARS = 4000;
const MAX_MESSAGES = 10;
const AI_DEBUG_PREFIX = "[plans-ai]";

export async function POST(request: Request) {
  const apiKey = String(process.env.OPENAI_API_KEY ?? "").trim();
  if (!apiKey) {
    return NextResponse.json(
      { error: "Missing OPENAI_API_KEY. Add it to .env.local and restart the dev server." },
      { status: 500 }
    );
  }

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid JSON body." }, { status: 400 });
  }

  const parsedRequest = parseAIPlanChatRequest(body);
  if (!parsedRequest) {
    return NextResponse.json({ error: "Please send at least one message." }, { status: 400 });
  }

  const client = new OpenAI({ apiKey });
  const recentMessages = parsedRequest.messages.slice(-MAX_MESSAGES);

  try {
    const response = await client.responses.create({
      model: "gpt-4.1-mini",
      input: [
        {
          role: "system",
          content: [
            "You are a planning assistant inside an event-based reminders app.",
            "Behave like a concise guided planning assistant, not a generic chatbot.",
            "Stay focused on helping the user shape a practical event workflow made of reminders, emails, and meetings.",
            "Produce a usable draft as early as reasonably possible.",
            "Ask targeted follow-up questions only when they materially improve the plan, and avoid blocking on optional details.",
            "If some details are missing, make practical business-friendly assumptions, mention them briefly, and still return a workable draft.",
            "Do not claim the app supports capabilities outside the existing builder.",
            "The structured draft must stay compatible with a builder that supports only base types earnings, conference, and press_release, so choose the closest one when the workflow is more general.",
            "Use these mapping heuristics when the workflow is general: investor/earnings/reporting -> earnings; conference/roadshow/event follow-up -> conference; announcement/launch/distribution/comms timeline -> press_release; other business preparation workflows should usually map to conference unless the request is clearly announcement-oriented.",
            "Use rowType 'calendar_event' for meetings.",
            "Use empty arrays instead of invented recipients or attendees when unknown.",
            "Prefer concrete, high-signal row titles over generic ones like 'Reminder 1' or 'Email task'.",
            "Prefer 3 to 6 useful rows unless the user clearly wants something larger.",
            "Sequence rows in a realistic order and avoid awkward duplicates.",
            "When exact times are unknown, use practical business-hour defaults and keep them consistent across the draft instead of leaving everything vague.",
            "For earnings-style plans, prefer a structure like prep reminder(s), internal prep meeting, day-before/day-of reminder, and follow-up email.",
            "For conference-style plans, prefer a structure like pre-event prep, day-of reminder, and post-event follow-up outreach.",
            "For press-release timelines, prefer a structure like internal review reminder(s), approval/alignment step, and distribution-day actions.",
            "For general business presentation or client-workflow requests, prefer a structure like early prep reminder, internal alignment meeting, final prep reminder, and optional follow-up email.",
            "If the date is unknown, set noEventDate true and leave anchorDate empty.",
            "If current builder context is provided in refine mode, treat it as the working plan you are editing.",
            "In refine mode, prefer returning a revised full draft that incorporates the user's requested changes, such as moving reminders earlier, removing rows, adding meetings, or making the timeline more aggressive.",
            "Do not respond with vague advice when you can return an updated draft.",
            "Preserve the overall plan intent unless the user clearly asks to rebuild it.",
            "If current builder context is absent or start_new mode is selected, behave as a fresh-plan assistant.",
            "Keep assistantMessage short, natural, and helpful, and use it to briefly explain what changed or what is still needed.",
            "In refine mode, also return a concise changeSummary list with 2 to 5 short items describing the main edits you made.",
            "Return a short confidenceNote that sounds human, such as 'High confidence', 'Needs confirmation on timing', or 'Assumed the follow-up email should be removed'.",
            "When useful, return 0 to 4 suggestedNextActions as short chip-friendly phrases like 'Move reminders earlier' or 'Add an internal prep meeting'.",
            "Only suggest actions that fit the existing builder constraints.",
            "If little changed or you still need clarification, say that plainly.",
            "Set status to 'ready_to_apply' when the draft is actionable enough to load into the builder, even if some optional details are still unknown.",
          ].join(" "),
        },
        {
          role: "user",
          content: [
            {
              type: "input_text",
              text: [
                "Generate the next assistant turn and the latest structured draft for the plans builder.",
                "Ask at most 3 follow-up questions in a turn, and only if they are worth asking.",
                "If you already have enough to build a useful first draft, do that instead of stalling for more information.",
                "When refine mode is active, revise the full draft instead of describing edits abstractly.",
                "Keep changeSummary short and practical. Do not write a verbose diff.",
                "Only return suggestedNextActions when they are genuinely useful, and keep them short.",
                "Choose practical, scenario-appropriate defaults instead of vague placeholder structure when exact details are missing.",
                "Prefer a coherent timeline with realistic sequencing over a long list of generic tasks.",
                "For general business requests, map to the closest supported builder type and still produce a useful plan draft rather than apologizing for the constraint.",
                "",
                "Current builder mode:",
                parsedRequest.builderContextMode === "refine_current" ? "Refine the current builder plan" : "Start a new plan draft",
                "",
                "Current builder summary:",
                parsedRequest.currentBuilderContext ? JSON.stringify(parsedRequest.currentBuilderContext) : "none",
                "",
                "Current draft summary:",
                parsedRequest.currentSummary || "none",
                "",
                "Current structured draft JSON:",
                parsedRequest.currentDraft ? JSON.stringify(parsedRequest.currentDraft) : "none",
                "",
                "Recent conversation:",
                recentMessages
                  .map((message) => `${message.role === "assistant" ? "Assistant" : "User"}: ${message.text.slice(0, MAX_MESSAGE_CHARS)}`)
                  .join("\n\n"),
              ].join("\n"),
            },
          ],
        },
      ],
      text: {
        format: {
          type: "json_schema",
          ...aiPlanChatSchema,
        },
      },
    });

    const rawOutput = response.output_text;
    if (!rawOutput) {
      console.error(AI_DEBUG_PREFIX, "model returned no output_text");
      return NextResponse.json({ error: "AI returned no output." }, { status: 500 });
    }

    console.info(AI_DEBUG_PREFIX, "raw model output preview", rawOutput.slice(0, 2000));

    let parsedJson: unknown;
    try {
      parsedJson = JSON.parse(rawOutput);
    } catch {
      console.error(AI_DEBUG_PREFIX, "invalid JSON from model", rawOutput.slice(0, 2000));
      return NextResponse.json({ error: "AI returned invalid JSON." }, { status: 500 });
    }

    const parsed = parseAIPlanChatTurnResult(parsedJson);
    if (!parsed?.draft || parsed.draft.rows.length === 0) {
      const issues = diagnoseAIPlanChatTurnResult(parsedJson);
      console.error(AI_DEBUG_PREFIX, "unusable draft response", {
        issues,
        parsedSummary: parsed?.summary ?? null,
        parsedAssistantMessage: parsed?.assistantMessage ?? null,
        parsedRowCount: parsed?.draft?.rows.length ?? 0,
        rawOutputPreview: rawOutput.slice(0, 2000),
      });
      const detail = issues.length ? ` ${issues.join(" ")}` : "";
      const message = parsed?.assistantMessage
        ? `AI returned assistant text, but the draft was still unusable after normalization.${detail}`
        : `AI could not build a usable plan draft.${detail}`;
      return NextResponse.json({ error: message }, { status: 500 });
    }

    console.info(AI_DEBUG_PREFIX, "accepted draft", {
      status: parsed.status,
      rowCount: parsed.draft.rows.length,
      baseType: parsed.draft.baseType,
    });

    return NextResponse.json(parsed);
  } catch (error) {
    const message = error instanceof Error ? error.message : "AI plan generation failed.";
    if (/429|quota/i.test(message)) {
      return NextResponse.json({ error: "OpenAI quota exceeded." }, { status: 429 });
    }
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
