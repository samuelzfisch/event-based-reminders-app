import Link from "next/link";
import { Fraunces, Manrope } from "next/font/google";
import { ExamplePlanCarousel } from "./components/example-plan-carousel";

const fraunces = Fraunces({
  subsets: ["latin"],
  weight: ["600", "700"],
  variable: "--font-fraunces",
});

const manrope = Manrope({
  subsets: ["latin"],
  weight: ["500", "600", "700"],
  variable: "--font-manrope",
});

function IconCreateEvent() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="h-[18px] w-[18px] shrink-0 text-slate-900" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3.5" y="5" width="17" height="15.5" rx="2.5" />
      <path d="M7.5 3.5v3M16.5 3.5v3M3.5 9.5h17M12 12.5v5M9.5 15h5" />
    </svg>
  );
}

function IconChecklistLayers() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="h-[18px] w-[18px] shrink-0 text-slate-900" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M12 4.5 19.5 8 12 11.5 4.5 8 12 4.5Z" />
      <path d="M19.5 12 12 15.5 4.5 12M19.5 16 12 19.5 4.5 16" />
    </svg>
  );
}

function IconSendArrow() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="h-[18px] w-[18px] shrink-0 text-slate-900" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M20 4 11 13" />
      <path d="M20 4 14 20l-3.5-7.5L3 9l17-5Z" />
    </svg>
  );
}

function IconHistoryActivity() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="h-[18px] w-[18px] shrink-0 text-slate-900" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M4 12a8 8 0 1 0 2.3-5.6" />
      <path d="M4 4.5v4h4" />
      <path d="M12 8.5v4l2.5 1.5" />
    </svg>
  );
}

function IconOutlookMail() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="h-[18px] w-[18px] shrink-0 text-slate-900" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3.5" y="6" width="17" height="12" rx="2.5" />
      <path d="m5.5 8 6.5 5 6.5-5" />
    </svg>
  );
}

function IconGoogleCalendarDoc() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="h-[18px] w-[18px] shrink-0 text-slate-900" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <rect x="4" y="4.5" width="16" height="15" rx="2.5" />
      <path d="M8 3.5v3M16 3.5v3M4 9h16M8 13h8M8 16h5" />
    </svg>
  );
}

function IconMeetingLink() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true" className="h-[18px] w-[18px] shrink-0 text-slate-900" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M10 13.5 8 15.5a3 3 0 1 1-4.2-4.2l3-3a3 3 0 0 1 4.2 0" />
      <path d="M14 10.5 16 8.5a3 3 0 1 1 4.2 4.2l-3 3a3 3 0 0 1-4.2 0" />
      <path d="M9.5 14.5 14.5 9.5" />
    </svg>
  );
}

const workflowSteps = [
  {
    step: "01",
    title: "Choose the event",
    description: "Everything in the workflow stays tied to the moment that matters.",
    icon: IconCreateEvent,
    numberClassName: "text-slate-500",
    cardClassName: "bg-white border-slate-200 shadow-[0_18px_40px_-34px_rgba(15,23,42,0.12)]",
  },
  {
    step: "02",
    title: "Build the full workflow",
    description: "Reminders, meetings, and emails — all set up once and reused every time.",
    icon: IconChecklistLayers,
    numberClassName: "text-slate-500",
    cardClassName: "bg-white border-slate-200 shadow-[0_18px_40px_-34px_rgba(15,23,42,0.12)]",
  },
  {
    step: "03",
    title: "Run it when it comes up",
    description: "Export the reminders, meetings, and emails whenever the event is back on the calendar.",
    icon: IconSendArrow,
    numberClassName: "text-slate-500",
    cardClassName: "bg-white border-slate-200 shadow-[0_18px_40px_-34px_rgba(15,23,42,0.12)]",
  },
  {
    step: "04",
    title: "See everything that happened",
    description: "Review what was created, sent, or updated after every run.",
    icon: IconHistoryActivity,
    numberClassName: "text-slate-500",
    cardClassName: "bg-white border-slate-200 shadow-[0_18px_40px_-34px_rgba(15,23,42,0.12)]",
  },
];

const integrationTiles = [
  {
    title: "Outlook",
    icon: IconOutlookMail,
    points: [
      "Create calendar events and meetings",
      "Draft and send emails",
      "Keep the whole workflow connected in Outlook",
    ],
  },
  {
    title: "Google",
    icon: IconGoogleCalendarDoc,
    points: [
      "Create Google Calendar events",
      "Draft Gmail messages",
      "Keep the whole workflow connected in Google",
    ],
  },
  {
    title: "Meeting links",
    icon: IconMeetingLink,
    points: [
      "Add Teams or Google Meet links automatically",
      "Keep meeting details attached to events",
    ],
  },
];

const historyPoints = [
  "Review what ran, what was created, and what changed",
  "Keep a clean record of every update",
  "Check the latest run without retracing your steps",
];

export default function Home() {
  return (
    <div className={`${fraunces.variable} ${manrope.variable} relative isolate min-h-[calc(100vh-4rem)] space-y-10 overflow-hidden bg-transparent pb-10 pt-1 text-slate-900`}>
      <section className="relative overflow-hidden rounded-[34px] border border-slate-200/80 bg-white shadow-[0_24px_70px_-44px_rgba(15,23,42,0.16)]">
        <div className="home-hero-grid relative grid gap-6 px-5 py-7 lg:px-6 lg:py-8">
          <div className="home-hero-intro min-w-0 space-y-5">
            <div className="space-y-3.5">
              <div className="inline-flex max-w-full items-center rounded-full border border-slate-200 bg-white px-3 py-1 text-[9px] font-semibold uppercase tracking-[0.24em] text-slate-600 shadow-[0_10px_22px_-18px_rgba(15,23,42,0.14)]">
                Reusable Event Workflows
              </div>
              <h1 className="home-hero-title max-w-[26rem] font-[family:var(--font-fraunces)] text-[32px] font-semibold leading-[1.03] tracking-[-0.04em] text-slate-950">
                Plan an event once. Reuse it every time.
              </h1>
              <p className="home-hero-copy max-w-[23rem] font-[family:var(--font-manrope)] text-[11px] leading-5 text-slate-600">
                Build the reminders, meetings, and emails once. Then reuse the same event workflow every time it comes back.
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2.5">
              <Link
                href="/plans?new=1"
                className="inline-flex items-center justify-center rounded-2xl bg-blue-600 px-4 py-2.5 text-[11px] font-medium text-white shadow-[0_18px_34px_-24px_rgba(38,72,104,0.34)] transition hover:bg-blue-700"
              >
                Start a Plan
              </Link>
              <Link
                href="/history"
                className="inline-flex items-center justify-center rounded-2xl border border-slate-200 bg-white px-4 py-2.5 text-[11px] font-medium text-slate-700 transition hover:bg-slate-50"
              >
                View History
              </Link>
              <Link
                href="/settings"
                className="inline-flex items-center justify-center rounded-2xl border border-slate-200 bg-white px-4 py-2.5 text-[11px] font-medium text-slate-700 transition hover:bg-slate-50"
              >
                Settings
              </Link>
            </div>
          </div>

          <div className="home-hero-example [&>div>div>div:nth-child(2)]:mt-3.5 [&>div>div>div:nth-child(3)]:mt-5 [&>div>div]:p-4 [&>div]:p-3">
            <ExamplePlanCarousel />
          </div>

          <div className="home-metrics-grid grid items-stretch gap-3 pt-1">
            <div className="flex min-h-[7.4rem] flex-col rounded-2xl border border-slate-200 bg-white px-3.5 py-3.5 shadow-[0_12px_24px_-28px_rgba(15,23,42,0.22)]">
              <div className="text-[8px] font-semibold uppercase tracking-[0.2em] text-slate-500">Reuse</div>
              <div className="mt-2.5 text-[9px] leading-5 text-slate-700">Stop rebuilding the same event checklist every time.</div>
            </div>
            <div className="flex min-h-[7.4rem] flex-col rounded-2xl border border-slate-200 bg-white px-3.5 py-3.5 shadow-[0_12px_24px_-28px_rgba(15,23,42,0.22)]">
              <div className="text-[8px] font-semibold uppercase tracking-[0.2em] text-slate-500">Run</div>
              <div className="mt-2.5 text-[9px] leading-5 text-slate-700">Send the full workflow out when the event is happening again.</div>
            </div>
            <div className="flex min-h-[7.4rem] flex-col rounded-2xl border border-slate-200 bg-white px-3.5 py-3.5 shadow-[0_12px_24px_-28px_rgba(15,23,42,0.22)]">
              <div className="text-[8px] font-semibold uppercase tracking-[0.2em] text-slate-500">Track</div>
              <div className="mt-2.5 text-[9px] leading-5 text-slate-700">See what was created, sent, and updated after every run.</div>
            </div>
          </div>
        </div>
      </section>

      <section className="space-y-6 rounded-[34px] border border-slate-200/80 bg-white px-5 py-6 shadow-[0_24px_70px_-46px_rgba(15,23,42,0.14)] lg:px-6 lg:py-7">
        <div className="space-y-3">
          <div className="text-xs font-semibold uppercase tracking-[0.24em] text-slate-500">How It Works</div>
          <h2 className="font-[family:var(--font-fraunces)] text-[2rem] font-semibold tracking-[-0.03em] text-slate-950">A reusable workflow, not a one-off checklist</h2>
          <p className="max-w-2xl text-sm leading-6 text-slate-600">Plan it once. Reuse it every time.</p>
        </div>

        <div className="home-workflow-grid grid gap-4">
          {workflowSteps.map((step) => (
            <div
              key={step.step}
              className={`flex h-full flex-col rounded-[24px] border p-4 ${step.cardClassName}`}
            >
              <div className={`text-[10px] font-semibold uppercase tracking-[0.24em] ${step.numberClassName}`}>
                {step.step}
              </div>
              <div className="mt-4 flex flex-col items-start gap-2.5">
                <step.icon />
                <div className="text-[12px] font-semibold leading-4 text-slate-950">{step.title}</div>
              </div>
              <p className="mt-3.5 text-[10px] leading-5 text-slate-600">{step.description}</p>
            </div>
          ))}
        </div>
      </section>

      <section className="rounded-[34px] border border-slate-200/80 bg-white shadow-[0_24px_70px_-46px_rgba(15,23,42,0.14)]">
        <div className="space-y-6 px-5 py-6 lg:px-6 lg:py-7">
          <div className="space-y-3">
            <div className="text-xs font-semibold uppercase tracking-[0.24em] text-slate-500">Connect Your Workflow</div>
            <h2 className="font-[family:var(--font-fraunces)] text-[2rem] font-semibold tracking-[-0.03em] text-slate-950">Use the tools your team already works in</h2>
            <p className="max-w-3xl text-sm leading-6 text-slate-600">
              Keep the workflow reusable without moving the work into another system.
            </p>
          </div>

          <div className="home-integrations-grid grid gap-4">
            {integrationTiles.map((tile) => (
              <div
                key={tile.title}
                className="flex h-full flex-col rounded-[24px] border border-slate-200 bg-white p-4 shadow-[0_14px_32px_-30px_rgba(15,23,42,0.12)]"
              >
                <div className="inline-flex flex-col items-start gap-2.5">
                  <tile.icon />
                  <span className="text-[11px] font-semibold uppercase tracking-[0.18em] text-slate-800">{tile.title}</span>
                </div>
                <ul className="mt-4 space-y-2.5 text-[10px] leading-5 text-slate-700">
                  {tile.points.map((point) => (
                    <li key={point} className="flex items-start gap-2">
                      <span className="mt-2 h-1.5 w-1.5 shrink-0 rounded-full bg-slate-300" />
                      <span>{point}</span>
                    </li>
                  ))}
                </ul>
              </div>
            ))}
          </div>
        </div>
      </section>

      <section className="rounded-[34px] border border-slate-200/80 bg-white shadow-[0_24px_70px_-46px_rgba(15,23,42,0.14)]">
        <div className="grid gap-8 px-5 py-6 md:grid-cols-[minmax(0,1fr)_minmax(0,270px)] md:items-start md:gap-7 lg:px-6 lg:py-7">
          <div className="min-w-0 space-y-5">
            <div className="text-xs font-semibold uppercase tracking-[0.24em] text-slate-500">History</div>
            <h2 className="flex min-w-0 items-center gap-2.5 font-[family:var(--font-fraunces)] text-[2rem] font-semibold tracking-[-0.03em] text-slate-950">
              <IconHistoryActivity />
              <span className="min-w-0">Track what happened</span>
            </h2>
            <p className="max-w-2xl text-base leading-7 text-slate-600">
              After you run a plan, you can see exactly what happened without digging through separate tools.
            </p>
            <p className="max-w-2xl text-sm leading-6 text-slate-600">
              Review what was created, sent, or updated, then keep a clean record of each run over time.
            </p>
          </div>

          <div className="w-full min-w-0 rounded-[24px] border border-slate-200 bg-white p-5 shadow-[0_18px_40px_-34px_rgba(15,23,42,0.12)] lg:p-6">
            <div className="text-sm font-semibold uppercase tracking-[0.18em] text-slate-500">Everything stays tracked</div>
            <ul className="mt-5 space-y-3.5 text-sm leading-6 text-slate-700">
              {historyPoints.map((point) => (
                <li key={point} className="flex items-start gap-2">
                  <span className="mt-2 h-1.5 w-1.5 shrink-0 rounded-full bg-slate-300" />
                  <span>{point}</span>
                </li>
              ))}
            </ul>
            <div className="mt-6">
              <Link
                href="/history"
                className="inline-flex items-center justify-center rounded-2xl border border-slate-200 bg-white px-5 py-3 text-sm font-medium text-slate-700 transition hover:bg-slate-50"
              >
                Go to History
              </Link>
            </div>
          </div>
        </div>
      </section>
    </div>
  );
}
