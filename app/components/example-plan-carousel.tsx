"use client";

import { useEffect, useEffectEvent, useState } from "react";

function IconCalendarTimeline() {
  return (
    <svg
      viewBox="0 0 24 24"
      aria-hidden="true"
      className="h-[18px] w-[18px] shrink-0 text-slate-900"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <rect x="3.5" y="5" width="17" height="15.5" rx="2.5" />
      <path d="M7.5 3.5v3M16.5 3.5v3M3.5 9.5h17" />
      <path d="M8 13h3M8 16h6" />
    </svg>
  );
}

type ExampleSlide = {
  title: string;
  description: string;
  actions: Array<{
    kind: string;
    title: string;
    schedule: string;
    time: string;
    accent: string;
    labelColor: string;
  }>;
};

const EXAMPLE_SLIDES: ExampleSlide[] = [
  {
    title: "Vacation",
    description: "A simple event schedule with useful follow-up before departure.",
    actions: [
      {
        kind: "Reminder",
        title: "Check passport",
        schedule: "10 Days Prior to Event",
        time: "9:00 AM",
        accent: "bg-blue-500",
        labelColor: "text-blue-700",
      },
      {
        kind: "Meeting",
        title: "Review itinerary with Alex",
        schedule: "3 Days Prior to Event",
        time: "2:00 PM",
        accent: "bg-violet-500",
        labelColor: "text-violet-700",
      },
      {
        kind: "Email",
        title: "Send final arrival details to Jordan",
        schedule: "1 Day Prior to Event",
        time: "5:00 PM",
        accent: "bg-green-500",
        labelColor: "text-green-700",
      },
    ],
  },
  {
    title: "Investor Meeting",
    description: "Keep leadership, prep, and follow-up aligned before an important business event.",
    actions: [
      {
        kind: "Reminder",
        title: "Finalize talking points",
        schedule: "3 Days Prior to Event",
        time: "8:30 AM",
        accent: "bg-blue-500",
        labelColor: "text-blue-700",
      },
      {
        kind: "Meeting",
        title: "Run prep review with finance",
        schedule: "1 Day Prior to Event",
        time: "4:00 PM",
        accent: "bg-violet-500",
        labelColor: "text-violet-700",
      },
      {
        kind: "Email",
        title: "Send meeting agenda and dial-in details",
        schedule: "Event Morning",
        time: "7:45 AM",
        accent: "bg-green-500",
        labelColor: "text-green-700",
      },
    ],
  },
  {
    title: "Annual Conference",
    description: "Reuse the same launch sequence for speakers, logistics, and attendee updates.",
    actions: [
      {
        kind: "Reminder",
        title: "Confirm booth materials arrival",
        schedule: "7 Days Prior to Event",
        time: "11:00 AM",
        accent: "bg-blue-500",
        labelColor: "text-blue-700",
      },
      {
        kind: "Meeting",
        title: "Speaker check-in with events team",
        schedule: "2 Days Prior to Event",
        time: "1:30 PM",
        accent: "bg-violet-500",
        labelColor: "text-violet-700",
      },
      {
        kind: "Email",
        title: "Send attendee arrival instructions",
        schedule: "1 Day Prior to Event",
        time: "6:15 PM",
        accent: "bg-green-500",
        labelColor: "text-green-700",
      },
    ],
  },
];

export function ExamplePlanCarousel() {
  const [activeIndex, setActiveIndex] = useState(0);
  const [isPaused, setIsPaused] = useState(false);

  const advanceSlide = useEffectEvent(() => {
    setActiveIndex((current) => (current + 1) % EXAMPLE_SLIDES.length);
  });

  useEffect(() => {
    if (isPaused) return;
    const intervalId = window.setInterval(() => {
      advanceSlide();
    }, 4800);

    return () => window.clearInterval(intervalId);
  }, [isPaused]);

  return (
    <div
      className="min-w-0 rounded-[28px] border border-slate-200 bg-white p-2 shadow-[0_28px_56px_-40px_rgba(15,23,42,0.18)]"
      onMouseEnter={() => setIsPaused(true)}
      onMouseLeave={() => setIsPaused(false)}
    >
      <div className="overflow-hidden rounded-[24px] border border-slate-200 bg-[var(--app-bg-soft)] p-3.5 shadow-[inset_0_1px_0_rgba(255,255,255,0.85)]">
        <div className="inline-flex items-center gap-2 text-[9px] font-semibold uppercase tracking-[0.2em] text-slate-500">
          <IconCalendarTimeline />
          <span>Example event plans</span>
        </div>

        <div className="mt-3 overflow-hidden">
          <div
            className="flex transition-transform duration-500 ease-[cubic-bezier(0.22,1,0.36,1)]"
            style={{ transform: `translateX(-${activeIndex * 100}%)` }}
          >
            {EXAMPLE_SLIDES.map((slide) => (
              <div key={slide.title} className="w-full shrink-0 overflow-hidden px-1">
                <div className="text-[13px] font-semibold text-slate-950">{slide.title}</div>
                <div className="mt-1 max-w-[26rem] text-[9px] leading-4 text-slate-500">{slide.description}</div>

                <div className="mt-4 space-y-2.5">
                  {slide.actions.map((action) => (
                    <div
                      key={`${slide.title}-${action.kind}-${action.title}`}
                      className="flex items-start gap-2 rounded-2xl border border-slate-200 bg-white px-2.5 py-2.5 shadow-[0_8px_18px_-20px_rgba(15,23,42,0.3)]"
                    >
                      <span className={`mt-1 h-2 w-2 shrink-0 rounded-full ${action.accent}`} />
                      <div className="min-w-0 flex-1">
                        <div className="flex flex-wrap items-baseline gap-x-2 gap-y-1">
                          <span className={`text-[10px] font-semibold ${action.labelColor}`}>{action.kind}:</span>
                          <span className="text-[10px] text-slate-900">{action.title}</span>
                        </div>
                        <div className="mt-1.5 text-[8px] text-slate-500">
                          {action.schedule} <span className="mx-2 text-slate-300">-</span> {action.time}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className="mt-5 flex justify-center">
          <div className="inline-flex items-center gap-1.5 rounded-full bg-slate-900/82 px-3 py-1.5 shadow-[0_10px_24px_-18px_rgba(15,23,42,0.45)]">
            {EXAMPLE_SLIDES.map((slide, index) => {
              const isActive = index === activeIndex;
              return (
                <button
                  key={slide.title}
                  type="button"
                  onClick={() => setActiveIndex(index)}
                  aria-label={`Show ${slide.title} example`}
                  aria-pressed={isActive}
                  className={`h-2 rounded-full transition-all duration-300 ${
                    isActive ? "w-5 bg-white" : "w-2 bg-white/45 hover:bg-white/70"
                  }`}
                />
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}
