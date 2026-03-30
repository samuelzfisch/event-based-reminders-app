import Link from "next/link";

export default function Home() {
  return (
    <div className="space-y-8 text-gray-900">
      <section className="space-y-2">
        <div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between">
          <div>
            <h1 className="text-3xl font-bold text-gray-900">Event-Based Reminders</h1>
            <p className="mt-2 max-w-2xl text-sm text-gray-600">
              A workspace for building event-based reminder plans, templates, and future delivery settings.
            </p>
          </div>
          <div className="rounded-xl border bg-white px-4 py-3 text-sm shadow-sm md:min-w-[280px] md:max-w-[320px]">
            <div className="text-xs font-semibold uppercase tracking-wide text-gray-500">Workspace</div>
            <div className="mt-1 font-medium text-gray-900">Event-Based Reminders</div>
            <div className="mt-1 text-xs text-gray-600">Plans and settings now share the same shell and control language.</div>
          </div>
        </div>
      </section>

      <section className="rounded-2xl border bg-white shadow-sm">
        <div className="space-y-5 p-6">
          <h2 className="text-lg font-semibold text-gray-900">Open a route</h2>
          <div className="flex flex-col gap-4 sm:flex-row">
            <Link
              href="/plans"
              className="inline-flex items-center justify-center rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700"
            >
              Open Plans
            </Link>
            <Link
              href="/history"
              className="inline-flex items-center justify-center rounded-lg border px-4 py-2 text-sm font-medium text-gray-700 hover:bg-gray-50"
            >
              Open History
            </Link>
            <Link
              href="/settings"
              className="inline-flex items-center justify-center rounded-lg border px-4 py-2 text-sm font-medium text-gray-700 hover:bg-gray-50"
            >
              Open Settings
            </Link>
          </div>
        </div>
      </section>
    </div>
  );
}
