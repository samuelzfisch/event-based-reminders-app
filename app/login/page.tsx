"use client";

import { InlineAuthCard } from "../components/inline-auth-card";

export default function LoginPage() {
  return (
    <div className="flex min-h-screen items-center justify-center bg-gray-50 px-4 py-8">
      <div className="w-full max-w-md">
        <InlineAuthCard />
      </div>
    </div>
  );
}
