export function isAuthBypassEnabled() {
  return String(process.env.NEXT_PUBLIC_AUTH_BYPASS_ENABLED ?? "").trim().toLowerCase() === "true";
}
