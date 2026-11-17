type Draft = { to: string[]; cc?: string[]; subject: string; body: string }

// Outlook on the web. Set NEXT_PUBLIC_OUTLOOK_COMPOSE_URL in .env.local to switch between
// "https://outlook.live.com/mail/deeplink/compose" (personal) and the default (work/school).
const OUTLOOK_COMPOSE =
  process.env.NEXT_PUBLIC_OUTLOOK_COMPOSE_URL ?? "https://outlook.office.com/mail/deeplink/compose"
const MAX_BODY = 1500 // links have a length limit, so long reports get truncated

export function outlookDraftUrl({ to, cc = [], subject, body }: Draft) {
  // Remove any "Sender: ..." line and leading blank space from the report text.
  const cleaned = body
    .split("\n")
    .filter((line) => !/^\s*sender\s*:/i.test(line))
    .join("\n")
    .replace(/^\s+/, "")

  const trimmed =
    cleaned.length > MAX_BODY
      ? cleaned.slice(0, MAX_BODY) + "\n\n[Report shortened. Full report attached separately.]"
      : cleaned

  const fields: [string, string][] = [
    ["to", to.join(",")],
    ["cc", cc.join(",")],
    ["subject", subject],
    ["body", trimmed],
  ]
  // encodeURIComponent writes spaces as %20, which Outlook reads correctly (unlike "+").
  const query = fields
    .filter(([, value]) => value)
    .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
    .join("&")

  return `${OUTLOOK_COMPOSE}?${query}`
}

export function openOutlookDraft(draft: Draft): boolean {
  const w = window.open(outlookDraftUrl(draft), "_blank")
  if (!w) return false // blocked by the browser's pop-up blocker
  w.opener = null
  return true
}