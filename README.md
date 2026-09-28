# Word Legal Copilot

A Word add-in for legal teams: it reads the document currently open in
Word, compares it against the company rulebook, and surfaces
recommended changes as Word comments the reviewer can accept or
dismiss.

## How it works

- **Task pane add-in** (`app/taskpane/page.tsx`) — loaded inside Word
  via Office.js. Pulls the full document text, sends it to the backend,
  and renders findings. Each finding can be inserted into the document
  as a Word comment anchored to the offending text.
- **Backend agent** (`app/api/review/route.ts` + `lib/agent.ts`) — a
  Next.js API route that sends the document text and the company
  rulebook to Claude, forcing a structured response (via tool use) of
  `{ quote, issue, severity, suggestion, rulebookCitation }` findings.
- **Rulebook** (`rulebook/company-rulebook.md`) — MVP uses full-context
  stuffing: the whole file is included in every request. Replace the
  placeholder with your real rulebook. Once it's large or changes
  often, swap this for retrieval (chunk + embed + fetch relevant
  sections per request) instead of sending the whole thing every time.

## Local development

```bash
npm install
cp .env.example .env.local   # fill in ANTHROPIC_API_KEY
npm run dev
```

This serves the app at `http://localhost:3000`. The task pane route is
`/taskpane`.

Office Add-ins require the task pane to be served over **HTTPS**, even
in local testing — Word will refuse to load a plain `http://localhost`
URL. Two ways to satisfy that during development:

1. **Deploy to a free HTTPS host** (recommended for MVP) — e.g. Vercel.
   Point the manifest at that URL and sideload against the deployed
   app. Fastest path to something you can actually test inside Word.
2. **Local HTTPS** — generate dev certs with
   [`office-addin-dev-certs`](https://www.npmjs.com/package/office-addin-dev-certs)
   and run Next.js behind a local HTTPS proxy (e.g. `local-ssl-proxy`
   in front of `next dev`). More moving parts; only worth it if you're
   iterating heavily offline.

## Sideloading into Word (testing)

The app is deployed at **https://word-legal-plugin.vercel.app**, and
`manifest/manifest.xml` already points at it — make sure `ANTHROPIC_API_KEY`
is set in the Vercel project's environment variables before testing.

1. In Word: **Insert → Add-ins → My Add-ins → Upload My Add-in**, and
   select `manifest/manifest.xml`.
2. The "Legal Copilot" group appears on the Home tab. Click
   **Review Rulebook** to open the task pane.

(If you redeploy under a different domain, update every URL in
`manifest/manifest.xml` to match before sideloading.)

(On Mac, or for org-wide testing without publishing, you can also use
`Insert → Add-ins → Admin Managed`, or shared-folder sideloading —
see Microsoft's
[sideloading docs](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins).)

## Publishing to your organization

Actual publishing is an admin action on your Microsoft 365 tenant, not
something done from this repo:

- **Org-wide (recommended for internal tools):** a Microsoft 365 admin
  uploads `manifest/manifest.xml` (pointing at your production URL) via
  the **Microsoft 365 admin center → Settings → Integrated apps →
  Upload custom apps**, and assigns it to specific users or the whole
  org.
- **AppSource (public listing):** submit through **Partner Center**;
  requires Microsoft's validation process. Usually unnecessary for an
  internal legal tool.

Whoever has tenant admin access should do this step once the app is
deployed at a stable URL.

## Current MVP scope / known limitations

- **Comments, not Track Changes.** Findings are applied as Word
  comments (`Range.insertComment`), not tracked-change edits. This is
  non-destructive and simpler to get right; a track-changes mode can be
  added later as an alternative apply path.
- **Full-context rulebook.** No retrieval/vector search yet — the whole
  rulebook file is sent on every request. Fine for a short policy doc,
  not for a large or frequently-updated one.
- **No auth yet.** The API route has no access control. Before any real
  deployment, add authentication (Entra ID / Azure AD SSO via
  `Office.context.auth.getAccessTokenAsync` is the standard pattern for
  Word add-ins in a Microsoft 365 tenant) and validate the token
  server-side.
- **Whole-document review only.** No selection-based or
  incremental/paragraph-level review yet.
