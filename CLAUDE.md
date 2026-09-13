# GTD-Eisenhower2 — Project Context

A personal task manager combining GTD (Getting Things Done) workflow staging with the
Eisenhower Priority Matrix, two-way synced with Microsoft To Do via a Cloudflare Worker.

## Architecture

- **Front end**: `index.html` — a single self-contained file (HTML/CSS/JS all inline).
  Hosted via GitHub Pages at https://kirkwoodnel.github.io/GTD-Eisenhower2/index.html
  Repo: `kirkwoodnel/GTD-Eisenhower2`, branch `main`. Pushing to `main` redeploys
  automatically — usually live within 10–20 seconds.
- **Backend**: Cloudflare Worker `gtd-todo-sync` (`worker.js`), at
  https://gtd-todo-sync.marknelson.workers.dev — handles Microsoft Graph OAuth (PKCE),
  syncs KV-stored tasks with Microsoft To Do's "Tasks" and "Flagged Emails" lists,
  manages webhook subscriptions for real-time sync, and runs a daily cron for
  subscription renewal + a safety-net sync.
- **Storage**: Cloudflare KV namespace `GTD_KV`.
- The front end talks to the Worker via `/api/tasks` (GET replaces/returns the full
  task array; PUT saves it) using a Bearer token — `WORKER_SECRET`, hardcoded in
  `index.html` since this is a single-user personal tool.

## How sync works

1. Microsoft Graph calls the Worker's `/webhook/notify` the instant something changes
   on either To Do list (create/update/delete) — this triggers a full `runSync()`.
2. A daily Cron Trigger also runs `runSync()` as a safety net and renews the webhook
   subscriptions (they expire after ~3 days).
3. `runSync()`: processes queued deletions → pulls "Tasks" (two-way, last-write-wins by
   `lastModified`) → pulls "Flagged Emails" (one-way; only completion status is ever
   written back to that list) → pushes anything the app created/edited
   (`source: "app"`) up to Microsoft.
4. Extra fields carried from Microsoft To Do: `notes` (round-trips both ways, from the
   task's `body`) and `url` (read-only — the linked Outlook email for flagged-email
   tasks, from `linkedResources`, requires `$expand=linkedResources` on the list fetch).

## Deploying changes

- **Front end**: commit + push `index.html` to `main`.
- **Worker**: `wrangler deploy` (once wrangler is configured with this account/Worker).
  Before wrangler was set up, changes had to be pasted manually into the Cloudflare
  dashboard's Worker editor — if that's still the case, say so up front.

## Design decisions / gotchas worth knowing before changing things

- Task IDs are always strings; every lookup/comparison uses `String(a) === String(b)`
  to avoid numeric-vs-string mismatches.
- Quick date-picker buttons (+1d..+5d) build the ISO date from local date components
  (`getFullYear()/getMonth()/getDate()`), not `toISOString()` — avoids a timezone-driven
  off-by-one-day bug.
- Completing a task holds a 2-second visual delay (checkbox flips instantly, border
  flashes green) before the list actually re-renders, so it doesn't vanish the instant
  it's checked.
- Clicking a task to edit it scrolls to the form container, not the page top.
- Section headers (4 GTD columns + 4 Eisenhower quadrants) and a total counter at the
  top show live counts of *non-completed* tasks in each bucket — recalculated in
  `renderAll()`.
- Drag-and-drop works between both the GTD columns and the Eisenhower quadrants.
- A "hide completed tasks" toggle persists via `localStorage`.

## Secrets (not committed — set via `wrangler secret put` or the Cloudflare dashboard)

`SHARED_SECRET`, `MSFT_CLIENT_ID`, `MSFT_TENANT_ID`, `SETUP_KEY` — see the header
comment in `worker.js` for what each does.
