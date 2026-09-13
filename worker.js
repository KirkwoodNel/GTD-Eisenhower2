/**
 * GTD-Eisenhower <-> Microsoft To Do sync Worker
 *
 * Bindings required (set in Cloudflare dashboard > Worker > Settings):
 *   KV namespace : GTD_KV
 *   Secret       : SHARED_SECRET     (long random string; the app sends it as a Bearer token)
 *   Secret       : MSFT_CLIENT_ID    (d924cf88-067c-49e5-b4bc-fc4105664437)
 *   Secret       : MSFT_TENANT_ID    (73234503-195d-4792-8e06-f09746f05f11)
 *   Secret       : SETUP_KEY         (one-off random string, only used to gate /auth/start so a
 *                                     stranger can't hijack your Microsoft sign-in on this Worker)
 *
 * KV keys used:
 *   "tasks"               -> JSON array, the single source of truth for the GTD app
 *   "pending_deletions"   -> JSON array of {externalId, listId}, tasks deleted in the app that
 *                            still need deleting from Microsoft To Do
 *   "msft_refresh_token"  -> current Microsoft refresh token (offline_access)
 *   "msft_tasks_list_id"  -> cached id of the "Tasks" To Do list
 *   "msft_flagged_list_id"-> cached id of the "Flagged Emails" To Do list
 *   "last_sync_time"      -> ISO string, last time runSync() completed (shown in the app)
 *   "graph_client_state"  -> random secret Microsoft echoes back on every webhook notification,
 *                            so we can tell a real notification from a spoofed one
 *   "sub_tasks"           -> {id, expirationDateTime} of the webhook subscription on the Tasks list
 *   "sub_flagged"         -> {id, expirationDateTime} of the webhook subscription on Flagged Emails
 *
 * Routes:
 *   GET    /api/tasks          -> full task array (what the app renders)
 *   PUT    /api/tasks          -> replace the full task array (what the app calls on every edit)
 *   GET    /auth/start?key=... -> begin the one-time Microsoft sign-in (PKCE, public client)
 *   GET    /auth/callback      -> Microsoft redirects here after consent
 *   POST   /webhook/notify     -> Microsoft Graph change-notification endpoint (validation
 *                                 handshake + real notifications for both To Do lists)
 *   GET    /webhook/setup      -> (Bearer-auth'd) manually (re)creates both subscriptions and
 *                                 runs a sync immediately, instead of waiting for the daily cron
 *
 * How syncing is triggered (webhook-driven, not polling):
 *   1. Microsoft Graph calls /webhook/notify the moment something changes on the "Tasks" list or
 *      the "Flagged Emails" list (create/update/delete). We don't try to interpret exactly what
 *      changed - we just kick off the same full runSync() used everywhere else, which is cheap
 *      and always correct.
 *   2. A daily Cron Trigger calls dailyMaintenance(), which renews both subscriptions (they
 *      expire after ~3 days if untouched) and also runs one full runSync() as a safety net, in
 *      case a notification was ever missed or dropped.
 *
 * runSync() itself, each time it runs:
 *   1. Deletes anything queued in "pending_deletions" from Microsoft To Do.
 *   2. Pulls the "Tasks" list from Microsoft To Do, merges into KV (last-write-wins by
 *      lastModified).
 *   3. Pulls "Flagged Emails" one-way; the only thing ever written back to that list is
 *      completion status (you can't create/edit items in it, only complete them).
 *   4. Pushes anything created/edited in the app (source:"app") up to the "Tasks" list.
 *
 * Extra fields carried over from Microsoft To Do:
 *   "notes" -> the task's body/note text (round-trips both ways, like title/dueDate).
 *   "url"   -> for tasks created by flagging an Outlook email, the link back to that email
 *              (from the task's linkedResources - $expand=linkedResources on the list fetch).
 *              Read-only: never written back to Microsoft.
 */

const GRAPH = "https://graph.microsoft.com/v1.0";
const CORS_ORIGIN = "https://kirkwoodnel.github.io";
// Azure AD requires an Origin header on token requests for SPA-registered clients
// (AADSTS9002327). Server-side fetch() doesn't send one automatically, so we add it.
const WORKER_ORIGIN = "https://gtd-todo-sync.marknelson.workers.dev";

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": CORS_ORIGIN,
    "Access-Control-Allow-Methods": "GET,PUT,OPTIONS",
    "Access-Control-Allow-Headers": "Authorization,Content-Type",
    // So the app's JS can read this header on the response (browsers hide custom
    // response headers cross-origin unless the server explicitly exposes them).
    "Access-Control-Expose-Headers": "X-Last-Sync",
  };
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders() },
  });
}

function unauthorized() {
  return json({ error: "unauthorized" }, 401);
}

function isAuthed(request, env) {
  return (request.headers.get("Authorization") || "") === `Bearer ${env.SHARED_SECRET}`;
}

async function getTasks(env) {
  const raw = await env.GTD_KV.get("tasks");
  return raw ? JSON.parse(raw) : [];
}

async function getPendingDeletions(env) {
  const raw = await env.GTD_KV.get("pending_deletions");
  return raw ? JSON.parse(raw) : [];
}

function newId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

// ---------- Microsoft Graph / OAuth helpers ----------

function b64url(buf) {
  return btoa(String.fromCharCode(...new Uint8Array(buf)))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

async function sha256(str) {
  return await crypto.subtle.digest("SHA-256", new TextEncoder().encode(str));
}

async function getValidAccessToken(env) {
  const refreshToken = await env.GTD_KV.get("msft_refresh_token");
  if (!refreshToken) throw new Error("Not connected to Microsoft yet - visit /auth/start once.");

  const body = new URLSearchParams({
    client_id: env.MSFT_CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: "offline_access Tasks.ReadWrite",
  });

  const resp = await fetch(
    `https://login.microsoftonline.com/${env.MSFT_TENANT_ID}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded", "Origin": WORKER_ORIGIN }, body }
  );
  if (!resp.ok) throw new Error(`Token refresh failed: ${await resp.text()}`);
  const tok = await resp.json();
  if (tok.refresh_token) await env.GTD_KV.put("msft_refresh_token", tok.refresh_token);
  return tok.access_token;
}

async function graphFetch(env, path, options = {}) {
  const accessToken = await getValidAccessToken(env);
  const resp = await fetch(`${GRAPH}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (resp.status === 404) return null;
  if (!resp.ok) throw new Error(`Graph ${path} failed: ${resp.status} ${await resp.text()}`);
  return resp.status === 204 ? null : resp.json();
}

async function getListId(env, wellknownName, kvKey) {
  const cached = await env.GTD_KV.get(kvKey);
  if (cached) return cached;
  const lists = await graphFetch(env, "/me/todo/lists");
  const match = lists.value.find((l) => l.wellknownListName === wellknownName);
  if (!match) throw new Error(`Could not find To Do list: ${wellknownName}`);
  await env.GTD_KV.put(kvKey, match.id);
  return match.id;
}

function toGraphTask(task) {
  const body = { title: task.title, status: task.completed ? "completed" : "notStarted" };
  body.body = { content: task.notes || "", contentType: "text" };
  if (task.dueDate) {
    body.dueDateTime = { dateTime: `${task.dueDate}T${task.dueTime || "00:00"}:00`, timeZone: "Europe/London" };
  }
  if (task.reminderDate) {
    body.reminderDateTime = { dateTime: `${task.reminderDate}T${task.reminderTime || "09:00"}:00`, timeZone: "Europe/London" };
    body.isReminderOn = true;
  }
  return body;
}

// Microsoft To Do notes come back as an itemBody ({content, contentType}). contentType is
// usually "text", but strip tags defensively in case it's ever "html".
function bodyToNotes(body) {
  if (!body || !body.content) return "";
  if (body.contentType === "html") {
    return body.content
      .replace(/<[^>]*>/g, "")
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .trim();
  }
  return body.content.trim();
}

// Flagging an email in Outlook creates a To Do task with a linkedResources entry pointing back
// at the message (webUrl). Requires $expand=linkedResources on the list request.
function firstLinkedUrl(graphTask) {
  if (Array.isArray(graphTask.linkedResources)) {
    const withUrl = graphTask.linkedResources.find((r) => r.webUrl);
    if (withUrl) return withUrl.webUrl;
  }
  return "";
}

function fromGraphTask(graphTask, listId, existing) {
  return {
    id: existing ? existing.id : newId(),
    title: graphTask.title,
    gtdStatus: existing ? existing.gtdStatus : "inbox",
    priority: existing ? existing.priority : "q4",
    dueDate: graphTask.dueDateTime ? graphTask.dueDateTime.dateTime.slice(0, 10) : "",
    dueTime: graphTask.dueDateTime ? graphTask.dueDateTime.dateTime.slice(11, 16) : "",
    reminderDate: graphTask.reminderDateTime ? graphTask.reminderDateTime.dateTime.slice(0, 10) : "",
    reminderTime: graphTask.reminderDateTime ? graphTask.reminderDateTime.dateTime.slice(11, 16) : "",
    context: existing ? existing.context : "",
    notes: bodyToNotes(graphTask.body),
    url: firstLinkedUrl(graphTask),
    completed: graphTask.status === "completed",
    externalId: graphTask.id,
    listId,
    lastModified: graphTask.lastModifiedDateTime || new Date().toISOString(),
    source: "todo",
  };
}

// ---------- Sync ----------

async function processPendingDeletions(env) {
  const pending = await getPendingDeletions(env);
  if (!pending.length) return;
  const remaining = [];
  for (const p of pending) {
    try {
      await graphFetch(env, `/me/todo/lists/${p.listId}/tasks/${p.externalId}`, { method: "DELETE" });
    } catch (e) {
      remaining.push(p); // retry next run
    }
  }
  await env.GTD_KV.put("pending_deletions", JSON.stringify(remaining));
}

async function runSync(env) {
  await processPendingDeletions(env);

  const tasksListId = await getListId(env, "defaultList", "msft_tasks_list_id");
  const flaggedListId = await getListId(env, "flaggedEmails", "msft_flagged_list_id");

  let localTasks = await getTasks(env);

  const remoteTasks = (await graphFetch(env, `/me/todo/lists/${tasksListId}/tasks?$expand=linkedResources`)).value;
  for (const rt of remoteTasks) {
    const idx = localTasks.findIndex((t) => t.externalId === rt.id);
    if (idx === -1) {
      localTasks.push(fromGraphTask(rt, tasksListId));
    } else {
      const local = localTasks[idx];
      const remoteNewer = new Date(rt.lastModifiedDateTime) > new Date(local.lastModified || 0);
      if (remoteNewer && local.source !== "app") {
        localTasks[idx] = fromGraphTask(rt, tasksListId, local);
      }
    }
  }

  const flaggedTasks = (await graphFetch(env, `/me/todo/lists/${flaggedListId}/tasks?$expand=linkedResources`)).value;
  for (const rt of flaggedTasks) {
    const idx = localTasks.findIndex((t) => t.externalId === rt.id);
    if (idx === -1) {
      const mapped = fromGraphTask(rt, flaggedListId);
      mapped.gtdStatus = "inbox";
      localTasks.push(mapped);
    } else if (localTasks[idx].completed && rt.status !== "completed") {
      await graphFetch(env, `/me/todo/lists/${flaggedListId}/tasks/${rt.id}`, {
        method: "PATCH",
        body: JSON.stringify({ status: "completed" }),
      });
    }
  }

  for (let i = 0; i < localTasks.length; i++) {
    const t = localTasks[i];
    if (t.source !== "app") continue; // only push tasks the app actually changed
    if (t.listId === flaggedListId) continue; // can only complete these, never create/edit

    if (!t.externalId) {
      const created = await graphFetch(env, `/me/todo/lists/${tasksListId}/tasks`, {
        method: "POST",
        body: JSON.stringify(toGraphTask(t)),
      });
      localTasks[i] = { ...t, externalId: created.id, listId: tasksListId, lastModified: created.lastModifiedDateTime, source: "todo" };
    } else {
      await graphFetch(env, `/me/todo/lists/${tasksListId}/tasks/${t.externalId}`, {
        method: "PATCH",
        body: JSON.stringify(toGraphTask(t)),
      });
      localTasks[i] = { ...t, source: "todo" };
    }
  }

  await env.GTD_KV.put("tasks", JSON.stringify(localTasks));
  await env.GTD_KV.put("last_sync_time", new Date().toISOString());
}

// ---------- Webhook subscriptions (Microsoft Graph change notifications) ----------

async function getClientState(env) {
  let cs = await env.GTD_KV.get("graph_client_state");
  if (!cs) {
    cs = b64url(crypto.getRandomValues(new Uint8Array(24)));
    await env.GTD_KV.put("graph_client_state", cs);
  }
  return cs;
}

// Creates a subscription if we don't have one yet, otherwise renews the existing one.
// If renewal fails (e.g. Microsoft already dropped it - a 404), falls back to creating fresh.
async function ensureSubscription(env, resource, kvKey, clientState) {
  // Just under the ~4230 minute (3 day) max lifetime for todoTask subscriptions.
  const expiration = new Date(Date.now() + 4200 * 60 * 1000).toISOString();
  const stored = await env.GTD_KV.get(kvKey);
  const record = stored ? JSON.parse(stored) : null;

  if (record && record.id) {
    try {
      const patched = await graphFetch(env, `/subscriptions/${record.id}`, {
        method: "PATCH",
        body: JSON.stringify({ expirationDateTime: expiration }),
      });
      if (patched) {
        await env.GTD_KV.put(kvKey, JSON.stringify({ id: record.id, expirationDateTime: expiration }));
        return;
      }
      // patched === null means Graph returned 404 - subscription is gone; fall through and recreate.
    } catch (e) {
      // Any other renewal failure - also fall through and recreate rather than leaving sync dark.
    }
  }

  const created = await graphFetch(env, "/subscriptions", {
    method: "POST",
    body: JSON.stringify({
      changeType: "created,updated,deleted",
      notificationUrl: `${WORKER_ORIGIN}/webhook/notify`,
      resource,
      expirationDateTime: expiration,
      clientState,
    }),
  });
  await env.GTD_KV.put(kvKey, JSON.stringify({ id: created.id, expirationDateTime: expiration }));
}

// Runs once a day (from the Cron Trigger): keeps both subscriptions alive and does one full
// sync as a safety net in case a webhook notification was ever missed or dropped.
async function dailyMaintenance(env) {
  const clientState = await getClientState(env);

  try {
    const tasksListId = await getListId(env, "defaultList", "msft_tasks_list_id");
    await ensureSubscription(env, `/me/todo/lists/${tasksListId}/tasks`, "sub_tasks", clientState);
  } catch (e) {
    console.error("ensureSubscription(tasks) failed:", e.message);
  }

  try {
    const flaggedListId = await getListId(env, "flaggedEmails", "msft_flagged_list_id");
    await ensureSubscription(env, `/me/todo/lists/${flaggedListId}/tasks`, "sub_flagged", clientState);
  } catch (e) {
    console.error("ensureSubscription(flagged) failed:", e.message);
  }

  try {
    await runSync(env);
  } catch (e) {
    console.error("runSync failed during daily maintenance:", e.message);
  }
}

// Handles both the one-time validation handshake and real change notifications from Graph.
async function handleWebhookNotify(request, env, ctx) {
  const url = new URL(request.url);
  const validationToken = url.searchParams.get("validationToken");
  if (validationToken !== null) {
    // Microsoft Graph validates a new subscription by POSTing ?validationToken=... and expects
    // the exact plain-text token echoed back within 10 seconds.
    return new Response(validationToken, {
      status: 200,
      headers: { "Content-Type": "text/plain" },
    });
  }

  let payload;
  try {
    payload = await request.json();
  } catch (e) {
    return new Response(null, { status: 202 }); // nothing usable - ack anyway
  }

  const clientState = await getClientState(env);
  const notifications = Array.isArray(payload.value) ? payload.value : [];
  const legit = notifications.some((n) => n.clientState === clientState);

  if (legit) {
    // Don't bother interpreting which task/list changed - just trigger the same full sync
    // used everywhere else. Respond immediately: Graph requires a 2xx within 3 seconds.
    ctx.waitUntil(runSync(env));
  }

  return new Response(null, { status: 202 });
}

// ---------- HTTP handlers ----------

async function handleAuthStart(request, env) {
  const url = new URL(request.url);
  if (url.searchParams.get("key") !== env.SETUP_KEY) return unauthorized();

  const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
  await env.GTD_KV.put("pkce_verifier", verifier, { expirationTtl: 600 });
  const challenge = b64url(await sha256(verifier));

  const authUrl = new URL(`https://login.microsoftonline.com/${env.MSFT_TENANT_ID}/oauth2/v2.0/authorize`);
  authUrl.search = new URLSearchParams({
    client_id: env.MSFT_CLIENT_ID,
    response_type: "code",
    redirect_uri: `${url.origin}/auth/callback`,
    response_mode: "query",
    scope: "offline_access Tasks.ReadWrite",
    code_challenge: challenge,
    code_challenge_method: "S256",
  }).toString();

  return Response.redirect(authUrl.toString(), 302);
}

async function handleAuthCallback(request, env) {
  const url = new URL(request.url);
  const code = url.searchParams.get("code");
  if (!code) return new Response(`Sign-in failed: ${url.searchParams.get("error_description") || "no code"}`, { status: 400 });

  const verifier = await env.GTD_KV.get("pkce_verifier");
  const body = new URLSearchParams({
    client_id: env.MSFT_CLIENT_ID,
    grant_type: "authorization_code",
    code,
    redirect_uri: `${url.origin}/auth/callback`,
    code_verifier: verifier || "",
    scope: "offline_access Tasks.ReadWrite",
  });

  const resp = await fetch(
    `https://login.microsoftonline.com/${env.MSFT_TENANT_ID}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded", "Origin": WORKER_ORIGIN }, body }
  );
  if (!resp.ok) return new Response(`Token exchange failed: ${await resp.text()}`, { status: 400 });
  const tok = await resp.json();
  await env.GTD_KV.put("msft_refresh_token", tok.refresh_token);

  return new Response("Connected to Microsoft To Do. You can close this tab.", {
    headers: { "Content-Type": "text/plain" },
  });
}

async function handleApi(request, env) {
  if (!isAuthed(request, env)) return unauthorized();

  if (request.method === "GET") {
    const tasks = await getTasks(env);
    const lastSync = await env.GTD_KV.get("last_sync_time");
    return new Response(JSON.stringify(tasks), {
      headers: {
        "Content-Type": "application/json",
        ...corsHeaders(),
        ...(lastSync ? { "X-Last-Sync": lastSync } : {}),
      },
    });
  }

  if (request.method === "PUT") {
    const incoming = await request.json();
    const previous = await getTasks(env);

    const stillPresent = new Set(incoming.filter((t) => t.externalId).map((t) => t.externalId));
    const newlyRemoved = previous.filter((p) => p.externalId && !stillPresent.has(p.externalId));
    if (newlyRemoved.length) {
      const pending = await getPendingDeletions(env);
      for (const t of newlyRemoved) pending.push({ externalId: t.externalId, listId: t.listId });
      await env.GTD_KV.put("pending_deletions", JSON.stringify(pending));
    }

    await env.GTD_KV.put("tasks", JSON.stringify(incoming));
    return json(incoming);
  }

  return json({ error: "not found" }, 404);
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (request.method === "OPTIONS") return new Response(null, { headers: corsHeaders() });
    if (url.pathname === "/auth/start") return handleAuthStart(request, env);
    if (url.pathname === "/auth/callback") return handleAuthCallback(request, env);
    if (url.pathname === "/api/tasks") return handleApi(request, env);
    if (url.pathname === "/webhook/notify") return handleWebhookNotify(request, env, ctx);
    if (url.pathname === "/webhook/setup") {
      if (!isAuthed(request, env)) return unauthorized();
      ctx.waitUntil(dailyMaintenance(env));
      return json({ ok: true, message: "Subscriptions being (re)created and a sync started in the background." });
    }
    return new Response("Not found", { status: 404 });
  },

  async scheduled(event, env, ctx) {
    ctx.waitUntil(dailyMaintenance(env));
  },
};
