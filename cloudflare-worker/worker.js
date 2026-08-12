// DPEG AI Summarize + To Do — Cloudflare Worker
// POST /                 → AI email summary (Groq, validated via MSAL user token)
// POST /todo             → Create Microsoft To Do task for a DPEG recipient (app credentials)
// POST /assignment       → Create/update a shared task-assignment record (D1)
// GET  /assignments      → Fetch assignments for the current user (assigned to me + by me)
// POST /assignment-status→ Recipient updates assignment status/progress (D1)
// POST /assignment-alert → Assigner sends a one-click "update required" nudge (D1)
// POST /assignment-alert-clear → Either party dismisses that alert (D1)

const ALLOWED_ORIGIN  = 'https://dpeg-software.github.io';
const ALLOWED_ORIGINS = new Set([
  ALLOWED_ORIGIN,
  'http://localhost:8765',
  'https://dpeg-task-manager-staging-test.pages.dev',
]);
const DPEG_TENANT_ID  = '9152bf5c-22ff-4e4a-8624-784a2d243006';
const AZURE_CLIENT_ID = '8d523e65-0163-49c7-881b-407c0222527e';
const STANDARD_DEPARTMENTS = new Set([
  'investor relations','accounting','acquisitions','development','software development','construction',
  'property management','maintenance','marketing','legal and title','leasing',
  'it','operations','lending','insurance','multifamily','eb-5',
]);

const CORS = {
  'Access-Control-Allow-Origin':  ALLOWED_ORIGIN,
  'Access-Control-Allow-Methods': 'GET, POST, PUT, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
};

const DATA_KEY = 'company-state';
const ADMIN_EMAILS = new Set(['systemmanager1@dhananipeg.com', 'propertymanagement2@dhananipeg.com']);
const PROOF_START = 'DPEG_PROOF_START';
const PROOF_END = 'DPEG_PROOF_END';
const PROOF_LINK_PREFIX = 'proof-link:';
let assignmentColumnsReady = false;
let taskMessagesTableReady = false;

// Staging must never mutate employee Microsoft To Do data or send application
// content to external AI services. Production behavior remains unchanged when
// APP_ENV is absent (as it is in the existing production Worker).
const STAGING_EXTERNAL_PATHS = new Set([
  '/',
  '/attachment-summary',
  '/todo',
  '/todo-update',
  '/poll-completions',
  '/proof-task',
  '/proof-submit',
]);

function isStaging(env) {
  return String(env.APP_ENV || '').trim().toLowerCase() === 'staging';
}

function externalEffectsAllowed(env) {
  return !isStaging(env);
}

function stagingSafetyResponse(path) {
  return json({
    error: 'staging_safety_block',
    message: 'This staging endpoint is disabled to protect production Microsoft and AI data.',
    path,
  }, 409);
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...CORS, 'Content-Type': 'application/json' },
  });
}

// Apply CORS after routing so every endpoint—including redirects and errors—
// gets the correct origin. Browsers accept only one Allow-Origin value, so a
// comma-separated list would not work; echo the requesting origin only when
// it is one of the two explicitly trusted app addresses.
function withRequestCors(response, request) {
  const origin = request.headers.get('Origin') || '';
  const allowedOrigin = ALLOWED_ORIGINS.has(origin) ? origin : ALLOWED_ORIGIN;
  const headers = new Headers(response.headers);
  headers.set('Access-Control-Allow-Origin', allowedOrigin);
  headers.set('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  headers.set('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  headers.append('Vary', 'Origin');
  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers,
  });
}

function graphDueDate(date) {
  const raw = String(date || '').trim();
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return null;
  const y = Number(match[1]);
  const m = Number(match[2]);
  const d = Number(match[3]);
  if (!y || m < 1 || m > 12 || d < 1 || d > 31) return null;
  return {
    dateTime: `${match[1]}-${match[2]}-${match[3]}T17:00:00.0000000`,
    timeZone: 'Central Standard Time',
  };
}

function extractEmailAddress(value) {
  const raw = String(value || '').trim();
  const angle = raw.match(/<([^<>@\s]+@[^<>\s]+)>/);
  if (angle) return angle[1].trim().toLowerCase();
  const email = raw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return email ? email[0].trim().toLowerCase() : raw.toLowerCase();
}

function proofSubmitUrl(body, listId, taskId) {
  const base = String(body.proofBaseUrl || 'https://dpeg-software.github.io/dhanani_task_manager/').split('#')[0];
  const url = new URL(base);
  url.searchParams.set('proof', '1');
  url.searchParams.set('taskId', String(body.appTaskId || ''));
  url.searchParams.set('recipientEmail', String(body.recipientEmail || ''));
  url.searchParams.set('assignedByName', String(body.assignedByName || ''));
  url.searchParams.set('assignedByEmail', String(body.assignedByEmail || ''));
  url.searchParams.set('title', String(body.title || 'Task'));
  url.searchParams.set('proofShareUrl', String(body.proofShareUrl || ''));
  url.searchParams.set('proofInstructions', String(body.proofInstructions || ''));
  url.searchParams.set('todoListId', listId);
  url.searchParams.set('todoTaskId', taskId);
  return url.toString();
}

function workerOrigin(request) {
  return new URL(request.url).origin;
}

function proofShortCode() {
  const bytes = new Uint8Array(9);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, b => b.toString(36).padStart(2, '0')).join('').slice(0, 12);
}

async function createProofShortUrl(request, env, targetUrl) {
  if (!env.DPEG_DATA) return targetUrl;
  const code = proofShortCode();
  await env.DPEG_DATA.put(`${PROOF_LINK_PREFIX}${code}`, JSON.stringify({
    targetUrl,
    createdAt: new Date().toISOString(),
  }), { expirationTtl: 60 * 60 * 24 * 180 });
  return `${workerOrigin(request)}/p/${code}`;
}

async function handleProofRedirect(request, env, code) {
  const cleanCode = String(code || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 12);
  if (!env.DPEG_DATA || !cleanCode) return new Response('Proof link not found', { status: 404, headers: CORS });
  const record = await env.DPEG_DATA.get(`${PROOF_LINK_PREFIX}${cleanCode}`, 'json');
  if (!record?.targetUrl) return new Response('Proof link expired or not found', { status: 404, headers: CORS });
  return Response.redirect(record.targetUrl, 302);
}

function parseProofs(text) {
  const raw = String(text || '');
  const start = raw.indexOf(PROOF_START);
  const end = raw.indexOf(PROOF_END);
  if (start < 0 || end < 0 || end <= start) return [];
  try {
    const parsed = JSON.parse(raw.slice(start + PROOF_START.length, end).trim());
    return Array.isArray(parsed?.proofs) ? parsed.proofs : [];
  } catch {
    return [];
  }
}

function userEmailFromClaims(claims) {
  return extractEmailAddress(claims.preferred_username || claims.upn || claims.email || '');
}

async function ensureAssignmentProofColumns(env) {
  if (!env.DPEG_ASSIGNMENTS || assignmentColumnsReady) return;
  const columns = [
    ['proof_status', "TEXT DEFAULT 'none'"],
    ['proof_submitted_at', 'TEXT'],
    ['proof_reviewed_at', 'TEXT'],
    ['proof_notification_id', 'TEXT'],
    ['update_alert_at', 'TEXT'],
    ['reminder_count', 'INTEGER NOT NULL DEFAULT 0'],
    ['assigner_message_seen_count', 'INTEGER NOT NULL DEFAULT 0'],
    ['recipient_message_seen_count', 'INTEGER NOT NULL DEFAULT 0'],
    ['recipient_reminder_seen_count', 'INTEGER'],
    ['version', 'INTEGER NOT NULL DEFAULT 1'],
    ['cancel_reason', 'TEXT'],
    ['cancelled_at', 'TEXT'],
  ];
  for (const [name, type] of columns) {
    try {
      await env.DPEG_ASSIGNMENTS.prepare(`ALTER TABLE assignments ADD COLUMN ${name} ${type}`).run();
    } catch (err) {
      if (!String(err?.message || '').toLowerCase().includes('duplicate column')) throw err;
    }
  }
  await env.DPEG_ASSIGNMENTS.prepare(
    'CREATE INDEX IF NOT EXISTS idx_assignments_recipient_created ON assignments(recipient_email, created_at)'
  ).run();
  await env.DPEG_ASSIGNMENTS.prepare(
    'CREATE INDEX IF NOT EXISTS idx_assignments_assigner_created ON assignments(assigner_email, created_at)'
  ).run();
  assignmentColumnsReady = true;
}

async function updateAssignmentProofState(env, details) {
  if (!env.DPEG_ASSIGNMENTS) return;
  await ensureAssignmentProofColumns(env);
  const now = new Date().toISOString();
  const appTaskId = String(details.appTaskId || '').trim();
  const recipientEmail = extractEmailAddress(details.recipientEmail || '');
  const senderEmail = extractEmailAddress(details.senderEmail || '');
  const proofStatus = String(details.proofStatus || '').trim();
  if (!appTaskId || !proofStatus) return;

  const submittedAt = proofStatus === 'submitted' ? now : null;
  const reviewedAt = proofStatus === 'approved' || proofStatus === 'declined' ? now : null;
  const notificationId = String(details.notificationId || '');
  await env.DPEG_ASSIGNMENTS.prepare(
    `UPDATE assignments
       SET proof_status = ?,
           proof_submitted_at = COALESCE(?, proof_submitted_at),
           proof_reviewed_at = CASE WHEN ? = 'submitted' THEN NULL ELSE COALESCE(?, proof_reviewed_at) END,
           proof_notification_id = COALESCE(NULLIF(?, ''), proof_notification_id),
           status = CASE WHEN ? = 'approved' THEN 'Done' ELSE status END,
           update_alert_at = CASE WHEN ? IN ('submitted','approved','declined') THEN NULL ELSE update_alert_at END,
           updated_at = ?,
           version = version + 1
     WHERE app_task_id = ?
       AND (? = '' OR recipient_email = ?)
       AND (? = '' OR assigner_email = ?)`
  ).bind(
    proofStatus,
    submittedAt,
    proofStatus,
    reviewedAt,
    notificationId,
    proofStatus,
    proofStatus,
    now,
    appTaskId,
    recipientEmail,
    recipientEmail,
    senderEmail,
    senderEmail,
  ).run();
}

// Task conversations use D1 rather than the shared KV document. Each message
// is an independent INSERT, so two people sending at the same time cannot
// overwrite one another. KV remains the source for proof records and legacy
// conversation history while the app is migrated.
async function ensureTaskMessagesTable(env) {
  if (!env.DPEG_ASSIGNMENTS || taskMessagesTableReady) return;
  // The staging database is always created from the checked-in migrations, so
  // its table and indexes already exist. Avoid issuing four redundant DDL
  // queries whenever Cloudflare starts a fresh Worker isolate.
  if (isStaging(env)) {
    taskMessagesTableReady = true;
    return;
  }
  await env.DPEG_ASSIGNMENTS.prepare(
    `CREATE TABLE IF NOT EXISTS task_messages (
       id TEXT PRIMARY KEY,
       app_task_id TEXT NOT NULL,
       task_title TEXT NOT NULL DEFAULT '',
       assigner_email TEXT NOT NULL,
       recipient_email TEXT NOT NULL,
       recipient_name TEXT NOT NULL DEFAULT '',
       sender_email TEXT NOT NULL,
       sender_name TEXT NOT NULL DEFAULT '',
       sender_role TEXT NOT NULL CHECK (sender_role IN ('assignor','assignee')),
       message TEXT NOT NULL,
       created_at TEXT NOT NULL
     )`
  ).run();
  await env.DPEG_ASSIGNMENTS.prepare(
    'CREATE INDEX IF NOT EXISTS idx_task_messages_thread ON task_messages(app_task_id, recipient_email, created_at)'
  ).run();
  await env.DPEG_ASSIGNMENTS.prepare(
    'CREATE INDEX IF NOT EXISTS idx_task_messages_assigner ON task_messages(assigner_email, created_at)'
  ).run();
  await env.DPEG_ASSIGNMENTS.prepare(
    'CREATE INDEX IF NOT EXISTS idx_task_messages_recipient ON task_messages(recipient_email, created_at)'
  ).run();
  await env.DPEG_ASSIGNMENTS.prepare(
    'CREATE INDEX IF NOT EXISTS idx_assignments_task_recipient ON assignments(app_task_id, recipient_email)'
  ).run();
  taskMessagesTableReady = true;
}

async function insertTaskMessage(env, body, claims) {
  if (!env.DPEG_ASSIGNMENTS) {
    return { error: json({ error: 'D1 task-message storage is not configured' }, 501) };
  }
  await ensureTaskMessagesTable(env);
  const appTaskId = String(body.appTaskId || '').trim();
  const recipientEmail = extractEmailAddress(body.recipientEmail || '');
  const assignerEmail = extractEmailAddress(body.assignerEmail || '');
  const senderEmail = userEmailFromClaims(claims);
  const message = String(body.message || '').trim();
  if (!appTaskId || !recipientEmail || !assignerEmail) {
    return { error: json({ error: 'Task and both participants are required' }, 400) };
  }
  if (!message) return { error: json({ error: 'Message is required' }, 400) };
  const assignment = await env.DPEG_ASSIGNMENTS.prepare(
    `SELECT title, assigner_email, recipient_email, recipient_name
       FROM assignments
      WHERE app_task_id = ? AND recipient_email = ?
      ORDER BY created_at DESC LIMIT 1`
  ).bind(appTaskId, recipientEmail).first();
  if (!assignment) {
    return { error: json({ error: 'Task assignment was not found' }, 404) };
  }
  const canonicalAssigner = extractEmailAddress(assignment.assigner_email || '');
  const canonicalRecipient = extractEmailAddress(assignment.recipient_email || '');
  if (senderEmail !== canonicalAssigner && senderEmail !== canonicalRecipient) {
    return { error: json({ error: 'You are not a participant in this task conversation' }, 403) };
  }
  const senderRole = senderEmail === canonicalRecipient ? 'assignee' : 'assignor';
  const id = `fm-${crypto.randomUUID()}`;
  const createdAt = new Date().toISOString();
  await env.DPEG_ASSIGNMENTS.prepare(
    `INSERT INTO task_messages
       (id, app_task_id, task_title, assigner_email, recipient_email, recipient_name,
        sender_email, sender_name, sender_role, message, created_at)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
  ).bind(
    id,
    appTaskId,
    String(assignment.title || body.taskTitle || ''),
    canonicalAssigner,
    canonicalRecipient,
    String(assignment.recipient_name || body.recipientName || ''),
    senderEmail,
    String(body.senderName || claims.name || ''),
    senderRole,
    message,
    createdAt,
  ).run();
  return { id, createdAt };
}

async function loadTaskMessageThreads(env, claims, options = {}) {
  if (!env.DPEG_ASSIGNMENTS) return [];
  await ensureTaskMessagesTable(env);
  const email = userEmailFromClaims(claims);
  const taskId = String(options.taskId || '').trim();
  const recipientEmail = extractEmailAddress(options.recipientEmail || '');
  const since = String(options.since || '').trim();
  let sql = `SELECT id, app_task_id, task_title, assigner_email, recipient_email, recipient_name,
                    sender_email, sender_name, sender_role, message, created_at
               FROM task_messages
              WHERE (assigner_email = ? OR recipient_email = ?)`;
  const bindings = [email, email];
  if (taskId) {
    sql += ' AND app_task_id = ?';
    bindings.push(taskId);
  }
  if (recipientEmail) {
    sql += ' AND recipient_email = ?';
    bindings.push(recipientEmail);
  }
  if (since) {
    // Inclusive comparison intentionally repeats at most the newest timestamp;
    // the browser deduplicates by message id. This avoids missing messages
    // created in the same millisecond as the previous polling cursor.
    sql += ' AND created_at >= ?';
    bindings.push(since);
  }
  sql += ' ORDER BY created_at ASC, id ASC';
  const result = await env.DPEG_ASSIGNMENTS.prepare(sql).bind(...bindings).all();
  const groups = new Map();
  for (const row of result.results || []) {
    const key = `${row.app_task_id}|${extractEmailAddress(row.recipient_email)}`;
    if (!groups.has(key)) {
      groups.set(key, {
        id: `d1-thread-${row.app_task_id}-${extractEmailAddress(row.recipient_email)}`,
        type: 'task_followup',
        appTaskId: String(row.app_task_id || ''),
        taskTitle: String(row.task_title || ''),
        senderEmail: String(row.assigner_email || ''),
        recipientEmail: String(row.recipient_email || ''),
        recipientName: String(row.recipient_name || ''),
        thread: [],
        followupStatus: '',
        status: 'open',
        createdAt: String(row.created_at || ''),
        seen: false,
      });
    }
    const group = groups.get(key);
    group.thread.push({
      id: String(row.id),
      by: row.sender_role === 'assignee' ? 'assignee' : 'assignor',
      email: String(row.sender_email || ''),
      name: String(row.sender_name || ''),
      message: String(row.message || ''),
      createdAt: String(row.created_at || ''),
    });
    group.taskTitle ||= String(row.task_title || '');
    group.updatedAt = String(row.created_at || '');
    group.followupStatus = row.sender_role === 'assignee' ? 'answered' : 'question';
  }
  return [...groups.values()];
}

function mergeD1TaskThreads(notifications, d1Threads) {
  const output = [...notifications];
  for (const d1Thread of d1Threads) {
    const sameTask = n =>
      String(n.appTaskId || '') === String(d1Thread.appTaskId || '') &&
      extractEmailAddress(n.recipientEmail || '') === extractEmailAddress(d1Thread.recipientEmail || '');
    const legacyIndex = output.findIndex(n => n.type === 'task_followup' && sameTask(n));
    let legacyThread = legacyIndex >= 0 && Array.isArray(output[legacyIndex].thread)
      ? output[legacyIndex].thread
      : [];
    if (!legacyThread.length) {
      let newestProofTime = -Infinity;
      output.forEach(n => {
        if (n.type === 'proof_submitted' && sameTask(n) && Array.isArray(n.thread) && n.thread.length) {
          const time = new Date(n.updatedAt || n.submittedAt || n.createdAt || 0).getTime();
          if (time >= newestProofTime) { legacyThread = n.thread; newestProofTime = time; }
        }
      });
    }
    const byId = new Map();
    [...legacyThread, ...d1Thread.thread].forEach(item => byId.set(String(item.id || ''), item));
    const merged = {
      ...(legacyIndex >= 0 ? output[legacyIndex] : {}),
      ...d1Thread,
      thread: [...byId.values()].sort((a, b) => {
        const timeDiff = new Date(a.createdAt || 0).getTime() - new Date(b.createdAt || 0).getTime();
        return timeDiff || String(a.id || '').localeCompare(String(b.id || ''));
      }),
    };
    if (legacyIndex >= 0) output[legacyIndex] = merged;
    else output.push(merged);
  }
  return output;
}

function todoTaskUrl(userEmail, listId, taskId) {
  return `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userEmail)}/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`;
}

function esc(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function parseProofBlock(text) {
  const raw = String(text || '');
  const start = raw.indexOf(PROOF_START);
  const end = raw.indexOf(PROOF_END);
  if (start < 0 || end < 0 || end <= start) return { proofs: [], base: raw.trim() };
  // Strip any surrounding HTML tag that wraps the proof markers
  const before = raw.slice(0, start).replace(/<[^>]*>\s*$/, '').trim();
  const after = raw.slice(end + PROOF_END.length).replace(/^\s*<\/[^>]*>/, '').trim();
  return { proofs: parseProofs(raw), base: [before, after].filter(Boolean).join('\n\n') };
}

function buildProofBlock(base, proofs) {
  return `${String(base || '').trim()}\n\n${PROOF_START}\n${JSON.stringify({ proofs }, null, 2)}\n${PROOF_END}`.trim();
}

function stripHtml(html) {
  return String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<hr[^>]*>/gi, '\n────────────────────────\n')
    .replace(/<\/p>/gi, '\n').replace(/<\/div>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&nbsp;/g, ' ')
    .replace(/\n{3,}/g, '\n\n').trim();
}

function decodeToken(token) {
  try {
    const payload = token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
    return JSON.parse(atob(payload));
  } catch {
    return null;
  }
}

// Validate the user's MSAL Bearer token (shared by every endpoint).
// decodeToken() only base64-decodes the payload — it does NOT check the
// signature. So the actual proof of identity here is the call to Graph
// /me: Graph only returns 200 for a genuine, unexpired token it issued.
// Once that succeeds, the token as a whole is known-authentic, so it's
// safe to also read auxiliary fields (like tid) off the unsigned decode —
// an attacker can't get Graph to accept a token whose payload was tampered.
async function validateUserToken(request) {
  const auth  = request.headers.get('Authorization') || '';
  const token = auth.startsWith('Bearer ') ? auth.slice(7) : '';
  if (!token) return { error: 'Missing authorization token', status: 401 };

  let me;
  try {
    const meRes = await fetch(
      'https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName,displayName',
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!meRes.ok) return { error: 'Invalid or expired token', status: 401 };
    me = await meRes.json();
  } catch {
    return { error: 'Could not verify token', status: 502 };
  }

  const decoded = decodeToken(token) || {};
  if (decoded.tid !== DPEG_TENANT_ID) return { error: 'Wrong tenant', status: 403 };

  const email = extractEmailAddress(me.mail || me.userPrincipalName || '');
  if (!email) return { error: 'Could not resolve account email', status: 403 };

  const claims = { ...decoded, preferred_username: email, name: me.displayName || decoded.name };
  return { claims };
}

// Acquire an app-only token using client_credentials (uses Tasks.ReadWrite.All Application permission).
// Cached per-isolate until shortly before expiry so we're not round-tripping
// to Azure AD on every single request.
let appTokenCache = { token: null, expiresAt: 0 };
async function getAppToken(env) {
  const now = Date.now();
  if (appTokenCache.token && now < appTokenCache.expiresAt - 30_000) {
    return appTokenCache.token;
  }
  const body = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     AZURE_CLIENT_ID,
    client_secret: env.AZURE_CLIENT_SECRET,
    scope:         'https://graph.microsoft.com/.default',
  });
  const res = await fetch(
    `https://login.microsoftonline.com/${DPEG_TENANT_ID}/oauth2/v2.0/token`,
    { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: body.toString() }
  );
  if (!res.ok) {
    const err = await res.text().catch(() => '');
    throw new Error(`App token request failed (${res.status}): ${err}`);
  }
  const data = await res.json();
  appTokenCache = { token: data.access_token, expiresAt: now + (data.expires_in || 3600) * 1000 };
  return appTokenCache.token;
}

// ── /todo endpoint ───────────────────────────────────────────────────────────
async function handleTodo(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const { recipientEmail, title, summary = '', priority = 'Normal', date, deadline, appTaskId = '' } = body;
  const recipient = extractEmailAddress(recipientEmail);
  const assignedByEmail = extractEmailAddress(body.assignedByEmail || userEmailFromClaims(claims));
  const assignedByName = String(body.assignedByName || claims.name || assignedByEmail || '').trim();
  if (!recipient || !title) {
    return json({ error: 'recipientEmail and title are required' }, 400);
  }
  if (!recipient.includes('@dhananipeg.com')) {
    return json({ error: 'Only @dhananipeg.com addresses are supported' }, 403);
  }

  // Get app-level token (Tasks.ReadWrite.All Application permission)
  let appToken;
  try { appToken = await getAppToken(env); }
  catch (err) { return json({ error: 'Could not acquire app token', detail: err.message }, 502); }

  // Find or create recipient list grouped by assigner
  const assignerLabel = String(assignedByName || assignedByEmail || 'DPEG Manager').trim();
  const desiredListName = `Tasks from ${assignerLabel}`;
  const listsRes = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(recipient)}/todo/lists`,
    { headers: { Authorization: `Bearer ${appToken}` } }
  );
  if (!listsRes.ok) {
    const err = await listsRes.text().catch(() => '');
    return json({ error: 'Cannot access recipient To Do', detail: err }, listsRes.status);
  }
  const listsData = await listsRes.json();
  let defaultList = (listsData.value || []).find(l => l.displayName === desiredListName);

  // If this assigner list does not exist, create it.
  if (!defaultList) {
    const createRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(recipient)}/todo/lists`,
      { method: 'POST', headers: { Authorization: `Bearer ${appToken}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ displayName: desiredListName }) }
    );
    if (!createRes.ok) {
      const err = await createRes.text().catch(() => '');
      return json({ error: 'Could not find or create task list for recipient', detail: err }, 502);
    }
    defaultList = await createRes.json();
  }

  // Build the To Do task
  const cleanSummary = summary.replace(/[•*▾▲◆]/g, '').slice(0, 1200);
  const task = {
    title,
    body: {
      content: [
        `Assigned by: ${assignerLabel}${assignedByEmail ? ` <${assignedByEmail}>` : ''}`,
        appTaskId ? `DPEG Task ID: ${appTaskId}` : '',
        '',
        cleanSummary,
        '',
        'Proof upload link will appear here after this task is created.',
      ].filter(Boolean).join('\n'),
      contentType: 'text',
    },
    importance: String(priority).toLowerCase() === 'high' ? 'high' : 'normal',
    status: 'notStarted',
  };
  const due = graphDueDate(deadline || date);
  if (due) task.dueDateTime = due;

  const taskRes = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(recipient)}/todo/lists/${defaultList.id}/tasks`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${appToken}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(task),
    }
  );
  if (!taskRes.ok) {
    const err = await taskRes.text().catch(() => '');
    return json({ error: 'Failed to create To Do task', detail: err }, taskRes.status);
  }

  const taskData = await taskRes.json().catch(() => ({}));
  if (taskData.id) {
    const longLink = proofSubmitUrl(body, defaultList.id, taskData.id);
    const link = await createProofShortUrl(request, env, longLink);
    const bodyHtml = [
      `<p><b>Assigned by:</b> ${esc(assignerLabel)}${assignedByEmail ? ` (${esc(assignedByEmail)})` : ''}</p>`,
      appTaskId ? `<p><b>Task ID:</b> ${esc(appTaskId)}</p>` : '',
      cleanSummary ? `<p>${esc(cleanSummary).replace(/\n/g, '<br>')}</p>` : '',
      `<p><b>Proof Submission</b><br><a href="${esc(link)}">Submit Proof</a></p>`,
    ].filter(Boolean).join('\n');
    await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(recipient)}/todo/lists/${defaultList.id}/tasks/${taskData.id}`,
      {
        method: 'PATCH',
        headers: { Authorization: `Bearer ${appToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ body: { content: bodyHtml, contentType: 'html' } }),
      }
    ).catch(() => {});
  }
  return json({ success: true, listId: defaultList.id, taskId: taskData.id || null });
}

// ── /data endpoint: shared company state for Action Log / Wednesday / Admin config
async function handleData(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_DATA) return json({ error: 'DPEG_DATA KV binding is not configured' }, 501);

  if (request.method === 'GET') {
    const data = await env.DPEG_DATA.get(DATA_KEY, 'json');
    return json(data || { tasks: [], archives: [], staffConfig: {}, customNotes: [], notifications: [] });
  }

  if (request.method === 'PUT' || request.method === 'POST') {
    let body;
    try { body = await request.json(); }
    catch { return json({ error: 'Invalid JSON body' }, 400); }

    const userEmail = extractEmailAddress(claims.preferred_username || claims.upn || claims.email || '');
    const existing = await env.DPEG_DATA.get(DATA_KEY, 'json') || {};

    // Optimistic concurrency: this endpoint reads the whole shared document
    // and writes the whole thing back, so without a version check, two
    // people saving around the same time silently clobber each other's
    // tasks/notes. The client must prove it last loaded the version it's
    // about to overwrite; otherwise it must reload and retry.
    if (existing.updatedAt && body.baseUpdatedAt !== existing.updatedAt) {
      return json({
        error: 'conflict',
        message: 'Company data changed since you last loaded it. Reload and try again.',
        current: existing,
      }, 409);
    }

    const payload = {
      tasks: Array.isArray(body.tasks) ? body.tasks : [],
      archives: Array.isArray(body.archives) ? body.archives : [],
      staffConfig: ADMIN_EMAILS.has(userEmail) && body.staffConfig && typeof body.staffConfig === 'object'
        ? body.staffConfig
        : (existing.staffConfig && typeof existing.staffConfig === 'object' ? existing.staffConfig : {}),
      customDepartments: Array.isArray(existing.customDepartments) ? existing.customDepartments : [],
      departmentAssignments: existing.departmentAssignments && typeof existing.departmentAssignments === 'object'
        ? existing.departmentAssignments : {},
      departmentsUpdatedAt: existing.departmentsUpdatedAt || null,
      customNotes: Array.isArray(body.customNotes) ? body.customNotes : [],
      notifications: Array.isArray(body.notifications) ? body.notifications : [],
      updatedAt: new Date().toISOString(),
    };
    await env.DPEG_DATA.put(DATA_KEY, JSON.stringify(payload));
    return json({ success: true, updatedAt: payload.updatedAt });
  }

  return json({ error: 'Method not allowed' }, 405);
}

async function sharedStorageFlag(env, name, fallback = 'off') {
  if (!env.DPEG_ASSIGNMENTS) return fallback;
  try {
    const row = await env.DPEG_ASSIGNMENTS.prepare(
      'SELECT value FROM feature_flags WHERE name = ?'
    ).bind(name).first();
    return String(row?.value || fallback).trim().toLowerCase();
  } catch {
    // Older deployments may not have the additive schema yet. A missing flag
    // must always fail closed rather than accidentally activating migration.
    return fallback;
  }
}

function normalizedShadowTask(task, ownerEmail, now) {
  const id = String(task?.id || '').trim().slice(0, 200);
  const title = String(task?.title || '').trim().slice(0, 500);
  if (!id || !title) return null;
  const rawStatus = String(task?.status || '').trim().toLowerCase();
  const status = rawStatus === 'done' || rawStatus === 'completed' ? 'Done'
    : rawStatus === 'cancelled' || rawStatus === 'canceled' ? 'Cancelled'
      : rawStatus === 'in progress' || rawStatus === 'inprogress' ? 'In Progress'
        : 'Pending';
  const createdAt = String(task?.createdAt || task?.date || now).trim().slice(0, 80) || now;
  const updatedAt = String(task?.updatedAt || task?.completedAt || task?.cancelledAt || now).trim().slice(0, 80) || now;
  return {
    id,
    ownerEmail,
    title,
    summary: String(task?.summary || task?.description || '').slice(0, 12000),
    departmentName: String(task?.dept || task?.department || 'Needs Department').trim().slice(0, 200) || 'Needs Department',
    priority: String(task?.priority || 'Normal').trim().slice(0, 40) || 'Normal',
    dueDate: String(task?.deadline || task?.dueDate || '').trim().slice(0, 80) || null,
    status,
    createdAt,
    updatedAt,
    completedAt: status === 'Done' ? String(task?.completedAt || task?.approvedAt || updatedAt).slice(0, 80) : null,
    cancelledAt: status === 'Cancelled' ? String(task?.cancelledAt || updatedAt).slice(0, 80) : null,
    version: Math.max(1, Number(task?.version || task?.assignmentVersion || 1) || 1),
  };
}

// Shadow-copy the signed-in user's legacy OneDrive tasks into normalized D1.
// The feature flag fails closed, OneDrive remains primary, and no rows are
// deleted. This endpoint is safe to deploy before dual-write is enabled.
async function handleSharedWorkflowSync(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (request.method !== 'POST') return json({ error: 'Method not allowed' }, 405);
  if (!env.DPEG_ASSIGNMENTS) return json({ error: 'D1 storage is not configured' }, 501);

  const flag = await sharedStorageFlag(env, 'shared_storage_dual_write');
  if (flag !== 'on') return json({ success: true, enabled: false, written: 0, skipped: 0 });

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }
  const incoming = Array.isArray(body.tasks) ? body.tasks.slice(0, 100) : [];
  const ownerEmail = userEmailFromClaims(claims);
  const ownedAssignments = await env.DPEG_ASSIGNMENTS.prepare(
    'SELECT DISTINCT app_task_id FROM assignments WHERE lower(assigner_email) = ?'
  ).bind(ownerEmail).all();
  const ownedIds = new Set((ownedAssignments.results || []).map(row => String(row.app_task_id || '')));
  const now = new Date().toISOString();
  const rows = [];
  let skipped = 0;

  for (const task of incoming) {
    const taskId = String(task?.id || '').trim();
    const declaredOwner = extractEmailAddress(task?.assignedByEmail || task?.assignerEmail || '');
    // Explicit ownership by somebody else always wins. Legacy/manual tasks
    // without an assigner are personal to the signed-in OneDrive owner.
    if (declaredOwner && declaredOwner !== ownerEmail && !ownedIds.has(taskId)) {
      skipped += 1;
      continue;
    }
    const row = normalizedShadowTask(task, ownerEmail, now);
    if (!row) { skipped += 1; continue; }
    rows.push(row);
  }

  const statements = rows.map(row => env.DPEG_ASSIGNMENTS.prepare(
    `INSERT INTO tasks
       (id, owner_email, title, summary, department_name, priority, due_date,
        status, source_type, created_at, updated_at, completed_at, cancelled_at, version)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'legacy_onedrive_shadow', ?, ?, ?, ?, ?)
     ON CONFLICT(id) DO UPDATE SET
       title = excluded.title,
       summary = excluded.summary,
       department_name = excluded.department_name,
       priority = excluded.priority,
       due_date = excluded.due_date,
       status = excluded.status,
       updated_at = excluded.updated_at,
       completed_at = excluded.completed_at,
       cancelled_at = excluded.cancelled_at,
       version = MAX(tasks.version, excluded.version)
     WHERE lower(tasks.owner_email) = lower(excluded.owner_email)`
  ).bind(
    row.id, row.ownerEmail, row.title, row.summary, row.departmentName,
    row.priority, row.dueDate, row.status, row.createdAt, row.updatedAt,
    row.completedAt, row.cancelledAt, row.version,
  ));
  const results = statements.length ? await env.DPEG_ASSIGNMENTS.batch(statements) : [];
  const written = results.reduce((sum, result) => sum + Number(result.meta?.changes || 0), 0);
  return json({ success: true, enabled: true, received: incoming.length, written, skipped });
}

// A short presence lock prevents two people from editing the same task form
// at once. It expires automatically if a browser closes without releasing it.
async function handleTaskEditLock(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_DATA) return json({ error: 'DPEG_DATA KV binding is not configured' }, 501);
  let body = {};
  if (request.method !== 'GET') {
    try { body = await request.json(); }
    catch { return json({ error: 'Invalid JSON body' }, 400); }
  }
  const url = new URL(request.url);
  const taskId = String(body.taskId || url.searchParams.get('taskId') || '').trim().slice(0, 160);
  if (!taskId) return json({ error: 'taskId is required' }, 400);
  const key = `task-edit-lock:${taskId}`;
  const email = userEmailFromClaims(claims);
  const now = Date.now();
  const existing = await env.DPEG_DATA.get(key, 'json');
  if (request.method === 'GET') return json({ lock: existing && Number(existing.expiresAt || 0) > now ? existing : null });
  if (request.method === 'DELETE') {
    if (existing && extractEmailAddress(existing.email) === email) await env.DPEG_DATA.delete(key);
    return json({ success: true });
  }
  if (request.method !== 'POST') return json({ error: 'Method not allowed' }, 405);
  if (existing && Number(existing.expiresAt || 0) > now && extractEmailAddress(existing.email) !== email) {
    return json({ error: 'currently_editing', editorName: existing.name || existing.email, expiresAt: existing.expiresAt }, 423);
  }
  const lock = { taskId, email, name: String(claims.name || email), expiresAt: now + 120000 };
  await env.DPEG_DATA.put(key, JSON.stringify(lock), { expirationTtl: 120 });
  let version=null;
  if(env.DPEG_ASSIGNMENTS){
    const row=await env.DPEG_ASSIGNMENTS.prepare('SELECT version FROM assignments WHERE id = ?').bind(taskId).first();
    if(row)version=Number(row.version||1);
  }
  return json({ success: true, lock, version });
}

// Shared department registry. Microsoft remains the contact/name source; this
// stores only company department names and email-to-department overrides.
async function handleDepartments(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_DATA) return json({ error: 'DPEG_DATA KV binding is not configured' }, 501);
  const existing = await env.DPEG_DATA.get(DATA_KEY, 'json') || {};

  if (request.method === 'GET') {
    return json({
      departments: Array.isArray(existing.customDepartments) ? existing.customDepartments : [],
      assignments: existing.departmentAssignments && typeof existing.departmentAssignments === 'object'
        ? existing.departmentAssignments : {},
      updatedAt: existing.departmentsUpdatedAt || null,
    });
  }

  if (request.method !== 'PUT' && request.method !== 'POST') {
    return json({ error: 'Method not allowed' }, 405);
  }
  const userEmail = extractEmailAddress(claims.preferred_username || claims.upn || claims.email || '');
  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  if (!ADMIN_EMAILS.has(userEmail)) return json({ error: 'Admin access only' }, 403);

  // Admins may map one verified Microsoft email to an existing department.
  if (request.method === 'POST' && body.assignment) {
    const email = extractEmailAddress(body.assignment.email || '');
    const name = String(body.assignment.name || '').trim().replace(/\s+/g, ' ');
    const dept = String(body.assignment.dept || '').trim().replace(/\s+/g, ' ');
    const departments = Array.isArray(existing.customDepartments) ? existing.customDepartments : [];
    const validDepartment = STANDARD_DEPARTMENTS.has(dept.toLowerCase()) ||
      departments.some(value => String(value || '').trim().toLowerCase() === dept.toLowerCase());
    if (!email || (!email.endsWith('@dhananipeg.com') && !email.endsWith('@dpeg.com'))) {
      return json({ error: 'A valid internal Microsoft email is required' }, 400);
    }
    if (!dept || dept === 'Needs Department' || !validDepartment) {
      return json({ error: 'Select an existing department' }, 400);
    }
    const assignments = existing.departmentAssignments && typeof existing.departmentAssignments === 'object'
      ? { ...existing.departmentAssignments } : {};
    assignments[email] = { email, name, dept };
    const now = new Date().toISOString();
    const payload = { ...existing, departmentAssignments: assignments, departmentsUpdatedAt: now };
    await env.DPEG_DATA.put(DATA_KEY, JSON.stringify(payload));
    if (env.DPEG_ASSIGNMENTS) {
      await env.DPEG_ASSIGNMENTS.prepare(
        'UPDATE assignments SET dept = ?, updated_at = ?, version = version + 1 WHERE lower(recipient_email) = ?'
      ).bind(dept, now, email).run();
    }
    return json({ success: true, assignment: assignments[email], updatedAt: now });
  }

  if (existing.departmentsUpdatedAt && body.baseUpdatedAt !== existing.departmentsUpdatedAt) {
    return json({ error: 'conflict', message: 'Departments changed since they were loaded.' }, 409);
  }

  const seen = new Set();
  const departments = (Array.isArray(body.departments) ? body.departments : [])
    .map(value => String(value || '').trim().replace(/\s+/g, ' '))
    .filter(value => value && !seen.has(value.toLowerCase()) && seen.add(value.toLowerCase()))
    .slice(0, 200);
  const assignments = {};
  for (const row of Object.values(body.assignments && typeof body.assignments === 'object' ? body.assignments : {})) {
    const email = extractEmailAddress(row?.email || '');
    const dept = String(row?.dept || '').trim();
    if (!email || !dept || !email.endsWith('@dhananipeg.com')) continue;
    assignments[email] = { email, name: String(row?.name || '').trim(), dept };
  }
  const now = new Date().toISOString();
  const payload = { ...existing, customDepartments: departments, departmentAssignments: assignments, departmentsUpdatedAt: now };
  await env.DPEG_DATA.put(DATA_KEY, JSON.stringify(payload));
  return json({ success: true, updatedAt: now });
}

// ── / endpoint (existing AI summary) ─────────────────────────────────────────
async function handleSummary(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const { subject = '', emailText = '', senderName = '', messageCount = 1, latestMessageText = '', latestSender = '', latestDate = '', attachmentNames = [] } = body;
  if (!emailText && !subject) {
    return json({ error: 'Provide emailText or subject' }, 400);
  }

  const attLine = attachmentNames.length
    ? `\nATTACHMENTS (${attachmentNames.length}): ${attachmentNames.join(', ')}. Do not analyze attachment content — only acknowledge that attachments exist if relevant.`
    : '';

  // Extract emailDate note if present in emailText (appended by client)
  const emailDateMatch = emailText.match(/\[TASK CONTEXT\][^\n]*email was received on ([^.]+)\. The task is being assigned today: ([^.]+)\./);
  const emailDateNote = emailDateMatch ? `Note: this email was originally sent on ${emailDateMatch[1].trim()}.` : '';

  const prompt = `You are an executive assistant at DPEG (Dhanani Private Equity Group). Summarise this email in 2-3 sentences. Focus on what action is needed, who needs to do it, and any deadline or amount mentioned.${emailDateNote ? ' ' + emailDateNote : ''} Be clear and concise. Do not reference any attachments — only the email body text.${attLine ? '\n\nATTACHMENT NOTE: ' + attLine : ''}

Subject: "${subject}"${senderName ? `\nFrom: ${senderName}` : ''}
${messageCount > 1 ? `Thread: ${messageCount} messages` : ''}

EMAIL BODY (summarise this only):
${emailText.replace(/\[TASK CONTEXT\][\s\S]*$/, '').trim().slice(0, 2800)}
${latestMessageText ? `\nLATEST MESSAGE${latestSender ? ` from ${latestSender}` : ''}${latestDate ? ` (${latestDate})` : ''}:\n${latestMessageText.slice(0, 700)}` : ''}

Write 3-5 sentences. No bullet points, no headers. State names, amounts, properties, and deadlines explicitly.`;

  const groqRes = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${env.GROQ_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'llama-3.3-70b-versatile',
      messages: [{ role: 'user', content: prompt }],
      max_tokens: 350,
      temperature: 0.1,
    }),
  });

  if (!groqRes.ok) {
    const err = await groqRes.text().catch(() => '');
    return json({ error: 'Groq call failed', detail: err }, 502);
  }

  const groqData = await groqRes.json();
  const summary = groqData.choices?.[0]?.message?.content?.trim() || '';
  return json({ summary });
}

// ── /attachment-summary endpoint ─────────────────────────────────────────────
async function handleAttachmentSummary(request, env) {
  const { error, status } = await validateUserToken(request);
  if (error) return json({ error }, status);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const { subject = '', attachmentContents = [] } = body;
  if (!attachmentContents.length) return json({ error: 'No attachment contents provided' }, 400);

  const text = attachmentContents.map(a => `[${a.name}]\n${String(a.text || '').slice(0, 800)}`).join('\n\n---\n\n');

  const prompt = `You are an executive assistant at DPEG (Dhanani Private Equity Group). Summarize the content of the following email attachments.

Email subject: "${subject}"

ATTACHMENT CONTENTS:
${text.slice(0, 3000)}

Write 1-3 bullet points (•) summarizing what these attachments contain. Be specific — name figures, dates, property addresses, and key details. Max 150 words total. Do not include preamble or labels.`;

  const groqRes = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${env.GROQ_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'llama-3.3-70b-versatile',
      messages: [{ role: 'user', content: prompt }],
      max_tokens: 250,
      temperature: 0.1,
    }),
  });

  if (!groqRes.ok) {
    const err = await groqRes.text().catch(() => '');
    return json({ error: 'Groq call failed', detail: err }, 502);
  }

  const groqData = await groqRes.json();
  const summary = groqData.choices?.[0]?.message?.content?.trim() || '';
  return json({ summary });
}

// ── /poll-completions endpoint ───────────────────────────────────────────────
async function handlePollCompletions(request, env) {
  const { error, status } = await validateUserToken(request);
  if (error) return json({ error }, status);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const { assignments = [] } = body;
  if (!assignments.length) return json({ completed: [] });

  let appToken;
  try { appToken = await getAppToken(env); }
  catch (err) { return json({ error: 'Could not acquire app token', detail: err.message }, 502); }

  const completed = [];
  for (const a of assignments) {
    const { recipientEmail, todoListId, todoTaskId, taskId } = a;
    if (!recipientEmail || !todoListId || !todoTaskId) continue;
    if (!recipientEmail.includes('@dhananipeg.com')) continue;
    try {
      const res = await fetch(
        `${todoTaskUrl(recipientEmail, todoListId, todoTaskId)}?$select=id,status,completedDateTime,body`,
        { headers: { Authorization: `Bearer ${appToken}` } }
      );
      if (!res.ok) continue;
      const taskData = await res.json();
      if (taskData.status === 'completed') {
        completed.push({
          taskId,
          todoTaskId,
          recipientEmail,
          completedDateTime: taskData.completedDateTime?.dateTime || null,
          proofs: parseProofs(taskData.body?.content || ''),
        });
      }
    } catch {}
  }

  return json({ completed });
}

async function handleProofTask(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const recipientEmail = extractEmailAddress(body.recipientEmail || '');
  const { todoListId, todoTaskId } = body;
  if (!recipientEmail || !todoListId || !todoTaskId) return json({ error: 'Missing proof task details' }, 400);
  if (!recipientEmail.includes('@dhananipeg.com')) return json({ error: 'Only @dhananipeg.com task recipients are supported' }, 403);

  let appToken;
  try { appToken = await getAppToken(env); }
  catch (err) { return json({ error: 'Could not acquire app token', detail: err.message }, 502); }

  const res = await fetch(
    `${todoTaskUrl(recipientEmail, todoListId, todoTaskId)}?$select=id,title,body,status`,
    { headers: { Authorization: `Bearer ${appToken}` } }
  );
  if (!res.ok) {
    const detail = await res.text().catch(() => '');
    return json({ error: 'Could not load assigned To Do task', detail }, res.status);
  }
  const task = await res.json();
  return json({ title: task.title || '', status: task.status || '', proofs: parseProofs(task.body?.content || '') });
}

async function handleProofSubmit(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const recipientEmail = extractEmailAddress(body.recipientEmail || '');
  const { todoListId, todoTaskId, proofs = [], markDone = true } = body;
  if (!recipientEmail || !todoListId || !todoTaskId) return json({ error: 'Missing proof task details' }, 400);
  if (!recipientEmail.includes('@dhananipeg.com')) return json({ error: 'Only @dhananipeg.com task recipients are supported' }, 403);
  if (!Array.isArray(proofs) || !proofs.length) return json({ error: 'No proof files provided' }, 400);

  let appToken;
  try { appToken = await getAppToken(env); }
  catch (err) { return json({ error: 'Could not acquire app token', detail: err.message }, 502); }

  const taskUrl = todoTaskUrl(recipientEmail, todoListId, todoTaskId);
  const currentRes = await fetch(`${taskUrl}?$select=id,body,status`, { headers: { Authorization: `Bearer ${appToken}` } });
  if (!currentRes.ok) {
    const detail = await currentRes.text().catch(() => '');
    return json({ error: 'Could not load assigned To Do task', detail }, currentRes.status);
  }
  const task = await currentRes.json();
  const parsed = parseProofBlock(task.body?.content || '');
  const nextProofs = [...parsed.proofs, ...proofs];
  const patch = {
    body: { content: buildProofBlock(parsed.base, nextProofs), contentType: 'html' },
  };
  if (markDone) patch.status = 'completed';

  const patchRes = await fetch(taskUrl, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${appToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(patch),
  });
  if (!patchRes.ok) {
    const detail = await patchRes.text().catch(() => '');
    return json({ error: 'Could not update assigned To Do task', detail }, patchRes.status);
  }
  return json({ success: true, proofs: nextProofs });
}

// ── /notify endpoint: append or update proof notifications in KV ──────────────
async function handleNotify(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_DATA) return json({ error: 'DPEG_DATA KV binding is not configured' }, 501);

  if (request.method === 'GET') {
    const data = await env.DPEG_DATA.get(DATA_KEY, 'json') || {};
    const legacyNotifications = Array.isArray(data.notifications) ? data.notifications : [];
    const url = new URL(request.url);
    const includeMessages = url.searchParams.get('includeMessages') === '1';
    const taskId = url.searchParams.get('taskId') || '';
    const recipientEmail = url.searchParams.get('recipientEmail') || '';
    const messagesSince = url.searchParams.get('messagesSince') || '';
    const wantsMessages = includeMessages || Boolean(taskId) || Boolean(messagesSince);
    const d1Threads = wantsMessages
      ? await loadTaskMessageThreads(env, claims, { taskId, recipientEmail, since: messagesSince })
      : [];
    const messageCursor = d1Threads.reduce((latest, thread) => {
      const value = String(thread.updatedAt || '');
      return value > latest ? value : latest;
    }, messagesSince);
    return json({
      notifications: wantsMessages ? mergeD1TaskThreads(legacyNotifications, d1Threads) : legacyNotifications,
      messageCursor,
    });
  }

  if (request.method !== 'POST') return json({ error: 'Method not allowed' }, 405);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const data = await env.DPEG_DATA.get(DATA_KEY, 'json') || {};
  const notifications = Array.isArray(data.notifications) ? data.notifications : [];

  if (body.type === 'proof_result') {
    // Mark the original proof_submitted notification as resolved
    const idx = notifications.findIndex(n => n.id === body.notifId && n.type === 'proof_submitted');
    if (idx >= 0) notifications[idx].status = body.result === 'approved' ? 'approved' : 'declined';
    // Add a result notification for the recipient
    notifications.push({
      id: `pr-${Date.now()}-${Math.random().toString(36).slice(2,7)}`,
      type: 'proof_result',
      appTaskId: String(body.appTaskId || ''),
      taskTitle: String(body.taskTitle || ''),
      senderEmail: String(body.senderEmail || ''),
      senderName: String(body.senderName || ''),
      recipientEmail: String(body.recipientEmail || ''),
      result: body.result === 'approved' ? 'approved' : 'declined',
      reason: String(body.reason || ''),
      createdAt: new Date().toISOString(),
      seen: false,
    });
    await updateAssignmentProofState(env, {
      appTaskId: body.appTaskId,
      recipientEmail: body.recipientEmail,
      proofStatus: body.result === 'approved' ? 'approved' : 'declined',
      notificationId: body.notifId,
    });
    // If declined, reset the recipient's To Do task back to notStarted
    if (externalEffectsAllowed(env) && body.result === 'declined' && body.todoListId && body.todoTaskId && body.recipientEmail) {
      try {
        const appToken = await getAppToken(env);
        const recipientEmail = extractEmailAddress(body.recipientEmail);
        if (recipientEmail.includes('@dhananipeg.com')) {
          await fetch(
            `${todoTaskUrl(recipientEmail, body.todoListId, body.todoTaskId)}`,
            {
              method: 'PATCH',
              headers: { Authorization: `Bearer ${appToken}`, 'Content-Type': 'application/json' },
              body: JSON.stringify({ status: 'notStarted' }),
            }
          );
        }
      } catch {}
    }
  } else if (body.type === 'proof_submitted') {
    const proofNotificationId = `pn-${Date.now()}-${Math.random().toString(36).slice(2,7)}`;
    notifications.push({
      id: proofNotificationId,
      type: 'proof_submitted',
      appTaskId: String(body.appTaskId || ''),
      taskTitle: String(body.taskTitle || ''),
      senderEmail: String(body.senderEmail || ''),
      recipientEmail: String(body.recipientEmail || ''),
      recipientName: String(body.recipientName || ''),
      proofs: Array.isArray(body.proofs) ? body.proofs : [],
      note: String(body.note || ''),
      thread: [],
      followupStatus: '',
      submittedAt: new Date().toISOString(),
      status: 'pending',
      seen: false,
    });
    await updateAssignmentProofState(env, {
      appTaskId: body.appTaskId,
      recipientEmail: body.recipientEmail,
      senderEmail: body.senderEmail,
      proofStatus: 'submitted',
      notificationId: proofNotificationId,
    });
  } else if (body.type === 'proof_followup_question') {
    const idx = notifications.findIndex(n => n.id === body.notifId && n.type === 'proof_submitted' && n.status === 'pending');
    if (idx < 0) return json({ error: 'Proof notification not found' }, 404);
    const question = String(body.message || '').trim();
    if (!question) return json({ error: 'Question is required' }, 400);
    const thread = Array.isArray(notifications[idx].thread) ? notifications[idx].thread : [];
    thread.push({
      id: `fq-${Date.now()}-${Math.random().toString(36).slice(2,7)}`,
      by: 'assignor',
      email: String(body.senderEmail || userEmailFromClaims(claims)),
      name: String(body.senderName || claims.name || ''),
      message: question,
      createdAt: new Date().toISOString(),
    });
    notifications[idx].thread = thread;
    notifications[idx].followupStatus = 'question';
    notifications[idx].updatedAt = new Date().toISOString();
  } else if (body.type === 'proof_followup_answer') {
    const recipientEmail = extractEmailAddress(body.recipientEmail || userEmailFromClaims(claims));
    const idx = notifications.findIndex(n =>
      n.type === 'proof_submitted' &&
      n.status === 'pending' &&
      (String(n.id) === String(body.notifId || '') ||
        (String(n.appTaskId) === String(body.appTaskId || '') && extractEmailAddress(n.recipientEmail) === recipientEmail))
    );
    if (idx < 0) return json({ error: 'Follow-up question not found' }, 404);
    const answer = String(body.message || '').trim();
    const attachments = Array.isArray(body.attachments) ? body.attachments : [];
    if (!answer && !attachments.length) return json({ error: 'Answer or attachment is required' }, 400);
    const thread = Array.isArray(notifications[idx].thread) ? notifications[idx].thread : [];
    thread.push({
      id: `fa-${Date.now()}-${Math.random().toString(36).slice(2,7)}`,
      by: 'assignee',
      email: recipientEmail,
      name: String(body.recipientName || claims.name || ''),
      message: answer,
      attachments,
      createdAt: new Date().toISOString(),
    });
    notifications[idx].thread = thread;
    notifications[idx].followupStatus = 'answered';
    notifications[idx].updatedAt = new Date().toISOString();
  } else if (body.type === 'task_followup_message') {
    // Generic per-task follow-up, usable from the Tasks tab regardless of
    // proof status — unlike proof_followup_question/answer above, this does
    // not require a prior proof submission to exist.
    const inserted = await insertTaskMessage(env, body, claims);
    if (inserted.error) return inserted.error;
    // Do not rewrite company-state here. That shared KV read/modify/write was
    // the source of lost simultaneous messages.
    return json({ success: true, messageId: inserted.id, createdAt: inserted.createdAt });
  } else {
    return json({ error: 'Unknown notification type' }, 400);
  }

  data.notifications = notifications;
  await env.DPEG_DATA.put(DATA_KEY, JSON.stringify(data));
  return json({ success: true });
}

// ── Update recipient To Do task (preserves proof block) ──────────────────────
async function handleTodoUpdate(request, env) {
  let body;
  try { body = await request.json(); } catch { return json({ error: 'Invalid JSON' }, 400); }
  const { recipientEmail, todoListId, todoTaskId, title, priority, date, followupNote, changes, assignedByName } = body;
  if (!recipientEmail || !todoListId || !todoTaskId) {
    return json({ error: 'recipientEmail, todoListId and todoTaskId are required' }, 400);
  }
  const recipient = extractEmailAddress(recipientEmail);
  if (!recipient.includes('@dhananipeg.com')) return json({ error: 'Only @dhananipeg.com supported' }, 403);
  let appToken;
  try { appToken = await getAppToken(env); }
  catch (err) { return json({ error: 'Could not acquire app token', detail: err.message }, 502); }

  // Fetch existing body — preserve HTML if already HTML, convert text to HTML if needed
  const taskRes = await fetch(
    `${todoTaskUrl(recipient, todoListId, todoTaskId)}?$select=id,body`,
    { headers: { Authorization: `Bearer ${appToken}` } }
  );
  let baseHtml = '';
  if (taskRes.ok) {
    const td = await taskRes.json().catch(() => ({}));
    const { base } = parseProofBlock(td.body?.content || '');
    if ((td.body?.contentType || 'text') === 'html') {
      baseHtml = base;
    } else {
      // Convert plain text to minimal HTML so appended updates stay structured
      baseHtml = base ? '<p>' + esc(base).replace(/\n\n+/g, '</p><p>').replace(/\n/g, '<br>') + '</p>' : '';
    }
  }

  // Filter out "Note: ..." entries from changes when followupNote is sent separately
  const significantChanges = (changes || []).filter(c => !followupNote || !String(c).startsWith('Note: '));
  const extrasHtml = [
    significantChanges.length ? `<p><b>Updated:</b> ${esc(significantChanges.join(' | '))}</p>` : '',
    followupNote ? `<p><b>Follow-up note from manager:</b><br>${esc(followupNote).replace(/\n/g, '<br>')}</p>` : '',
  ].filter(Boolean).join('\n');
  const newContent = [baseHtml, extrasHtml].filter(Boolean).join('\n');

  const due = graphDueDate(date);
  const patch = {
    body: { content: newContent, contentType: 'html' },
    importance: String(priority || '').toLowerCase() === 'high' ? 'high' : 'normal',
  };
  if (due) patch.dueDateTime = due;
  if (followupNote) patch.status = 'notStarted';

  const patchRes = await fetch(todoTaskUrl(recipient, todoListId, todoTaskId), {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${appToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(patch),
  });
  if (!patchRes.ok) {
    const err = await patchRes.text().catch(() => '');
    return json({ error: 'Failed to update recipient To Do', detail: err }, patchRes.status);
  }
  return json({ success: true });
}

// ── /assignment endpoint: create/update a shared assignment record (D1) ──────
async function handleCreateAssignment(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_ASSIGNMENTS) return json({ error: 'DPEG_ASSIGNMENTS D1 binding is not configured' }, 501);
  await ensureAssignmentProofColumns(env);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const id = String(body.id || '').trim();
  const title = String(body.title || '').trim();
  const recipientEmail = extractEmailAddress(body.recipientEmail || '');
  if (!id || !title || !recipientEmail) {
    return json({ error: 'id, title and recipientEmail are required' }, 400);
  }
  if (!recipientEmail.includes('@dhananipeg.com')) {
    return json({ error: 'Only @dhananipeg.com addresses are supported' }, 403);
  }

  const assignerEmail = userEmailFromClaims(claims);
  const assignerName = String(body.assignerName || claims.name || assignerEmail || '').trim();
  const now = new Date().toISOString();
  const existingRow = await env.DPEG_ASSIGNMENTS.prepare(
    'SELECT assigner_email, version FROM assignments WHERE id = ?'
  ).bind(id).first();
  if(existingRow&&extractEmailAddress(existingRow.assigner_email)!==assignerEmail){
    return json({error:'Only the assigner can edit this task'},403);
  }
  if(existingRow&&body.expectedVersion!=null&&Number(body.expectedVersion)!==Number(existingRow.version||1)){
    return json({error:'version_conflict',message:'This task was updated by another user. Reload it before making changes.',currentVersion:Number(existingRow.version||1)},409);
  }

  // Single atomic upsert: on conflict, only touch assigner-owned columns.
  // Recipient-owned columns (status, progress_note, proof_*) and created_at
  // are deliberately left out of the DO UPDATE SET so a concurrent
  // /assignment-status write from the recipient can never be clobbered by
  // a racing /assignment write from the assigner (or vice versa).
  await env.DPEG_ASSIGNMENTS.prepare(
    `INSERT INTO assignments
      (id, app_task_id, title, summary, dept, priority, due_date,
       assigner_email, assigner_name, recipient_email, recipient_name,
       status, progress_note, proof_status, proof_submitted_at, proof_reviewed_at, proof_notification_id, recipient_todo_list_id, recipient_todo_task_id,
       proof_instructions, created_at, updated_at)
     VALUES (?,?,?,?,?,?,?,?,?,?,?,'Assigned',NULL,'none',NULL,NULL,NULL,?,?,?,?,?)
     ON CONFLICT(id) DO UPDATE SET
       app_task_id = excluded.app_task_id,
       title = excluded.title,
       summary = excluded.summary,
       dept = excluded.dept,
       priority = excluded.priority,
       due_date = excluded.due_date,
       assigner_email = excluded.assigner_email,
       assigner_name = excluded.assigner_name,
       recipient_email = excluded.recipient_email,
       recipient_name = excluded.recipient_name,
       recipient_todo_list_id = excluded.recipient_todo_list_id,
       recipient_todo_task_id = excluded.recipient_todo_task_id,
       proof_instructions = excluded.proof_instructions,
       updated_at = excluded.updated_at,
       version = assignments.version + 1`
  ).bind(
    id,
    String(body.appTaskId || ''),
    title,
    String(body.summary || '').slice(0, 8000),
    String(body.dept || ''),
    String(body.priority || 'Normal'),
    String(body.dueDate || ''),
    assignerEmail,
    assignerName,
    recipientEmail,
    String(body.recipientName || ''),
    String(body.recipientTodoListId || ''),
    String(body.recipientTodoTaskId || ''),
    String(body.proofInstructions || ''),
    now,
    now,
  ).run();

  const updatedRow=await env.DPEG_ASSIGNMENTS.prepare('SELECT version FROM assignments WHERE id = ?').bind(id).first();
  return json({ success: true, id, version:Number(updatedRow?.version||1) });
}

// ── /assignments endpoint: fetch both directions for the current user (D1) ───
async function handleAssignments(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_ASSIGNMENTS) return json({ error: 'DPEG_ASSIGNMENTS D1 binding is not configured' }, 501);
  await ensureAssignmentProofColumns(env);
  if (request.method !== 'GET') return json({ error: 'Method not allowed' }, 405);

  const url = new URL(request.url);
  const requestedEmail = extractEmailAddress(url.searchParams.get('email') || '');
  const tokenEmail = userEmailFromClaims(claims);
  if (!requestedEmail || requestedEmail !== tokenEmail) {
    return json({ error: 'You can only fetch your own assignments' }, 403);
  }

  const [toMe, byMe] = await Promise.all([
    env.DPEG_ASSIGNMENTS.prepare(
      'SELECT * FROM assignments WHERE recipient_email = ? ORDER BY created_at DESC'
    ).bind(tokenEmail).all(),
    env.DPEG_ASSIGNMENTS.prepare(
      'SELECT * FROM assignments WHERE assigner_email = ? ORDER BY created_at DESC'
    ).bind(tokenEmail).all(),
  ]);

  const shape = (row) => ({
    id: row.id,
    appTaskId: row.app_task_id,
    title: row.title,
    summary: row.summary,
    dept: row.dept,
    priority: row.priority,
    dueDate: row.due_date,
    assignerEmail: row.assigner_email,
    assignerName: row.assigner_name,
    recipientEmail: row.recipient_email,
    recipientName: row.recipient_name,
    status: row.status,
    progressNote: row.progress_note,
    proofStatus: row.proof_status || 'none',
    proofSubmittedAt: row.proof_submitted_at,
    proofReviewedAt: row.proof_reviewed_at,
    proofNotificationId: row.proof_notification_id,
    proofInstructions: row.proof_instructions,
    updateAlertAt: row.update_alert_at || null,
    reminderCount: Number(row.reminder_count || 0),
    assignerMessageSeenCount: Number(row.assigner_message_seen_count || 0),
    recipientMessageSeenCount: Number(row.recipient_message_seen_count || 0),
    recipientReminderSeenCount: row.recipient_reminder_seen_count == null ? null : Number(row.recipient_reminder_seen_count),
    cancelReason: row.cancel_reason || '',
    cancelledAt: row.cancelled_at || null,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
    version: Number(row.version || 1),
  });

  return json({
    assignedToMe: (toMe.results || []).map(shape),
    assignedByMe: (byMe.results || []).map(shape),
  });
}

// ── /assignment-status endpoint: recipient updates progress (D1) ─────────────
// 'Done' is deliberately excluded — it is only ever set server-side by the
// proof-approval path (see the proof_status sync above), never chosen manually.
// 'Accepted' was removed from the flow (Assigned → In Progress → Submitted → Done).
const ASSIGNMENT_STATUSES = new Set(['Assigned', 'In Progress']);

async function handleAssignmentStatus(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_ASSIGNMENTS) return json({ error: 'DPEG_ASSIGNMENTS D1 binding is not configured' }, 501);
  await ensureAssignmentProofColumns(env);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const id = String(body.id || '').trim();
  const newStatus = String(body.status || '').trim();
  const expectedVersion = Number(body.expectedVersion);
  if (!id || !ASSIGNMENT_STATUSES.has(newStatus) || !Number.isInteger(expectedVersion) || expectedVersion < 1) {
    return json({ error: 'id, a valid status and expectedVersion are required' }, 400);
  }

  const row = await env.DPEG_ASSIGNMENTS
    .prepare('SELECT recipient_email FROM assignments WHERE id = ?')
    .bind(id).first();
  if (!row) return json({ error: 'Assignment not found' }, 404);

  const tokenEmail = userEmailFromClaims(claims);
  if (extractEmailAddress(row.recipient_email) !== tokenEmail) {
    return json({ error: 'Only the recipient can update this assignment' }, 403);
  }
  if(env.DPEG_DATA){
    const editLock=await env.DPEG_DATA.get(`task-edit-lock:${id}`,'json');
    if(editLock&&Number(editLock.expiresAt||0)>Date.now()&&extractEmailAddress(editLock.email)!==tokenEmail){
      return json({error:'currently_editing',editorName:editLock.name||editLock.email},423);
    }
  }

  const updatedAt = new Date().toISOString();
  const result = await env.DPEG_ASSIGNMENTS.prepare(
    `UPDATE assignments
        SET status = ?, progress_note = ?, update_alert_at = NULL,
            updated_at = ?, version = version + 1
      WHERE id = ? AND version = ?`
  ).bind(
    newStatus,
    body.progressNote != null ? String(body.progressNote).slice(0, 2000) : null,
    updatedAt,
    id,
    expectedVersion,
  ).run();

  if (!result.meta?.changes) {
    return json({
      error: 'This assignment was changed by another user. Refresh and try again.',
      code: 'VERSION_CONFLICT',
    }, 409);
  }

  return json({ success: true, updatedAt, version: expectedVersion + 1 });
}

// ── /assignment-alert endpoint: assigner sends a one-click "update required"
// nudge (D1). Independent of the KV follow-up thread — just a timestamp flag
// on the assignment row, cleared automatically once the recipient updates
// their status or submits proof.
async function handleAssignmentAlert(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_ASSIGNMENTS) return json({ error: 'DPEG_ASSIGNMENTS D1 binding is not configured' }, 501);
  await ensureAssignmentProofColumns(env);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const id = String(body.id || '').trim();
  if (!id) return json({ error: 'id is required' }, 400);

  const row = await env.DPEG_ASSIGNMENTS
    .prepare('SELECT assigner_email FROM assignments WHERE id = ?')
    .bind(id).first();
  if (!row) return json({ error: 'Assignment not found' }, 404);

  const tokenEmail = userEmailFromClaims(claims);
  if (extractEmailAddress(row.assigner_email) !== tokenEmail) {
    return json({ error: 'Only the assigner can send an update-required alert' }, 403);
  }

  const updateAlertAt = new Date().toISOString();
  await env.DPEG_ASSIGNMENTS.prepare(
    `UPDATE assignments
        SET update_alert_at = ?,
            reminder_count = COALESCE(reminder_count, 0) + 1,
            updated_at = ?,
            version = version + 1
      WHERE id = ?`
  ).bind(updateAlertAt, updateAlertAt, id).run();

  const updated = await env.DPEG_ASSIGNMENTS.prepare(
    'SELECT reminder_count, version, updated_at FROM assignments WHERE id = ?'
  ).bind(id).first();

  return json({
    success: true,
    updateAlertAt,
    reminderCount: Number(updated?.reminder_count || 0),
    version: Number(updated?.version || 1),
    updatedAt: updated?.updated_at || updateAlertAt,
  });
}

// ── /assignment-alert-clear endpoint: manually dismiss an alert (D1). Either
// party can clear it — the recipient acknowledging it, or the assigner
// retracting it — since the automatic clear-on-submit/approve/decline path
// above doesn't cover every case (e.g. an alert sent while proof is already
// sitting in review, with no status change left for the recipient to make).
async function handleAssignmentAlertClear(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_ASSIGNMENTS) return json({ error: 'DPEG_ASSIGNMENTS D1 binding is not configured' }, 501);
  await ensureAssignmentProofColumns(env);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const id = String(body.id || '').trim();
  if (!id) return json({ error: 'id is required' }, 400);

  const row = await env.DPEG_ASSIGNMENTS
    .prepare('SELECT assigner_email, recipient_email FROM assignments WHERE id = ?')
    .bind(id).first();
  if (!row) return json({ error: 'Assignment not found' }, 404);

  const tokenEmail = userEmailFromClaims(claims);
  const isAssigner = extractEmailAddress(row.assigner_email) === tokenEmail;
  const isRecipient = extractEmailAddress(row.recipient_email) === tokenEmail;
  if (!isAssigner && !isRecipient) {
    return json({ error: 'Only the assigner or recipient can dismiss this alert' }, 403);
  }

  await env.DPEG_ASSIGNMENTS.prepare(
    'UPDATE assignments SET update_alert_at = NULL WHERE id = ?'
  ).bind(id).run();

  return json({ success: true });
}

// Persists per-user acknowledgement so messages/reminders do not become new
// again after a refresh, a new day, or a different browser.
async function handleAssignmentSeen(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_ASSIGNMENTS) return json({ error: 'DPEG_ASSIGNMENTS D1 binding is not configured' }, 501);
  await ensureAssignmentProofColumns(env);
  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }
  const id=String(body.id||'').trim();
  if(!id)return json({error:'id is required'},400);
  const row=await env.DPEG_ASSIGNMENTS.prepare(
    'SELECT assigner_email, recipient_email FROM assignments WHERE id = ?'
  ).bind(id).first();
  if(!row)return json({error:'Assignment not found'},404);
  const email=userEmailFromClaims(claims);
  const isAssigner=extractEmailAddress(row.assigner_email)===email;
  const isRecipient=extractEmailAddress(row.recipient_email)===email;
  if(!isAssigner&&!isRecipient)return json({error:'Not authorized for this assignment'},403);
  const threadLen=Math.max(0,Number(body.threadLen||0));
  const reminderCount=Math.max(0,Number(body.reminderCount||0));
  if(isAssigner){
    await env.DPEG_ASSIGNMENTS.prepare(
      'UPDATE assignments SET assigner_message_seen_count = MAX(COALESCE(assigner_message_seen_count,0), ?) WHERE id = ?'
    ).bind(threadLen,id).run();
  }else{
    await env.DPEG_ASSIGNMENTS.prepare(
      `UPDATE assignments
          SET recipient_message_seen_count = MAX(COALESCE(recipient_message_seen_count,0), ?),
              recipient_reminder_seen_count = MAX(COALESCE(recipient_reminder_seen_count,0), ?)
        WHERE id = ?`
    ).bind(threadLen,reminderCount,id).run();
  }
  return json({success:true,threadLen,reminderCount});
}

// ── /assignment-cancel endpoint: assigner calls off a delegated task (D1) ────
// Assigner-only, mirroring the ownership check on /assignment-alert. Sets a
// terminal 'Cancelled' status distinct from the proof-approval 'Done' path —
// the recipient's copy simply stops accepting status/proof changes once they
// see it, the same way an approved task already does.
async function handleAssignmentCancel(request, env) {
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_ASSIGNMENTS) return json({ error: 'DPEG_ASSIGNMENTS D1 binding is not configured' }, 501);
  await ensureAssignmentProofColumns(env);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const id = String(body.id || '').trim();
  if (!id) return json({ error: 'id is required' }, 400);
  const reason = String(body.reason || '').trim().slice(0, 2000);

  const row = await env.DPEG_ASSIGNMENTS
    .prepare('SELECT assigner_email, status, proof_status FROM assignments WHERE id = ?')
    .bind(id).first();
  if (!row) return json({ error: 'Assignment not found' }, 404);

  const tokenEmail = userEmailFromClaims(claims);
  if (extractEmailAddress(row.assigner_email) !== tokenEmail) {
    return json({ error: 'Only the assigner can cancel this task' }, 403);
  }
  if (row.status === 'Cancelled') return json({ error: 'This task is already cancelled' }, 409);
  if (row.proof_status === 'approved') return json({ error: 'A completed task cannot be cancelled' }, 409);

  const now = new Date().toISOString();
  await env.DPEG_ASSIGNMENTS.prepare(
    `UPDATE assignments
        SET status = 'Cancelled', cancel_reason = ?, cancelled_at = ?,
            update_alert_at = NULL, updated_at = ?, version = version + 1
      WHERE id = ?`
  ).bind(reason, now, now, id).run();

  return json({ success: true, cancelledAt: now, reason });
}

// ── Staging-only normalized task API ─────────────────────────────────────────
// This endpoint exercises the new D1 workflow schema with synthetic records.
// It is intentionally unavailable unless APP_ENV is exactly "staging".
const STAGING_TASK_STATUSES = new Set(['Pending', 'In Progress', 'Done', 'Cancelled']);
const STAGING_TASK_PRIORITIES = new Set(['Low', 'Normal', 'High']);

function stagingTaskShape(row) {
  return {
    id: row.id,
    title: row.title,
    summary: row.summary,
    departmentName: row.department_name,
    priority: row.priority,
    dueDate: row.due_date,
    status: row.status,
    sourceType: row.source_type,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
    completedAt: row.completed_at,
    cancelledAt: row.cancelled_at,
    version: Number(row.version || 1),
  };
}

async function recordStagingTaskEvent(env, taskId, actorEmail, eventType, eventData, createdAt) {
  await env.DPEG_ASSIGNMENTS.prepare(
    `INSERT INTO task_events
       (id, assignment_id, app_task_id, actor_email, event_type, event_data, created_at)
     VALUES (?, NULL, ?, ?, ?, ?, ?)`
  ).bind(
    `evt-${crypto.randomUUID()}`,
    taskId,
    actorEmail,
    eventType,
    JSON.stringify(eventData || {}),
    createdAt,
  ).run();
}

async function handleStagingTasksCore(request, env) {
  if (!isStaging(env)) return json({ error: 'Not found' }, 404);
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_ASSIGNMENTS) return json({ error: 'Staging D1 binding is not configured' }, 501);

  const ownerEmail = userEmailFromClaims(claims);
  if (request.method === 'GET') {
    const [taskResult, assignmentResult, proofResult, reminderResult, allMessageThreads] = await Promise.all([
      env.DPEG_ASSIGNMENTS.prepare(
      `SELECT id, title, summary, department_name, priority, due_date, status,
              source_type, created_at, updated_at, completed_at, cancelled_at, version
         FROM tasks
        WHERE owner_email = ? AND source_type = 'staging_test'
        ORDER BY updated_at DESC, id DESC`
      ).bind(ownerEmail).all(),
      env.DPEG_ASSIGNMENTS.prepare(
        `SELECT a.*
           FROM assignments a
           JOIN tasks t ON t.id = a.app_task_id
          WHERE t.source_type = 'staging_test'
            AND (a.assigner_email = ? OR a.recipient_email = ?)
          ORDER BY a.updated_at DESC, a.id DESC`
      ).bind(ownerEmail, ownerEmail).all(),
      env.DPEG_ASSIGNMENTS.prepare(
        `SELECT p.*
           FROM proof_submissions p
           JOIN assignments a ON a.id = p.assignment_id
           JOIN tasks t ON t.id = a.app_task_id
          WHERE t.source_type = 'staging_test'
            AND (a.assigner_email = ? OR a.recipient_email = ?)
          ORDER BY p.submitted_at DESC, p.id DESC`
      ).bind(ownerEmail, ownerEmail).all(),
      env.DPEG_ASSIGNMENTS.prepare(
        `SELECT r.*
           FROM reminders r
           JOIN assignments a ON a.id = r.assignment_id
           JOIN tasks t ON t.id = a.app_task_id
          WHERE t.source_type = 'staging_test'
            AND (a.assigner_email = ? OR a.recipient_email = ?)
          ORDER BY r.created_at DESC, r.id DESC`
      ).bind(ownerEmail, ownerEmail).all(),
      loadTaskMessageThreads(env, claims),
    ]);
    const messageThreads = allMessageThreads
      .filter(thread => String(thread.appTaskId || '').startsWith('stg-task-'));
    return json({
      tasks: (taskResult.results || []).map(stagingTaskShape),
      assignments: assignmentResult.results || [],
      proofs: proofResult.results || [],
      reminders: reminderResult.results || [],
      messageThreads,
    });
  }

  if (request.method !== 'POST') return json({ error: 'Method not allowed' }, 405);
  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const action = String(body.action || '').trim().toLowerCase();
  if (action === 'assignment_status') {
    const assignmentId = String(body.assignmentId || '').trim();
    const expectedVersion = Number(body.expectedVersion);
    const nextStatus = String(body.status || '').trim();
    if (!assignmentId || !Number.isInteger(expectedVersion) || expectedVersion < 1
        || !new Set(['Assigned', 'In Progress']).has(nextStatus)) {
      return json({ error: 'assignmentId, expectedVersion and a valid status are required' }, 400);
    }
    const assignment = await env.DPEG_ASSIGNMENTS.prepare(
      `SELECT a.* FROM assignments a
        JOIN tasks t ON t.id = a.app_task_id
       WHERE a.id = ? AND t.source_type = 'staging_test'`
    ).bind(assignmentId).first();
    if (!assignment) return json({ error: 'Staging assignment not found' }, 404);
    if (ownerEmail !== extractEmailAddress(assignment.recipient_email)) {
      return json({ error: 'Only the assignee can update staging status' }, 403);
    }
    const now = new Date().toISOString();
    const updated = await env.DPEG_ASSIGNMENTS.prepare(
      `UPDATE assignments
          SET status = ?, progress_note = ?, update_alert_at = NULL,
              updated_at = ?, version = version + 1
        WHERE id = ? AND version = ?`
    ).bind(
      nextStatus,
      body.progressNote == null ? assignment.progress_note : String(body.progressNote).slice(0, 2000),
      now,
      assignmentId,
      expectedVersion,
    ).run();
    if (!updated.meta?.changes) {
      return json({ error: 'version_conflict', message: 'Reload before changing this status.' }, 409);
    }
    await env.DPEG_ASSIGNMENTS.prepare(
      `INSERT INTO task_events
         (id, assignment_id, app_task_id, actor_email, event_type, event_data, created_at)
       VALUES (?, ?, ?, ?, 'staging_status_changed', ?, ?)`
    ).bind(
      `evt-${crypto.randomUUID()}`,
      assignmentId,
      assignment.app_task_id,
      ownerEmail,
      JSON.stringify({ previousStatus: assignment.status, status: nextStatus }),
      now,
    ).run();
    return json({ success: true, status: nextStatus, version: expectedVersion + 1 });
  }

  if (action === 'message') {
    const assignmentId = String(body.assignmentId || '').trim();
    const assignment = await env.DPEG_ASSIGNMENTS.prepare(
      `SELECT a.* FROM assignments a
        JOIN tasks t ON t.id = a.app_task_id
       WHERE a.id = ? AND t.source_type = 'staging_test'`
    ).bind(assignmentId).first();
    if (!assignment) return json({ error: 'Staging assignment not found' }, 404);
    if (ownerEmail !== extractEmailAddress(assignment.assigner_email)
        && ownerEmail !== extractEmailAddress(assignment.recipient_email)) {
      return json({ error: 'You are not a participant in this staging assignment' }, 403);
    }
    const inserted = await insertTaskMessage(env, {
      appTaskId: assignment.app_task_id,
      taskTitle: assignment.title,
      assignerEmail: assignment.assigner_email,
      recipientEmail: assignment.recipient_email,
      recipientName: assignment.recipient_name,
      senderName: String(claims.name || ownerEmail),
      message: body.message,
    }, claims);
    if (inserted.error) return inserted.error;
    return json({ success: true, message: inserted }, 201);
  }

  if (action === 'remind') {
    const assignmentId = String(body.assignmentId || '').trim();
    const expectedVersion = Number(body.expectedVersion);
    if (!assignmentId || !Number.isInteger(expectedVersion) || expectedVersion < 1) {
      return json({ error: 'assignmentId and expectedVersion are required' }, 400);
    }
    const assignment = await env.DPEG_ASSIGNMENTS.prepare(
      `SELECT a.* FROM assignments a
        JOIN tasks t ON t.id = a.app_task_id
       WHERE a.id = ? AND t.source_type = 'staging_test'`
    ).bind(assignmentId).first();
    if (!assignment) return json({ error: 'Staging assignment not found' }, 404);
    if (ownerEmail !== extractEmailAddress(assignment.assigner_email)) {
      return json({ error: 'Only the assigner can send a staging reminder' }, 403);
    }
    const idempotencyKey = String(body.idempotencyKey || '').trim().slice(0, 200) || null;
    if (idempotencyKey) {
      const existing = await env.DPEG_ASSIGNMENTS.prepare(
        'SELECT id, created_at FROM reminders WHERE idempotency_key = ?'
      ).bind(idempotencyKey).first();
      if (existing) return json({ success: true, duplicate: true, reminder: existing });
    }
    const now = new Date().toISOString();
    const updated = await env.DPEG_ASSIGNMENTS.prepare(
      `UPDATE assignments
          SET reminder_count = COALESCE(reminder_count, 0) + 1,
              update_alert_at = ?, updated_at = ?, version = version + 1
        WHERE id = ? AND version = ?`
    ).bind(now, now, assignmentId, expectedVersion).run();
    if (!updated.meta?.changes) {
      return json({ error: 'version_conflict', message: 'Reload this assignment before reminding again.' }, 409);
    }
    const reminderId = `rem-${crypto.randomUUID()}`;
    await env.DPEG_ASSIGNMENTS.batch([
      env.DPEG_ASSIGNMENTS.prepare(
        `INSERT INTO reminders
           (id, assignment_id, sender_email, recipient_email, created_at, idempotency_key)
         VALUES (?, ?, ?, ?, ?, ?)`
      ).bind(reminderId, assignmentId, ownerEmail, assignment.recipient_email, now, idempotencyKey),
      env.DPEG_ASSIGNMENTS.prepare(
        `INSERT INTO task_events
           (id, assignment_id, app_task_id, actor_email, event_type, event_data, created_at)
         VALUES (?, ?, ?, ?, 'staging_reminder_sent', ?, ?)`
      ).bind(`evt-${crypto.randomUUID()}`, assignmentId, assignment.app_task_id, ownerEmail, '{}', now),
    ]);
    return json({ success: true, reminder: { id: reminderId, createdAt: now }, version: expectedVersion + 1 }, 201);
  }

  if (action === 'submit_proof') {
    const assignmentId = String(body.assignmentId || '').trim();
    const expectedVersion = Number(body.expectedVersion);
    if (!assignmentId || !Number.isInteger(expectedVersion) || expectedVersion < 1) {
      return json({ error: 'assignmentId and expectedVersion are required' }, 400);
    }
    const assignment = await env.DPEG_ASSIGNMENTS.prepare(
      `SELECT a.* FROM assignments a
        JOIN tasks t ON t.id = a.app_task_id
       WHERE a.id = ? AND t.source_type = 'staging_test'`
    ).bind(assignmentId).first();
    if (!assignment) return json({ error: 'Staging assignment not found' }, 404);
    if (ownerEmail !== extractEmailAddress(assignment.recipient_email)) {
      return json({ error: 'Only the assignee can submit staging proof' }, 403);
    }
    const idempotencyKey = String(body.idempotencyKey || '').trim().slice(0, 200) || null;
    if (idempotencyKey) {
      const existing = await env.DPEG_ASSIGNMENTS.prepare(
        'SELECT id, status, submitted_at FROM proof_submissions WHERE idempotency_key = ?'
      ).bind(idempotencyKey).first();
      if (existing) return json({ success: true, duplicate: true, proof: existing });
    }
    const now = new Date().toISOString();
    const updated = await env.DPEG_ASSIGNMENTS.prepare(
      `UPDATE assignments
          SET status = 'Submitted', proof_status = 'submitted', proof_submitted_at = ?,
              proof_reviewed_at = NULL, update_alert_at = NULL,
              updated_at = ?, version = version + 1
        WHERE id = ? AND version = ?`
    ).bind(now, now, assignmentId, expectedVersion).run();
    if (!updated.meta?.changes) {
      return json({ error: 'version_conflict', message: 'Reload this assignment before submitting proof.' }, 409);
    }
    const proofId = `proof-${crypto.randomUUID()}`;
    const statements = [
      env.DPEG_ASSIGNMENTS.prepare(
        `INSERT INTO proof_submissions
           (id, assignment_id, app_task_id, submitter_email, submitter_name, note,
            status, submitted_at, idempotency_key)
         VALUES (?, ?, ?, ?, ?, ?, 'pending', ?, ?)`
      ).bind(
        proofId,
        assignmentId,
        assignment.app_task_id,
        ownerEmail,
        String(claims.name || ownerEmail),
        String(body.note || '').slice(0, 8000),
        now,
        idempotencyKey,
      ),
    ];
    for (const file of (Array.isArray(body.files) ? body.files : []).slice(0, 10)) {
      const fileName = String(file?.fileName || file?.name || '').trim().slice(0, 500);
      if (!fileName) continue;
      statements.push(env.DPEG_ASSIGNMENTS.prepare(
        `INSERT INTO proof_files
           (id, submission_id, drive_provider, drive_item_id, file_name, mime_type,
            size_bytes, web_url, created_at)
         VALUES (?, ?, 'staging', NULL, ?, ?, ?, ?, ?)`
      ).bind(
        `pf-${crypto.randomUUID()}`,
        proofId,
        fileName,
        String(file?.mimeType || file?.type || '').slice(0, 200) || null,
        Number.isFinite(Number(file?.sizeBytes || file?.size)) ? Number(file?.sizeBytes || file?.size) : null,
        String(file?.webUrl || 'about:blank').slice(0, 2000),
        now,
      ));
    }
    statements.push(env.DPEG_ASSIGNMENTS.prepare(
      `INSERT INTO task_events
         (id, assignment_id, app_task_id, actor_email, event_type, event_data, created_at)
       VALUES (?, ?, ?, ?, 'staging_proof_submitted', ?, ?)`
    ).bind(`evt-${crypto.randomUUID()}`, assignmentId, assignment.app_task_id, ownerEmail, JSON.stringify({ proofId }), now));
    await env.DPEG_ASSIGNMENTS.batch(statements);
    return json({ success: true, proof: { id: proofId, status: 'pending', submittedAt: now }, version: expectedVersion + 1 }, 201);
  }

  if (action === 'review_proof') {
    const proofId = String(body.proofId || '').trim();
    const expectedVersion = Number(body.expectedVersion);
    const decision = String(body.decision || '').trim().toLowerCase();
    if (!proofId || !Number.isInteger(expectedVersion) || expectedVersion < 1
        || !new Set(['approved', 'changes_requested']).has(decision)) {
      return json({ error: 'proofId, expectedVersion and a valid decision are required' }, 400);
    }
    const proof = await env.DPEG_ASSIGNMENTS.prepare(
      `SELECT p.*, a.assigner_email, a.recipient_email, a.version AS assignment_version
         FROM proof_submissions p
         JOIN assignments a ON a.id = p.assignment_id
         JOIN tasks t ON t.id = a.app_task_id
        WHERE p.id = ? AND t.source_type = 'staging_test'`
    ).bind(proofId).first();
    if (!proof) return json({ error: 'Staging proof not found' }, 404);
    if (ownerEmail !== extractEmailAddress(proof.assigner_email)) {
      return json({ error: 'Only the assigner can review staging proof' }, 403);
    }
    if (proof.status !== 'pending') return json({ error: 'This proof was already reviewed' }, 409);
    if (Number(proof.assignment_version || 1) !== expectedVersion) {
      return json({ error: 'version_conflict', currentVersion: Number(proof.assignment_version || 1) }, 409);
    }
    const now = new Date().toISOString();
    const assignmentStatus = decision === 'approved' ? 'Done' : 'Assigned';
    const assignmentProofStatus = decision === 'approved' ? 'approved' : 'declined';
    const updated = await env.DPEG_ASSIGNMENTS.prepare(
      `UPDATE assignments
          SET status = ?, proof_status = ?, proof_reviewed_at = ?, updated_at = ?, version = version + 1
        WHERE id = ? AND version = ?`
    ).bind(assignmentStatus, assignmentProofStatus, now, now, proof.assignment_id, expectedVersion).run();
    if (!updated.meta?.changes) {
      return json({ error: 'version_conflict', message: 'Reload this proof before reviewing it.' }, 409);
    }
    await env.DPEG_ASSIGNMENTS.batch([
      env.DPEG_ASSIGNMENTS.prepare(
        `UPDATE proof_submissions
            SET status = ?, reviewed_at = ?, reviewer_email = ?, review_reason = ?
          WHERE id = ? AND status = 'pending'`
      ).bind(decision, now, ownerEmail, String(body.reason || '').slice(0, 8000), proofId),
      env.DPEG_ASSIGNMENTS.prepare(
        `INSERT INTO task_events
           (id, assignment_id, app_task_id, actor_email, event_type, event_data, created_at)
         VALUES (?, ?, ?, ?, 'staging_proof_reviewed', ?, ?)`
      ).bind(
        `evt-${crypto.randomUUID()}`,
        proof.assignment_id,
        proof.app_task_id,
        ownerEmail,
        JSON.stringify({ proofId, decision }),
        now,
      ),
    ]);
    return json({ success: true, decision, version: expectedVersion + 1 });
  }

  if (action === 'delegate') {
    const title = String(body.title || '').trim().slice(0, 500);
    const recipientEmail = extractEmailAddress(body.recipientEmail || '');
    if (!title || !recipientEmail) return json({ error: 'title and recipientEmail are required' }, 400);
    if (!recipientEmail.endsWith('@dhananipeg.com')) {
      return json({ error: 'Only @dhananipeg.com staging recipients are supported' }, 403);
    }
    const priority = STAGING_TASK_PRIORITIES.has(String(body.priority || ''))
      ? String(body.priority)
      : 'Normal';
    const now = new Date().toISOString();
    const taskId = `stg-task-${crypto.randomUUID()}`;
    const assignmentId = `stg-asg-${crypto.randomUUID()}`;
    const summary = String(body.summary || '').slice(0, 8000);
    const department = String(body.departmentName || 'Needs Department').trim().slice(0, 200) || 'Needs Department';
    const dueDate = String(body.dueDate || '').trim().slice(0, 40) || null;
    await env.DPEG_ASSIGNMENTS.batch([
      env.DPEG_ASSIGNMENTS.prepare(
        `INSERT INTO tasks
           (id, owner_email, title, summary, department_name, priority, due_date,
            status, source_type, created_at, updated_at, version)
         VALUES (?, ?, ?, ?, ?, ?, ?, 'Pending', 'staging_test', ?, ?, 1)`
      ).bind(taskId, ownerEmail, title, summary, department, priority, dueDate, now, now),
      env.DPEG_ASSIGNMENTS.prepare(
        `INSERT INTO assignments
           (id, app_task_id, title, summary, dept, priority, due_date,
            assigner_email, assigner_name, recipient_email, recipient_name,
            status, proof_status, proof_instructions, created_at, updated_at, version)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Assigned', 'none', ?, ?, ?, 1)`
      ).bind(
        assignmentId,
        taskId,
        title,
        summary,
        department,
        priority,
        dueDate,
        ownerEmail,
        String(claims.name || ownerEmail),
        recipientEmail,
        String(body.recipientName || '').trim().slice(0, 300),
        String(body.proofInstructions || '').slice(0, 4000),
        now,
        now,
      ),
      env.DPEG_ASSIGNMENTS.prepare(
        `INSERT INTO task_events
           (id, assignment_id, app_task_id, actor_email, event_type, event_data, created_at)
         VALUES (?, ?, ?, ?, 'staging_task_delegated', ?, ?)`
      ).bind(
        `evt-${crypto.randomUUID()}`,
        assignmentId,
        taskId,
        ownerEmail,
        JSON.stringify({ recipientEmail, title }),
        now,
      ),
    ]);
    return json({
      success: true,
      task: { id: taskId, title, status: 'Pending', version: 1, createdAt: now },
      assignment: { id: assignmentId, recipientEmail, status: 'Assigned', version: 1 },
    }, 201);
  }

  if (action === 'create') {
    const title = String(body.title || '').trim().slice(0, 500);
    if (!title) return json({ error: 'title is required' }, 400);
    const priority = STAGING_TASK_PRIORITIES.has(String(body.priority || ''))
      ? String(body.priority)
      : 'Normal';
    const now = new Date().toISOString();
    const id = `stg-task-${crypto.randomUUID()}`;
    await env.DPEG_ASSIGNMENTS.prepare(
      `INSERT INTO tasks
         (id, owner_email, title, summary, department_name, priority, due_date,
          status, source_type, created_at, updated_at, version)
       VALUES (?, ?, ?, ?, ?, ?, ?, 'Pending', 'staging_test', ?, ?, 1)`
    ).bind(
      id,
      ownerEmail,
      title,
      String(body.summary || '').slice(0, 8000),
      String(body.departmentName || 'Needs Department').trim().slice(0, 200) || 'Needs Department',
      priority,
      String(body.dueDate || '').trim().slice(0, 40) || null,
      now,
      now,
    ).run();
    await recordStagingTaskEvent(env, id, ownerEmail, 'staging_task_created', { title }, now);
    return json({ success: true, task: { id, title, status: 'Pending', version: 1, createdAt: now } }, 201);
  }

  if (action === 'update') {
    const id = String(body.id || '').trim();
    const expectedVersion = Number(body.expectedVersion);
    if (!id || !Number.isInteger(expectedVersion) || expectedVersion < 1) {
      return json({ error: 'id and expectedVersion are required' }, 400);
    }
    const current = await env.DPEG_ASSIGNMENTS.prepare(
      `SELECT * FROM tasks
        WHERE id = ? AND owner_email = ? AND source_type = 'staging_test'`
    ).bind(id, ownerEmail).first();
    if (!current) return json({ error: 'Staging task not found' }, 404);
    if (Number(current.version || 1) !== expectedVersion) {
      return json({
        error: 'version_conflict',
        message: 'This task changed since it was opened. Reload and try again.',
        currentVersion: Number(current.version || 1),
      }, 409);
    }

    const title = body.title == null ? String(current.title) : String(body.title).trim().slice(0, 500);
    if (!title) return json({ error: 'title cannot be empty' }, 400);
    const nextStatus = body.status == null ? String(current.status) : String(body.status);
    if (!STAGING_TASK_STATUSES.has(nextStatus)) return json({ error: 'Invalid status' }, 400);
    const nextPriority = body.priority == null ? String(current.priority) : String(body.priority);
    if (!STAGING_TASK_PRIORITIES.has(nextPriority)) return json({ error: 'Invalid priority' }, 400);
    const now = new Date().toISOString();
    const completedAt = nextStatus === 'Done' ? (current.completed_at || now) : null;
    const cancelledAt = nextStatus === 'Cancelled' ? (current.cancelled_at || now) : null;
    const result = await env.DPEG_ASSIGNMENTS.prepare(
      `UPDATE tasks
          SET title = ?, summary = ?, department_name = ?, priority = ?, due_date = ?,
              status = ?, completed_at = ?, cancelled_at = ?, updated_at = ?, version = version + 1
        WHERE id = ? AND owner_email = ? AND source_type = 'staging_test' AND version = ?`
    ).bind(
      title,
      body.summary == null ? String(current.summary || '') : String(body.summary).slice(0, 8000),
      body.departmentName == null
        ? String(current.department_name || 'Needs Department')
        : (String(body.departmentName).trim().slice(0, 200) || 'Needs Department'),
      nextPriority,
      body.dueDate == null ? current.due_date : (String(body.dueDate).trim().slice(0, 40) || null),
      nextStatus,
      completedAt,
      cancelledAt,
      now,
      id,
      ownerEmail,
      expectedVersion,
    ).run();
    if (!result.meta?.changes) {
      const latest = await env.DPEG_ASSIGNMENTS.prepare(
        'SELECT version FROM tasks WHERE id = ? AND owner_email = ?'
      ).bind(id, ownerEmail).first();
      return json({
        error: 'version_conflict',
        message: 'This task was updated by another request. Reload and try again.',
        currentVersion: Number(latest?.version || expectedVersion),
      }, 409);
    }
    await recordStagingTaskEvent(env, id, ownerEmail, 'staging_task_updated', {
      previousStatus: current.status,
      status: nextStatus,
      previousVersion: expectedVersion,
      version: expectedVersion + 1,
    }, now);
    const updated = await env.DPEG_ASSIGNMENTS.prepare(
      'SELECT * FROM tasks WHERE id = ? AND owner_email = ?'
    ).bind(id, ownerEmail).first();
    return json({ success: true, task: stagingTaskShape(updated) });
  }

  return json({
    error: 'Supported actions are delegate, create, update, assignment_status, message, remind, submit_proof and review_proof',
  }, 400);
}

async function broadcastStagingChange(env) {
  if (!isStaging(env) || !env.STAGING_REALTIME_HUB) return;
  try {
    const hub = env.STAGING_REALTIME_HUB.getByName('dpeg-staging-workflow');
    await hub.fetch('https://staging-realtime.internal/broadcast', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'workflow_changed', at: new Date().toISOString() }),
    });
  } catch (error) {
    console.warn('Staging realtime broadcast failed:', error?.message || error);
  }
}

async function handleStagingTasks(request, env) {
  const response = await handleStagingTasksCore(request, env);
  if (request.method === 'POST' && response.status >= 200 && response.status < 300) {
    await broadcastStagingChange(env);
  }
  return response;
}

async function withStagingRealtimeBroadcast(request, env, responsePromise) {
  const response = await responsePromise;
  if (isStaging(env) && request.method !== 'GET'
      && response.status >= 200 && response.status < 300) {
    await broadcastStagingChange(env);
  }
  return response;
}

async function handleStagingRealtimeTicket(request, env) {
  if (!isStaging(env)) return json({ error: 'Not found' }, 404);
  if (request.method !== 'POST') return json({ error: 'Method not allowed' }, 405);
  const { error, status, claims } = await validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.STAGING_REALTIME_HUB) return json({ error: 'Staging realtime binding is not configured' }, 501);
  const hub = env.STAGING_REALTIME_HUB.getByName('dpeg-staging-workflow');
  const response = await hub.fetch('https://staging-realtime.internal/ticket', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email: userEmailFromClaims(claims) }),
  });
  const result = await response.json().catch(() => ({}));
  if (!response.ok) return json({ error: result.error || 'Could not create realtime ticket' }, response.status);
  return json(result);
}

async function handleStagingRealtimeSocket(request, env) {
  if (!isStaging(env)) return json({ error: 'Not found' }, 404);
  if (request.headers.get('Upgrade')?.toLowerCase() !== 'websocket') {
    return json({ error: 'WebSocket upgrade required' }, 426);
  }
  const origin = request.headers.get('Origin') || '';
  if (!ALLOWED_ORIGINS.has(origin)) return json({ error: 'Origin is not allowed' }, 403);
  if (!env.STAGING_REALTIME_HUB) {
    return json({ error: 'Staging realtime bindings are not configured' }, 501);
  }
  const ticket = String(new URL(request.url).searchParams.get('ticket') || '').trim();
  if (!ticket) return json({ error: 'Realtime ticket is required' }, 401);
  const hub = env.STAGING_REALTIME_HUB.getByName('dpeg-staging-workflow');
  return hub.fetch(request);
}

export class StagingRealtimeHub {
  constructor(state) {
    this.state = state;
    this.state.setWebSocketAutoResponse(new WebSocketRequestResponsePair('ping', 'pong'));
  }

  async fetch(request) {
    if (request.headers.get('Upgrade')?.toLowerCase() === 'websocket') {
      const ticket = String(new URL(request.url).searchParams.get('ticket') || '').trim();
      const ticketKey = `ticket:${ticket}`;
      const record = ticket ? await this.state.storage.get(ticketKey) : null;
      if (ticket) await this.state.storage.delete(ticketKey);
      if (!record || Number(record.expiresAt || 0) < Date.now()) {
        return new Response('Realtime ticket is invalid or expired', { status: 401 });
      }
      const pair = new WebSocketPair();
      const [client, server] = Object.values(pair);
      this.state.acceptWebSocket(server);
      server.serializeAttachment({
        email: extractEmailAddress(record.email || ''),
        connectedAt: Date.now(),
      });
      return new Response(null, { status: 101, webSocket: client });
    }
    if (request.method === 'POST' && new URL(request.url).pathname === '/ticket') {
      const body = await request.json().catch(() => ({}));
      const email = extractEmailAddress(body.email || '');
      if (!email) return json({ error: 'Ticket user is required' }, 400);
      const ticket = crypto.randomUUID().replaceAll('-', '') + crypto.randomUUID().replaceAll('-', '');
      const expiresAt = Date.now() + 60_000;
      await this.state.storage.put(`ticket:${ticket}`, { email, expiresAt });
      return new Response(JSON.stringify({ ticket, expiresAt }), {
        headers: { 'Content-Type': 'application/json' },
      });
    }
    if (request.method === 'POST' && new URL(request.url).pathname === '/broadcast') {
      const payload = await request.text();
      for (const socket of this.state.getWebSockets()) {
        try { socket.send(payload); } catch {}
      }
      return new Response(null, { status: 204 });
    }
    return new Response('Not found', { status: 404 });
  }

  webSocketMessage(socket, message) {
    if (message === 'ping') socket.send('pong');
  }

  webSocketClose(socket, code, reason) {
    socket.close(code, reason);
  }
}

// ── Router ────────────────────────────────────────────────────────────────────
async function routeRequest(request, env) {
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS });
    }
    const path = new URL(request.url).pathname.replace(/\/$/, '') || '/';
    if (path === '/environment' && request.method === 'GET') {
      return json({
        environment: isStaging(env) ? 'staging' : 'production',
        externalEffectsEnabled: externalEffectsAllowed(env),
      });
    }
    if (path === '/staging/realtime-ticket') return handleStagingRealtimeTicket(request, env);
    if (path === '/staging/realtime') return handleStagingRealtimeSocket(request, env);
    if (path === '/staging/tasks') return handleStagingTasks(request, env);
    if (isStaging(env) && STAGING_EXTERNAL_PATHS.has(path)) {
      return stagingSafetyResponse(path);
    }
    const proofMatch = path.match(/^\/p\/([A-Za-z0-9_%-]+)/);
    if (proofMatch && request.method === 'GET') return handleProofRedirect(request, env, proofMatch[1]);
    if (path === '/data') return handleData(request, env);
    if (path === '/shared-workflow-sync') return withStagingRealtimeBroadcast(request, env, handleSharedWorkflowSync(request, env));
    if (path === '/departments') return handleDepartments(request, env);
    if (path === '/department-assignment') return handleDepartments(request, env);
    if (path === '/task-edit-lock') return handleTaskEditLock(request, env);
    if (path === '/notify') return withStagingRealtimeBroadcast(request, env, handleNotify(request, env));
    if (path === '/assignments') return handleAssignments(request, env);

    if (request.method !== 'POST') {
      return json({ error: 'Only POST allowed' }, 405);
    }

    if (path === '/todo') return handleTodo(request, env);
    if (path === '/todo-update') return handleTodoUpdate(request, env);
    if (path === '/poll-completions') return handlePollCompletions(request, env);
    if (path === '/proof-task') return handleProofTask(request, env);
    if (path === '/proof-submit') return handleProofSubmit(request, env);
    if (path === '/attachment-summary') return handleAttachmentSummary(request, env);
    if (path === '/assignment') return handleCreateAssignment(request, env);
    if (path === '/assignment-status') return withStagingRealtimeBroadcast(request, env, handleAssignmentStatus(request, env));
    if (path === '/assignment-alert') return withStagingRealtimeBroadcast(request, env, handleAssignmentAlert(request, env));
    if (path === '/assignment-alert-clear') return withStagingRealtimeBroadcast(request, env, handleAssignmentAlertClear(request, env));
    if (path === '/assignment-seen') return handleAssignmentSeen(request, env);
    if (path === '/assignment-cancel') return withStagingRealtimeBroadcast(request, env, handleAssignmentCancel(request, env));
    return handleSummary(request, env);
}

export default {
  async fetch(request, env) {
    const response = await routeRequest(request, env);
    // A WebSocket upgrade carries a Cloudflare-specific webSocket property.
    // Reconstructing that 101 Response to add CORS headers strips the socket
    // and leaves the browser reconnecting forever. WebSockets validate Origin
    // before routing, so return successful upgrades exactly as provided.
    if (response.status === 101) return response;
    return withRequestCors(response, request);
  },
};
