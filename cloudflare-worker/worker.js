// DPEG AI Summarize + To Do — Cloudflare Worker
// POST /        → AI email summary (Groq, validated via MSAL user token)
// POST /todo    → Create Microsoft To Do task for a DPEG recipient (app credentials)

const ALLOWED_ORIGIN  = 'https://dpeg-software.github.io';
const DPEG_TENANT_ID  = '9152bf5c-22ff-4e4a-8624-784a2d243006';
const AZURE_CLIENT_ID = '8d523e65-0163-49c7-881b-407c0222527e';

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

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...CORS, 'Content-Type': 'application/json' },
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
  if (!env.DPEG_DATA || !code) return new Response('Proof link not found', { status: 404, headers: CORS });
  const record = await env.DPEG_DATA.get(`${PROOF_LINK_PREFIX}${code}`, 'json');
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

// Validate the user's MSAL Bearer token (shared by both endpoints)
function validateUserToken(request) {
  const auth  = request.headers.get('Authorization') || '';
  const token = auth.startsWith('Bearer ') ? auth.slice(7) : '';
  if (!token) return { error: 'Missing authorization token', status: 401 };
  const claims = decodeToken(token);
  if (!claims) return { error: 'Invalid token', status: 401 };
  if (claims.tid !== DPEG_TENANT_ID) return { error: 'Wrong tenant', status: 403 };
  if (claims.exp && Date.now() / 1000 > claims.exp) return { error: 'Token expired', status: 401 };
  return { claims };
}

// Acquire an app-only token using client_credentials (uses Tasks.ReadWrite.All Application permission)
async function getAppToken(env) {
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
  return data.access_token;
}

// ── /todo endpoint ───────────────────────────────────────────────────────────
async function handleTodo(request, env) {
  const { error, status, claims } = validateUserToken(request);
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
  const { error, status, claims } = validateUserToken(request);
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
    const payload = {
      tasks: Array.isArray(body.tasks) ? body.tasks : [],
      archives: Array.isArray(body.archives) ? body.archives : [],
      staffConfig: ADMIN_EMAILS.has(userEmail) && body.staffConfig && typeof body.staffConfig === 'object'
        ? body.staffConfig
        : (existing.staffConfig && typeof existing.staffConfig === 'object' ? existing.staffConfig : {}),
      customNotes: Array.isArray(body.customNotes) ? body.customNotes : [],
      notifications: Array.isArray(body.notifications) ? body.notifications : [],
      updatedAt: new Date().toISOString(),
    };
    await env.DPEG_DATA.put(DATA_KEY, JSON.stringify(payload));
    return json({ success: true, updatedAt: payload.updatedAt });
  }

  return json({ error: 'Method not allowed' }, 405);
}

// ── / endpoint (existing AI summary) ─────────────────────────────────────────
async function handleSummary(request, env) {
  const { error, status, claims } = validateUserToken(request);
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
  const { error, status } = validateUserToken(request);
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
  const { error, status } = validateUserToken(request);
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
  const { error, status, claims } = validateUserToken(request);
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
  const { error, status, claims } = validateUserToken(request);
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
  const { error, status, claims } = validateUserToken(request);
  if (error) return json({ error }, status);
  if (!env.DPEG_DATA) return json({ error: 'DPEG_DATA KV binding is not configured' }, 501);

  if (request.method === 'GET') {
    const data = await env.DPEG_DATA.get(DATA_KEY, 'json') || {};
    return json({ notifications: Array.isArray(data.notifications) ? data.notifications : [] });
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
    // If declined, reset the recipient's To Do task back to notStarted
    if (body.result === 'declined' && body.todoListId && body.todoTaskId && body.recipientEmail) {
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
    notifications.push({
      id: `pn-${Date.now()}-${Math.random().toString(36).slice(2,7)}`,
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

// ── Router ────────────────────────────────────────────────────────────────────
export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS });
    }
    const path = new URL(request.url).pathname.replace(/\/$/, '') || '/';
    const proofMatch = path.match(/^\/p\/([A-Za-z0-9_-]+)$/);
    if (proofMatch && request.method === 'GET') return handleProofRedirect(request, env, proofMatch[1]);
    if (path === '/data') return handleData(request, env);
    if (path === '/notify') return handleNotify(request, env);

    if (request.method !== 'POST') {
      return json({ error: 'Only POST allowed' }, 405);
    }

    if (path === '/todo') return handleTodo(request, env);
    if (path === '/todo-update') return handleTodoUpdate(request, env);
    if (path === '/poll-completions') return handlePollCompletions(request, env);
    if (path === '/proof-task') return handleProofTask(request, env);
    if (path === '/proof-submit') return handleProofSubmit(request, env);
    if (path === '/attachment-summary') return handleAttachmentSummary(request, env);
    return handleSummary(request, env);
  },
};
