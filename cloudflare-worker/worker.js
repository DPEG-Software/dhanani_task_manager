// DPEG AI Summarize + To Do — Cloudflare Worker
// POST /        → AI email summary (Groq, validated via MSAL user token)
// POST /todo    → Create Microsoft To Do task for a DPEG recipient (app credentials)

const ALLOWED_ORIGIN  = 'https://dpeg-software.github.io';
const DPEG_TENANT_ID  = '9152bf5c-22ff-4e4a-8624-784a2d243006';
const AZURE_CLIENT_ID = '8d523e65-0163-49c7-881b-407c0222527e';

const CORS = {
  'Access-Control-Allow-Origin':  ALLOWED_ORIGIN,
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
};

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...CORS, 'Content-Type': 'application/json' },
  });
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

  const { recipientEmail, title, summary = '', priority = 'Normal', date } = body;
  if (!recipientEmail || !title) {
    return json({ error: 'recipientEmail and title are required' }, 400);
  }
  if (!recipientEmail.toLowerCase().includes('@dhananipeg.com')) {
    return json({ error: 'Only @dhananipeg.com addresses are supported' }, 403);
  }

  // Get app-level token (Tasks.ReadWrite.All Application permission)
  let appToken;
  try { appToken = await getAppToken(env); }
  catch (err) { return json({ error: 'Could not acquire app token', detail: err.message }, 502); }

  // Find recipient's default To Do list
  const listsRes = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(recipientEmail)}/todo/lists`,
    { headers: { Authorization: `Bearer ${appToken}` } }
  );
  if (!listsRes.ok) {
    const err = await listsRes.text().catch(() => '');
    return json({ error: 'Cannot access recipient To Do', detail: err }, listsRes.status);
  }
  const listsData = await listsRes.json();
  const defaultList =
    (listsData.value || []).find(l => l.isDefaultList) ||
    (listsData.value || [])[0];
  if (!defaultList) return json({ error: 'No task list found for recipient' }, 404);

  // Build the To Do task
  const task = {
    title,
    body: {
      content: `Assigned to you via DPEG Task Manager\n\n${summary.replace(/[•*▾▲◆]/g, '').slice(0, 600)}`,
      contentType: 'text',
    },
    importance: String(priority).toLowerCase() === 'high' ? 'high' : 'normal',
    status: 'notStarted',
  };
  if (date) {
    const d = new Date(date);
    if (!isNaN(d)) task.dueDateTime = { dateTime: d.toISOString(), timeZone: 'UTC' };
  }

  const taskRes = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(recipientEmail)}/todo/lists/${defaultList.id}/tasks`,
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

  return json({ success: true });
}

// ── / endpoint (existing AI summary) ─────────────────────────────────────────
async function handleSummary(request, env) {
  const { error, status, claims } = validateUserToken(request);
  if (error) return json({ error }, status);

  let body;
  try { body = await request.json(); }
  catch { return json({ error: 'Invalid JSON body' }, 400); }

  const { subject = '', emailText = '', senderName = '', messageCount = 1 } = body;
  if (!emailText && !subject) {
    return json({ error: 'Provide emailText or subject' }, 400);
  }

  const prompt = `You are an executive assistant briefing a busy property investment executive at DPEG (Dhanani Private Equity Group).

Email subject: "${subject}"${senderName ? `\nFrom: ${senderName}` : ''}${messageCount > 1 ? `\nThread: ${messageCount} messages` : ''}

Email content:
${emailText.slice(0, 3500)}

Write a concise executive briefing using ONLY these bullet labels. Skip any bullet that has no meaningful content:

• About: [One sentence — what is this email actually about? Be specific.]
• Action needed: [One sentence — what must DPEG do or decide? Use a direct verb.]
• Key info: [One concrete fact — a dollar amount, date, deadline, property name, or percentage. Skip if none.]
• Latest: [One sentence — the most recent development or what was last said.]

Rules: Never say "the email discusses". State specifics directly. Max 200 words.`;

  const groqRes = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${env.GROQ_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'llama-3.3-70b-versatile',
      messages: [{ role: 'user', content: prompt }],
      max_tokens: 250,
      temperature: 0.2,
    }),
  });

  if (!groqRes.ok) {
    const err = await groqRes.text().catch(() => '');
    return json({ error: 'Groq call failed', detail: err }, 502);
  }

  const groqData = await groqRes.json();
  const summary  = groqData.choices?.[0]?.message?.content?.trim() || '';
  return json({ summary });
}

// ── Router ────────────────────────────────────────────────────────────────────
export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS });
    }
    if (request.method !== 'POST') {
      return json({ error: 'Only POST allowed' }, 405);
    }

    const path = new URL(request.url).pathname.replace(/\/$/, '') || '/';
    if (path === '/todo') return handleTodo(request, env);
    return handleSummary(request, env);
  },
};
