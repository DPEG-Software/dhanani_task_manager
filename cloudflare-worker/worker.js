// DPEG AI Summarize — Cloudflare Worker
// Receives email text + Microsoft MSAL token from the app,
// validates the token belongs to DPEG, calls Groq, returns summary.

const ALLOWED_ORIGIN = 'https://dpeg-software.github.io';
const DPEG_TENANT_ID = '9152bf5c-22ff-4e4a-8624-784a2d243006';

const CORS = {
  'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
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

export default {
  async fetch(request, env) {

    // ── CORS preflight ───────────────────────────────────────────
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS });
    }

    if (request.method !== 'POST') {
      return json({ error: 'Only POST allowed' }, 405);
    }

    // ── Validate Microsoft token ─────────────────────────────────
    const auth = request.headers.get('Authorization') || '';
    const token = auth.startsWith('Bearer ') ? auth.slice(7) : '';
    if (!token) return json({ error: 'Missing authorization token' }, 401);

    const claims = decodeToken(token);
    if (!claims)            return json({ error: 'Invalid token'        }, 401);
    if (claims.tid !== DPEG_TENANT_ID)
                            return json({ error: 'Wrong tenant'         }, 403);
    if (claims.exp && Date.now() / 1000 > claims.exp)
                            return json({ error: 'Token expired'        }, 401);

    // ── Parse request body ───────────────────────────────────────
    let body;
    try { body = await request.json(); }
    catch { return json({ error: 'Invalid JSON body' }, 400); }

    const {
      subject      = '',
      emailText    = '',
      senderName   = '',
      messageCount = 1,
    } = body;

    if (!emailText && !subject) {
      return json({ error: 'Provide emailText or subject' }, 400);
    }

    // ── Build prompt ─────────────────────────────────────────────
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

    // ── Call Groq ────────────────────────────────────────────────
    const groqRes = await fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${env.GROQ_API_KEY}`,
        'Content-Type': 'application/json',
      },
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
  },
};
