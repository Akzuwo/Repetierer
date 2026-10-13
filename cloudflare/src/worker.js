import html from '../web/index.html';
import css from '../web/survey.css';
import js from '../web/survey.js';

const BASE = '/app/services/repetierer/';
const DAY = 86400000;
const headers = {
  'Cache-Control': 'no-store, no-transform',
  'Content-Security-Policy': "default-src 'none'; script-src 'self'; style-src 'self'; font-src 'self'; connect-src 'self'; img-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'",
  'Referrer-Policy': 'no-referrer',
  'X-Content-Type-Options': 'nosniff',
  'Permissions-Policy': 'camera=(), microphone=(), geolocation=()',
  'Strict-Transport-Security': 'max-age=31536000'
};

function json(data, status = 200) {
  return new Response(JSON.stringify({ success: status < 400, ...data }), { status, headers: { ...headers, 'Content-Type': 'application/json; charset=utf-8' } });
}
function fail(status, error) { throw Object.assign(new Error(error), { status, publicCode: error }); }
export async function hash(value) {
  return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(value))), x => x.toString(16).padStart(2, '0')).join('');
}
export function normalizeName(value) {
  return value.normalize('NFKC').replace(/\p{Cf}/gu, '').trim().replace(/\s+/gu, ' ').toLocaleLowerCase('de-CH');
}
function text(value, max, required = false) {
  if (typeof value !== 'string' || value.length > max || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(value) || (required && !value.trim())) fail(400, 'invalid_input');
  return value.trim();
}
async function body(request) {
  if (!request.headers.get('Content-Type')?.startsWith('application/json')) fail(415, 'json_required');
  const reader = request.body?.getReader();
  if (!reader) fail(400, 'invalid_input');
  let size = 0;
  const parts = [];
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    size += value.byteLength;
    if (size > 32768) { await reader.cancel(); fail(413, 'payload_too_large'); }
    parts.push(value);
  }
  const bytes = new Uint8Array(size);
  let offset = 0;
  for (const part of parts) { bytes.set(part, offset); offset += part.length; }
  let data;
  try { data = JSON.parse(new TextDecoder().decode(bytes)); } catch { fail(400, 'invalid_json'); }
  if (!data || Array.isArray(data) || typeof data !== 'object') fail(400, 'invalid_input');
  return data;
}
async function teacher(request, env) {
  const token = request.headers.get('Authorization')?.match(/^Bearer (rt_[A-Za-z0-9_-]{43})$/)?.[1];
  if (!token) fail(401, 'unauthorized');
  const row = await env.DB.prepare('SELECT id FROM teachers WHERE token_hash = ? AND revoked_at IS NULL AND expires_at > ?').bind(await hash(token), Date.now()).first();
  if (!row) fail(401, 'unauthorized');
  return row.id;
}
async function ownedRound(env, id, owner) {
  const row = await env.DB.prepare('SELECT * FROM rounds WHERE id = ? AND teacher_id = ? AND expires_at > ?').bind(id, owner, Date.now()).first();
  if (!row) fail(404, 'round_not_found');
  return row;
}
function open(round) { return !round.closed_at && round.deadline > Date.now() && round.expires_at > Date.now(); }

async function route(request, env) {
  const url = new URL(request.url);
  if (!url.pathname.startsWith(BASE)) return json({ error: 'not_found' }, 404);
  const path = url.pathname.slice(BASE.length);
  if (request.method === 'GET' && ['', 'survey.css', 'survey.js', 'fonts/Geist-Variable.woff2'].includes(path)) {
    if (path.startsWith('fonts/')) {
      const asset = await env.ASSETS.fetch(new Request(new URL('/' + path, url), request));
      return new Response(asset.body, {status:asset.status, headers:{...headers, 'Content-Type':'font/woff2', 'Cache-Control':'public, max-age=86400'}});
    }
    return new Response(path === '' ? html : path.endsWith('.css') ? css : js, {headers:{...headers, 'Content-Type':path === '' ? 'text/html; charset=utf-8' : path.endsWith('.css') ? 'text/css; charset=utf-8' : 'application/javascript; charset=utf-8'}});
  }
  if (path === 'api/health' && request.method === 'GET') {
    await env.DB.prepare('SELECT 1 FROM teachers LIMIT 1').first();
    return json({ service: 'repetierer-assessment' });
  }
  const publicMatch = path.match(/^api\/public\/rounds\/([a-f0-9-]{36})(\/responses)?$/);
  if (publicMatch) {
    const id = publicMatch[1];
    const round = await env.DB.prepare('SELECT * FROM rounds WHERE id = ? AND expires_at > ?').bind(id, Date.now()).first();
    if (!round) fail(404, 'round_not_found');
    if (!publicMatch[2] && request.method === 'GET') return json({round:{ id, title:round.title, criteria:JSON.parse(round.criteria), deadline:new Date(round.deadline).toISOString(), expiresAt:new Date(round.expires_at).toISOString(), closed:!open(round) }});
    if (publicMatch[2] && request.method === 'POST') {
      const origin = request.headers.get('Origin');
      if (origin && origin !== url.origin) fail(403, 'origin_denied');
      if (!open(round)) fail(410, 'round_closed');
      const data = await body(request);
      const name = text(data.name, 200, true);
      if (!normalizeName(name)) fail(400, 'invalid_input');
      const email = text(data.email, 254, true).toLowerCase();
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) fail(400, 'invalid_email');
      if (!/^[A-Za-z0-9_-]{32,100}$/.test(data.browserToken || '')) fail(400, 'invalid_browser_token');
      const criteria = JSON.parse(round.criteria);
      if (!data.answers || typeof data.answers !== 'object' || Array.isArray(data.answers) || Object.keys(data.answers).length !== criteria.length) fail(400, 'invalid_answers');
      const answers = {};
      for (const criterion of criteria) {
        const value = data.answers[criterion.id];
        if (!Number.isInteger(value) || value < 1 || value > 5) fail(400, 'invalid_answers');
        answers[criterion.id] = value;
      }
      const additionalQuality = text(data.additionalQuality ?? '', 2000);
      const comment = text(data.comment ?? '', 2000);
      const now = Date.now();
      const responseId = crypto.randomUUID();
      let result;
      try {
        result = await env.DB.prepare(`INSERT INTO responses (id, round_id, name, name_key, email, browser_hash, answers, additional_quality, comment, submitted_at)
          SELECT ?, id, ?, ?, ?, ?, ?, ?, ?, ? FROM rounds
          WHERE id = ? AND closed_at IS NULL AND deadline > ? AND expires_at > ?
          AND (SELECT COUNT(*) FROM responses WHERE round_id = ?) < max_responses`).bind(responseId, name, normalizeName(name), email, await hash(id + ':' + data.browserToken), JSON.stringify(answers), additionalQuality, comment, now, id, now, now, id).run();
      } catch (error) {
        if (/UNIQUE constraint failed: responses\.(?:round_id|name_key|browser_hash)/.test(error.message)) fail(409, 'duplicate_response');
        throw error;
      }
      if (!result.meta.changes) fail(410, 'round_closed_or_full');
      return json({responseId}, 201);
    }
    fail(405, 'method_not_allowed');
  }
  if (path.startsWith('api/teacher/')) {
    const owner = await teacher(request, env);
    if (path === 'api/teacher/me' && request.method === 'GET') return json({teacherId:owner});
    const match = path.match(/^api\/teacher\/rounds\/([a-f0-9-]{36})(\/(?:responses|close))?$/);
    if (!match) fail(404, 'not_found');
    const id = match[1];
    if (!match[2] && request.method === 'PUT') {
      const data = await body(request);
      const title = text(data.title, 160, true);
      if (!Array.isArray(data.criteria) || data.criteria.length < 1 || data.criteria.length > 20) fail(400, 'invalid_criteria');
      const criteria = data.criteria.map(c => {
        if (!c || !/^[a-z0-9_]{1,64}$/.test(c.id) || !Array.isArray(c.scaleLabels) || c.scaleLabels.length !== 5) fail(400, 'invalid_criteria');
        return {id:c.id, label:text(c.label, 250, true), scaleLabels:c.scaleLabels.map(v=>text(v, 80, true))};
      });
      if (new Set(criteria.map(c=>c.id)).size !== criteria.length) fail(400, 'invalid_criteria');
      const now = Date.now();
      const existing = await env.DB.prepare('SELECT * FROM rounds WHERE id = ?').bind(id).first();
      const deadline = data.deadline ? Date.parse(data.deadline) : (existing?.deadline ?? now + 30 * DAY);
      const max = data.maxResponses;
      if (!Number.isInteger(max) || max < 1 || max > 200 || !Number.isFinite(deadline) || (!existing && (deadline <= now || deadline > now + 90 * DAY))) fail(400, 'invalid_input');
      const encoded = JSON.stringify(criteria);
      if (existing) {
        if (existing.teacher_id !== owner || existing.expires_at <= now) fail(404, 'round_not_found');
        if (existing.title !== title || existing.criteria !== encoded || existing.deadline !== deadline || existing.max_responses !== max) fail(409, 'round_conflict');
        return json({roundId:id});
      }
      const count = await env.DB.prepare('SELECT COUNT(*) AS n FROM rounds WHERE teacher_id = ? AND expires_at > ?').bind(owner, now).first();
      if (count.n >= 200) fail(429, 'round_limit');
      await env.DB.prepare('INSERT INTO rounds (id, teacher_id, title, criteria, deadline, expires_at, created_at, max_responses) VALUES (?, ?, ?, ?, ?, ?, ?, ?)').bind(id, owner, title, encoded, deadline, now + 90 * DAY, now, max).run();
      return json({roundId:id}, 201);
    }
    const round = await ownedRound(env, id, owner);
    if (match[2] === '/responses' && request.method === 'GET') {
      const cursor = Number(url.searchParams.get('cursor') || 0);
      if (!Number.isSafeInteger(cursor) || cursor < 0) fail(400, 'invalid_cursor');
      const rows = await env.DB.prepare('SELECT * FROM responses WHERE round_id = ? AND seq > ? ORDER BY seq LIMIT 100').bind(id, cursor).all();
      return json({closed:!open(round), responses:rows.results.map(r=>({responseId:r.id, roundId:id, timestamp:new Date(r.submitted_at).toISOString(), name:r.name, email:r.email, answers:JSON.parse(r.answers), additionalQuality:r.additional_quality, comment:r.comment})), nextCursor:rows.results.length === 100 ? rows.results.at(-1).seq : null});
    }
    if (match[2] === '/close' && request.method === 'POST') {
      const now = Date.now();
      await env.DB.prepare('UPDATE rounds SET closed_at = COALESCE(closed_at, ?), expires_at = MIN(expires_at, ?) WHERE id = ? AND teacher_id = ?').bind(now, (round.closed_at || now) + 30 * DAY, id, owner).run();
      return json({closed:true});
    }
    if (!match[2] && request.method === 'DELETE') {
      await env.DB.prepare('DELETE FROM rounds WHERE id = ? AND teacher_id = ?').bind(id, owner).run();
      return json({deleted:true});
    }
    fail(405, 'method_not_allowed');
  }
  fail(404, 'not_found');
}
export default {
  async fetch(request, env) {
    try { return await route(request, env); }
    catch (error) { return json({error:error.publicCode || 'internal_error'}, error.status || 500); }
  },
  async scheduled(controller, env) {
    await env.DB.prepare('DELETE FROM rounds WHERE expires_at <= ?').bind(Date.now()).run();
  }
};
