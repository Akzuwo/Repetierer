// Provision/revoke an application-specific teacher credential. Never prints the token.
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { spawnSync } = require('node:child_process');
const root = path.join(__dirname, '..');
const privateDir = path.join(root, '.cloudflare-private');
const wrangler = path.join(root, 'node_modules/wrangler/bin/wrangler.js');
const args = process.argv.slice(2);
const staging = args.includes('--staging');
const revoke = args.includes('--revoke');
const rotate = args.includes('--rotate');
const value = args.find(v => !v.startsWith('--'));
if (!value || ((revoke || rotate) && !/^[a-f0-9-]{36}$/.test(value)) || (revoke && rotate)) throw new Error('Aufruf: node scripts/cloudflare-teacher.js "Lehrer-Kürzel" [--staging] oder --revoke/--rotate <Lehrer-ID>');
fs.mkdirSync(privateDir, {recursive:true});
const sqlFile = path.join(privateDir, `teacher-${crypto.randomUUID()}.sql`);
const id = revoke || rotate ? value : crypto.randomUUID();
const quote = v => "'" + v.replace(/'/g,"''") + "'";
let envFile;
if (revoke) {
  const now = Date.now();
  fs.writeFileSync(sqlFile, `UPDATE teachers SET revoked_at = ${now} WHERE id = ${quote(id)}; SELECT changes() AS changed; UPDATE rounds SET closed_at = COALESCE(closed_at, ${now}), expires_at = MIN(expires_at, ${now + 30*86400000}) WHERE teacher_id = ${quote(id)};`, {mode:0o600});
} else {
  const token = 'rt_' + crypto.randomBytes(32).toString('base64url');
  const hash = crypto.createHash('sha256').update(token).digest('hex');
  const expires = Date.now() + 365*86400000;
  const sql = rotate ? `UPDATE teachers SET token_hash = ${quote(hash)}, expires_at = ${expires}, revoked_at = NULL WHERE id = ${quote(id)}; SELECT changes() AS changed;` : `INSERT INTO teachers (id, token_hash, label, expires_at) VALUES (${quote(id)}, ${quote(hash)}, ${quote(value.slice(0,80))}, ${expires});`;
  fs.writeFileSync(sqlFile, sql, {mode:0o600});
  envFile = path.join(privateDir, `teacher-${id}${rotate ? '-'+Date.now() : ''}.env`);
  const api = staging ? process.env.ASSESSMENT_STAGING_API_URL : 'https://timolu.ch/app/services/repetierer/api';
  if (!api || !api.startsWith('https://')) throw new Error('Für Staging ASSESSMENT_STAGING_API_URL setzen.');
  fs.writeFileSync(envFile, `ASSESSMENT_API_URL=${api}\n${staging?'':'ASSESSMENT_FALLBACK_API_URL=https://akzuwo.ch/app/services/repetierer/api\n'}ASSESSMENT_TEACHER_TOKEN=${token}\n`, {mode:0o600,flag:'wx'});
}
const cliArgs = [wrangler, 'd1', 'execute', staging ? 'repetierer-assessment-staging' : 'repetierer-assessment', '--config', 'cloudflare/wrangler.jsonc', '--remote', '--file', sqlFile, '--json'];
if (staging) cliArgs.push('--env','staging');
const result = spawnSync(process.execPath, cliArgs, {cwd:root, env:{...process.env,NODE_USE_SYSTEM_CA:'1',WRANGLER_SEND_METRICS:'false'}, encoding:'utf8', windowsHide:true});
// The SQL contains only a token hash, but keep it out of diagnostic output too.
if (result.status !== 0) throw new Error('Lehrer-Provisionierung fehlgeschlagen. Datei bleibt für Diagnose erhalten; Wrangler-Anmeldung und D1-Zugriff prüfen.');
if (revoke || rotate) {
  // File-import results do not contain SELECT rows; verify with a separate query.
  const checkArgs = [wrangler, 'd1', 'execute', staging ? 'repetierer-assessment-staging' : 'repetierer-assessment', '--config', 'cloudflare/wrangler.jsonc', '--remote', '--json', '--command', `SELECT COUNT(*) AS count FROM teachers WHERE id = ${quote(id)} AND revoked_at IS ${revoke ? 'NOT ' : ''}NULL;`];
  if (staging) checkArgs.push('--env','staging');
  const check = spawnSync(process.execPath, checkArgs, {cwd:root,env:{...process.env,NODE_USE_SYSTEM_CA:'1',WRANGLER_SEND_METRICS:'false'},encoding:'utf8',windowsHide:true});
  if (check.status !== 0 || !JSON.parse(check.stdout).some(item=>item.results?.some(row=>row.count===1))) throw new Error('Lehrer-ID nicht gefunden oder Zustandsprüfung fehlgeschlagen.');
}
fs.unlinkSync(sqlFile);
console.log(JSON.stringify({teacherId:id, revoked:revoke, rotated:rotate, credentialFile:envFile || null}));
