// Runs on Cloudflare's Linux build image without an Electron GUI or cloud credentials.
const {spawnSync} = require('node:child_process');
const path = require('node:path');
const root = path.join(__dirname,'..');
const env = {...process.env,REPETIERER_BACKEND_ONLY:'1',WRANGLER_SEND_METRICS:'false'};
const tests = ['test/cloudflareE2E.test.js','test/assessmentApiClient.test.js','test/assessmentStore.test.js','test/delivery.test.js','test/integrationConfig.test.js'];
for (const args of [
  ['--test','--test-concurrency=1',...tests],
  [path.join(root,'node_modules/wrangler/bin/wrangler.js'),'deploy','--dry-run','--config','cloudflare/wrangler.jsonc','--env=']
]) {
  const result=spawnSync(process.execPath,args,{cwd:root,env,stdio:'inherit',windowsHide:true});
  if(result.status!==0) process.exit(result.status || 1);
}
