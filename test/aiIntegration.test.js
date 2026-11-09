const test=require('node:test'),assert=require('node:assert/strict');
const {createAiIntegration}=require('../src/ai/integration');
test('KI entfernen sperrt Statusabfragen und Chat, bewahrt Ollama und kann bestätigt wieder eingerichtet werden',async()=>{
  let enabled=true,calls=0;
  const service={status:async()=>{calls++;return {state:'ready'};},chat:async()=>{calls++;return 'Antwort';},setup:async()=>({state:'ready'})};
  const integration=createAiIntegration({service,getEnabled:()=>enabled,setEnabled:value=>enabled=value});
  assert.equal((await integration.status()).state,'ready');
  integration.remove();assert.equal(enabled,false);const before=calls;
  assert.equal((await integration.status()).state,'disabled');await assert.rejects(integration.chat([]),/entfernt/);assert.equal(calls,before);
  assert.throws(integration.requireEnabled,/entfernt/);
  await integration.setup(()=>{});assert.equal(enabled,true);assert.equal(await integration.chat([]),'Antwort');
});
test('Entfernen verwirft laufende KI-Antworten; fehlgeschlagene Einrichtung aktiviert keine Integration',async()=>{
  let enabled=true,release;
  const integration=createAiIntegration({getEnabled:()=>enabled,setEnabled:v=>enabled=v,service:{chat:()=>new Promise(r=>release=r),setup:async()=>{throw new Error('Download fehlgeschlagen');}}});
  const pending=integration.chat([]);integration.remove();release('späte Antwort');await assert.rejects(pending,/entfernt/);
  await assert.rejects(integration.setup(),/Download/);assert.equal(enabled,false);
});
