const test=require('node:test'),assert=require('node:assert/strict');
const {createHealthMonitor,HEALTH_INTERVAL_MS}=require('../src/assessment/healthMonitor');
const flush=()=>new Promise(resolve=>setImmediate(resolve));
test('startup check, five-minute main-process cadence, shared cached reads and no overlapping checks',async()=>{
 let tick,release,calls=0,cleared;const updates=[];
 const monitor=createHealthMonitor({check:async()=>{calls++;if(calls===1)await new Promise(resolve=>release=resolve);return {connection:{reachable:true},brevo:{mode:'cloud'}};},onUpdate:state=>updates.push(state),setIntervalFn:(fn,ms)=>{assert.equal(ms,300000);tick=fn;return 42;},clearIntervalFn:id=>cleared=id});
 assert.equal(HEALTH_INTERVAL_MS,300000);
 const first=monitor.start();await flush();const parallel=[monitor.get(),monitor.get(),tick()];assert.equal(calls,1);
 release();await Promise.all([first,...parallel]);assert.equal(updates.length,1);assert.equal((await monitor.get()).connection.reachable,true);assert.equal(calls,1);
 await tick();assert.equal(calls,2);assert.equal(updates.length,2);assert.ok(updates[1].checkedAt);
 monitor.stop();assert.equal(cleared,42);await tick();await monitor.start();assert.equal(calls,2);
});
test('failure is visible and a later periodic check recovers without stopping the monitor',async()=>{
 let tick,calls=0;const updates=[];
 const monitor=createHealthMonitor({check:async()=>{if(++calls===1)throw new Error('private details');return {connection:{reachable:true},brevo:{configured:true}};},onUpdate:state=>updates.push(state),setIntervalFn:fn=>{tick=fn;return 1;},clearIntervalFn:()=>{}});
 await monitor.start();assert.equal(updates[0].brevo.invalid,true);assert.doesNotMatch(JSON.stringify(updates),/private details/);
 await tick();assert.equal(updates[1].connection.reachable,true);assert.equal(updates[1].brevo.configured,true);monitor.stop();
});
test('config changes discard a slow old result and immediately recheck once, without overlap',async()=>{
 let release,calls=0,active=0,peak=0;const updates=[];
 const monitor=createHealthMonitor({check:async()=>{calls++;active++;peak=Math.max(peak,active);if(calls===1)await new Promise(resolve=>release=resolve);active--;return {connection:{source:calls===1?'old':'new'}};},onUpdate:state=>updates.push(state)});
 const old=monitor.get();await flush();const changed=monitor.invalidate();release();
 const results=await Promise.all([old,changed]);assert.equal(calls,2);assert.equal(peak,1);assert.equal(updates.length,1);assert.equal(results[0].connection.source,'new');assert.equal(results[1].connection.source,'new');monitor.stop();
});
test('quit suppresses late results and notifications',async()=>{
 let release,updates=0;
 const monitor=createHealthMonitor({check:()=>new Promise(resolve=>release=resolve),onUpdate:()=>updates++});
 const pending=monitor.get();monitor.stop();release({connection:{reachable:true}});await pending;assert.equal(updates,0);assert.equal(await monitor.get(),null);
});
