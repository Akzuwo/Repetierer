const test=require('node:test'),assert=require('node:assert/strict');
const {validateDeadline,deadlineEndZurich,deadlineBounds,DAY}=require('../src/assessment/deadlinePolicy');
const {createAutoSync,SYNC_INTERVAL_MS}=require('../src/assessment/autoSync');
test('Fristen: echte Kalendertage, Zürich Sommer/Winter, Vergangenheit und 90-Tage-Grenze',()=>{
  const now=Date.parse('2026-11-06T10:00:00Z');
  assert.equal(deadlineEndZurich('2026-12-18'),'2026-12-18T22:59:59.000Z');
  assert.equal(deadlineEndZurich('2027-06-18'),'2027-06-18T21:59:59.000Z');
  for(const date of ['2027-02-29','2026-02-30','2099-01-01','2026-11-05'])assert.throws(()=>validateDeadline(date,now));
  const bounds=deadlineBounds(now);
  assert.equal(validateDeadline(bounds.min,now),bounds.min);
  assert.equal(validateDeadline(bounds.max,now),bounds.max);
  assert.ok(Date.parse(deadlineEndZurich(bounds.max))<=now+90*DAY);
  assert.equal(validateDeadline('',now),null);assert.throws(()=>validateDeadline('',now,true));
});
test('Auto-Sync läuft alle 30 Sekunden, überlappt nicht, überlebt Fehler und teilt die Warteschlange mit manuellen Aktionen',async()=>{
  const target={filePath:'fixture.xlsx',className:'A',roundId:'round'};
  let interval,timer,release,requests=0,failures=0,successes=0;
  const sync=createAutoSync({getTargets:()=>[target],sync:async()=>{requests++;if(requests===1)await new Promise(r=>release=r);if(requests===2)throw new Error('offline');return 'ok';},onResult:()=>successes++,onError:()=>failures++,setIntervalFn:(callback,ms)=>{timer=callback;interval=ms;return 17;},clearIntervalFn:id=>assert.equal(id,17)});
  assert.equal(interval,SYNC_INTERVAL_MS);assert.equal(interval,30000);assert.equal(typeof timer,'function');
  const first=sync.tick();await new Promise(r=>setImmediate(r));
  await sync.tick();assert.equal(requests,1);release();await first;
  await sync.tick();assert.equal(failures,1);await sync.tick();assert.equal(successes,2);
  const sequence=[];let completeManual;
  const manual=sync.run(target,async()=>{sequence.push('manual');await new Promise(r=>completeManual=r);});
  const queued=sync.run(target,()=>sequence.push('close'));
  await new Promise(r=>setImmediate(r));await sync.tick();assert.equal(requests,3);
  completeManual();await Promise.all([manual,queued]);assert.deepEqual(sequence,['manual','close']);
  sync.stop();await sync.tick();assert.equal(requests,3);
});
