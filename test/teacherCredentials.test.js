const test = require('node:test');
const assert = require('node:assert/strict');
const {protectTeacherToken,revealTeacherToken} = require('../src/assessment/teacherCredentials');
const token='rt_'+'a'.repeat(43);
const storage={isEncryptionAvailable:()=>true,encryptString:s=>Buffer.from('sealed:'+s),decryptString:b=>b.toString().slice(7)};
test('Lehrerzugänge benötigen den Betriebssystem-Schlüsselspeicher',()=>{
  assert.throws(()=>protectTeacherToken(token,{isEncryptionAvailable:()=>false}),/nicht verfügbar/);
  assert.throws(()=>protectTeacherToken(token,{isEncryptionAvailable:()=>true,getSelectedStorageBackend:()=> 'basic_text'}),/nicht verfügbar/);
  assert.throws(()=>revealTeacherToken(token,storage),/erneut importieren/);
  const protectedToken=protectTeacherToken(token,storage);
  assert.equal(protectedToken.includes(token),false);
  assert.equal(revealTeacherToken(protectedToken,storage),token);
});
