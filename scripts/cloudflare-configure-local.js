const {app,net,safeStorage} = require('electron');
const path = require('node:path');
const {importIntegrationEnv} = require('../src/assessment/integrationConfig');
const {protectTeacherToken,revealTeacherToken} = require('../src/assessment/teacherCredentials');
const {AssessmentApiClient} = require('../src/assessment/assessmentApiClient');
app.setName(require('../package.json').name);
app.whenReady().then(async()=>{
  const userData=app.getPath('userData');
  const result=await importIntegrationEnv(path.resolve(process.argv[2]),path.join(userData,'repetierer.env'),{
    legacyBrevoPath:path.join(userData,'brevo-smtp.env'),
    protectTeacherToken:token=>protectTeacherToken(token,safeStorage),
    verify:async({assessment})=>{
      if(assessment.teacherToken.startsWith('encrypted:')) assessment.teacherToken=revealTeacherToken(assessment.teacherToken,safeStorage);
      await new AssessmentApiClient(assessment,{fetch:net.fetch}).checkStatus();
    }
  });
  console.log(JSON.stringify({ok:true,userData,encrypted:true,originalDeleted:result.originalDeleted}));app.quit();
}).catch(error=>{console.error(error.message);app.exit(1);});
