// Workbook-owned homework columns. Existing columns and repetition limits stay intact.
const empty = value => value === null || value === undefined || value === '';
const validGrade = value => typeof value === 'number' && Number.isFinite(value) && value >= 1 && value <= 6;
const labels = {grade:'Hausaufgaben Note',date:'Hausaufgaben Datum',weight:'Hausaufgaben Gewicht',meta:'Hausaufgaben Eintrag'};
function layout(sheet) {
  const slots=new Map();let average,rounded;
  for(const row of [3,5]) sheet.getRow(row).eachCell((cell,column)=>{
    const text=String(cell.value || '');
    if(text==='Ø Repetition + Hausaufgaben')average={row,column};
    if(text==='Hausaufgaben Gesamtnote gerundet')rounded={row,column};
    for(const [key,label] of Object.entries(labels)) {
      const match=new RegExp(`^${label} (\\d+)$`).exec(text);
      if(match){const index=Number(match[1]);const slot=slots.get(index)||{index,row};if(slot[key])throw new Error('Doppelte Hausaufgabenspalten.');slot[key]=column;slots.set(index,slot);}
    }
  });
  const result=[...slots.values()].sort((a,b)=>a.index-b.index);
  if(result.some((slot,index)=>slot.index!==index+1 || Object.keys(labels).some(key=>!slot[key])) || (!!result.length !== !!average) || (!!average !== !!rounded))throw new Error('Die Hausaufgabenspalten sind unvollständig. Bitte die Originaldatei prüfen.');
  return {slots:result,average,rounded};
}
function inspect(sheet) {
  const current=layout(sheet);
  const start=Math.max(23,sheet.columnCount+1);
  if(!current.slots.length && start+25>16384)throw new Error('Kein Platz für Hausaufgabenspalten.');
  return {name:sheet.name,needsMigration:!current.slots.length,start,columns:26};
}
function header(sheet,column,text,hidden=false) {
  const row=sheet.getCell('A3').value ? 3 : 5;
  const cell=sheet.getCell(row,column);cell.style=JSON.parse(JSON.stringify(sheet.getCell(row,11).style));cell.value=text;
  sheet.getColumn(column).width=text.includes('Ø')||text.includes('gerundet')?26:22;
  sheet.getColumn(column).hidden=hidden;
  if(row===3)sheet.mergeCells(3,column,4,column);
  for(let r=6;r<=30;r++)sheet.getCell(r,column).style=JSON.parse(JSON.stringify(sheet.getCell(r,11).style));
}
function addSlot(sheet,index,start) {
  for(const [offset,[key,label]] of Object.entries(labels).entries()) {
    const column=start+Number(offset);header(sheet,column,`${label} ${index}`,key==='meta'||key==='weight');
    for(let row=6;row<=30;row++) {
      const cell=sheet.getCell(row,column);
      if(key==='weight'){cell.numFmt='0.0';}
      if(key==='grade'){cell.numFmt='0.0';cell.dataValidation={type:'decimal',operator:'between',allowBlank:true,formulae:[1,6],showErrorMessage:true,error:'Note zwischen 1 und 6 eingeben.'};}
      if(key==='date')cell.numFmt='dd.mm.yyyy';
    }
  }
}
function migrateWorksheet(sheet) {
  const plan=inspect(sheet);if(!plan.needsMigration)return plan;
  header(sheet,plan.start,'Ø Repetition + Hausaufgaben');header(sheet,plan.start+1,'Hausaufgaben Gesamtnote gerundet');
  for(let index=0;index<6;index++)addSlot(sheet,index+1,plan.start+2+index*4);
  refreshFormulas(sheet);return plan;
}
function history(sheet,row) {
  return layout(sheet).slots.filter(slot=>!empty(sheet.getCell(row,slot.grade).value)).map(slot=>{
    const grade=sheet.getCell(row,slot.grade).value,weight=sheet.getCell(row,slot.weight).value;
    if(!validGrade(grade)||typeof weight!=='number'||!Number.isFinite(weight)||weight<=0||weight>5)throw new Error('Eine Hausaufgabennote oder ihr Gewicht ist ungültig.');
    let meta={};const raw=sheet.getCell(row,slot.meta).value;
    if(!empty(raw)){try{meta=JSON.parse(raw);if(!meta || typeof meta!=='object' || Array.isArray(meta))throw new Error();}catch{throw new Error('Ein Hausaufgabeneintrag ist beschädigt.');}}
    return {index:slot.index,columns:{...slot},grade,weight,date:sheet.getCell(row,slot.date).value,id:meta.id,awardExtraJoker:meta.awardExtraJoker===true};
  });
}
function weightedAverage(sheet,row) {
  const grades=[];for(let c=2;c<=7;c++){const raw=sheet.getCell(row,c).value;const grade=typeof raw==='string'&&raw.trim()?Number(raw.replace(',','.')):raw;if(validGrade(grade))grades.push({grade,weight:1});}
  grades.push(...history(sheet,row));
  return grades.length ? grades.reduce((sum,e)=>sum+e.grade*e.weight,0)/grades.reduce((sum,e)=>sum+e.weight,0) : null;
}
function refreshFormulas(sheet) {
  const current=layout(sheet);if(!current.slots.length)return;
  for(let row=6;row<=30;row++) {
    const refs=current.slots.map(slot=>({g:sheet.getCell(row,slot.grade).address,w:sheet.getCell(row,slot.weight).address}));
    const numerator=`SUM(B${row}:G${row})+${refs.map(({g,w})=>`IF(ISNUMBER(${g}),${g}*${w},0)`).join('+')}`;
    const denominator=`COUNT(B${row}:G${row})+${refs.map(({g,w})=>`IF(ISNUMBER(${g}),${w},0)`).join('+')}`;
    const average=weightedAverage(sheet,row);
    sheet.getCell(row,current.average.column).value={formula:`IF((${denominator})=0,"",(${numerator})/(${denominator}))`,result:average ?? ''};
    const address=sheet.getCell(row,current.average.column).address;
    sheet.getCell(row,current.rounded.column).value={formula:`IF(${address}="","",ROUND(${address},1))`,result:average===null?'':Math.round(average*10)/10};
    sheet.getCell(row,current.average.column).numFmt='0.00';sheet.getCell(row,current.rounded.column).numFmt='0.0';
  }
}
function append(sheet,row,entry) {
  if(!validGrade(entry.grade)||!Number.isFinite(entry.weight)||entry.weight<=0||entry.weight>5||!entry.id)throw new Error('Ungültiger Hausaufgabeneintrag.');
  const current=layout(sheet);if(!current.slots.length)throw new Error('Bitte zuerst die Hausaufgabenspalten anlegen.');
  if(history(sheet,row).some(item=>item.id===(entry.homeworkId||entry.id)))return false;
  let slot=current.slots.find(item=>empty(sheet.getCell(row,item.grade).value));
  if(!slot){if(sheet.columnCount+4>16384)throw new Error('Keine weiteren Excelspalten verfügbar.');addSlot(sheet,current.slots.length+1,sheet.columnCount+1);slot=layout(sheet).slots.at(-1);}
  sheet.getCell(row,slot.grade).value=entry.grade;sheet.getCell(row,slot.date).value=entry.date;sheet.getCell(row,slot.weight).value=entry.weight;
  sheet.getCell(row,slot.meta).value=JSON.stringify({id:entry.homeworkId||entry.id,awardExtraJoker:!!entry.awardExtraJoker});
  if(entry.awardExtraJoker)sheet.getCell(row,8).value=Number(sheet.getCell(row,8).value||0)+1;
  refreshFormulas(sheet);return true;
}
function undo(sheet,row,entry) {
  const slot=history(sheet,row).find(item=>item.id===(entry.homeworkId||entry.id));
  if(!slot || slot.grade!==entry.grade)throw new Error('Der Hausaufgabeneintrag wurde in Excel verändert.');
  for(const key of ['grade','date','weight','meta'])sheet.getCell(row,slot.columns[key]).value=null;
  if(slot.awardExtraJoker)sheet.getCell(row,8).value=Math.max(0,Number(sheet.getCell(row,8).value||0)-1);
  refreshFormulas(sheet);
}
module.exports={layout,inspect,migrateWorksheet,history,weightedAverage,refreshFormulas,append,undo,validGrade};
