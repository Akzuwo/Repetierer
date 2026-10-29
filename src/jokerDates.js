// Keep legacy Q histories readable until the teacher explicitly migrates.
const FIRST = 17;
const MIN_COLUMNS = 6;
const title = index => `Joker Datum ${index}`;
const empty = value => value === null || value === undefined || value === '';
function dates(value) {
  if (empty(value)) return [];
  if (value instanceof Date) return [`${String(value.getUTCDate()).padStart(2,'0')}.${String(value.getUTCMonth()+1).padStart(2,'0')}.${value.getUTCFullYear()}`];
  if (typeof value === 'object') throw new Error('Eine Joker-Datumszelle enthält eine Formel oder einen nicht unterstützten Wert.');
  return String(value).split(/[;,\n]/).map(part => part.trim()).filter(Boolean);
}
function layout(sheet) {
  for (const row of [3,5]) {
    const columns=[];
    sheet.getRow(row).eachCell((cell,column) => {
      const match = /^Joker Datum (\d+)$/.exec(String(cell.value || ''));
      if (match && column >= FIRST) columns.push({ index:Number(match[1]), column });
    });
    if (columns.length) return { row, columns: columns.sort((a,b)=>a.index-b.index) };
  }
  return { row: sheet.getCell('Q3').value ? 3 : 5, columns:[] };
}
function inspect(sheet) {
  const current=layout(sheet);
  if (current.columns.length) {
    const valid=current.columns.every((item,index)=>item.index===index+1);
    if(!valid) throw new Error(`Uneindeutige Joker-Datumsspalten in ${sheet.name}.`);
    return { name:sheet.name, needsMigration:false, columns:current.columns.length, dates:0, students:0 };
  }
  let count=0, students=0, maximum=MIN_COLUMNS;
  for(let row=6;row<=30;row++) {
    const history=dates(sheet.getCell(row,FIRST).value);
    count+=history.length; maximum=Math.max(maximum,history.length);
    if(sheet.getCell(row,1).value) students++;
  }
  // Do not shift or overwrite custom columns to the right of the legacy cell.
  let rightOccupied=false;
  sheet.eachRow(row=>row.eachCell((cell,column)=>{ if(column>FIRST && !empty(cell.value)) rightOccupied=true; }));
  const start=rightOccupied ? Math.max(FIRST+1,sheet.columnCount+1) : FIRST;
  if(start+maximum-1>16384) throw new Error('Kein Platz für weitere Joker-Datumsspalten.');
  return {name:sheet.name,needsMigration:true,dates:count,students,columns:maximum,start,headerRow:current.row,keepsLegacyColumn:rightOccupied};
}
function addColumn(sheet,column,index,headerRow) {
  const source=sheet.getCell(headerRow, FIRST);
  const cell=sheet.getCell(headerRow,column);
  cell.style=JSON.parse(JSON.stringify(source.style));
  cell.value=title(index);
  sheet.getColumn(column).width=18;
  if(headerRow===3 && column!==FIRST && !sheet.getCell(4,column).isMerged) sheet.mergeCells(3,column,4,column);
  for(let row=6;row<=30;row++) {
    sheet.getCell(row,column).style=JSON.parse(JSON.stringify(sheet.getCell(row,FIRST).style));
    sheet.getCell(row,column).numFmt='dd.mm.yyyy';
  }
}
function migrateWorksheet(sheet) {
  const plan=inspect(sheet);
  if(!plan.needsMigration) return plan;
  const histories=[];
  for(let row=6;row<=30;row++) histories.push(dates(sheet.getCell(row,FIRST).value));
  const count=Math.max(plan.columns,...histories.map(history=>history.length));
  for(let index=0;index<count;index++) addColumn(sheet,plan.start+index,index+1,plan.headerRow);
  histories.forEach((history,index)=>history.forEach((date,slot)=>{sheet.getCell(index+6,plan.start+slot).value=date;}));
  if(plan.keepsLegacyColumn) sheet.getCell(plan.headerRow,FIRST).value='Joker Daten (alt)';
  return plan;
}
function append(sheet,row,date) {
  const current=layout(sheet);
  if(!current.columns.length) {
    const cell=sheet.getCell(row,FIRST); cell.value=[...dates(cell.value),date].join('; '); return;
  }
  let slot=current.columns.find(item=>empty(sheet.getCell(row,item.column).value));
  if(!slot) {
    const column=Math.max(sheet.columnCount+1,FIRST+current.columns.length);
    addColumn(sheet,column,current.columns.length+1,current.row);
    slot={column};
  }
  sheet.getCell(row,slot.column).value=date;
}
function remove(sheet,row,date) {
  const current=layout(sheet);
  if(!current.columns.length) {
    const cell=sheet.getCell(row,FIRST), history=dates(cell.value);
    const index=date ? history.lastIndexOf(date) : history.length-1;
    if(index<0) return false;
    history.splice(index,1);cell.value=history.length?history.join('; '):null;return true;
  }
  const slot=[...current.columns].reverse().find(item=>!empty(sheet.getCell(row,item.column).value) && (!date || dates(sheet.getCell(row,item.column).value).includes(date)));
  if(!slot) return false;
  sheet.getCell(row,slot.column).value=null;return true;
}
module.exports={ inspect, migrateWorksheet, append, remove, layout, dates };
