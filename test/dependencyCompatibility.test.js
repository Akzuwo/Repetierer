const {test}=require('node:test');
const assert=require('node:assert/strict');
const Excel=require('exceljs');
test('updated UUID dependency preserves extended Excel conditional formatting',async()=>{
 const workbook=new Excel.Workbook();
 const sheet=workbook.addWorksheet('Fixture');
 sheet.getCell('A1').value=3;
 sheet.addConditionalFormatting({ref:'A1:A3',rules:[{type:'iconSet',iconSet:'3Stars',cfvo:[{type:'percent',value:0},{type:'percent',value:33},{type:'percent',value:67}]}]});
 const copy=new Excel.Workbook();
 await copy.xlsx.load(await workbook.xlsx.writeBuffer());
 assert.equal(copy.worksheets[0].getCell('A1').value,3);
 const rule=copy.worksheets[0].conditionalFormattings[0].rules[0];
 assert.equal(rule.iconSet,'3Stars');
 assert.match(rule.x14Id,/^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/i);
});
