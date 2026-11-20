const {pdfHtml}=require('./studentDossierView');
async function renderDossierPdf(BrowserWindow,record){
  const win=new BrowserWindow({show:false,webPreferences:{nodeIntegration:false,contextIsolation:true,sandbox:true}});
  try{
    await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(pdfHtml(record))}`);
    return await win.webContents.printToPDF({printBackground:true,pageSize:'A4',preferCSSPageSize:true,displayHeaderFooter:true,headerTemplate:'<span></span>',footerTemplate:'<div style="font:8px Arial;width:100%;text-align:center"><span class="pageNumber"></span> / <span class="totalPages"></span></div>'});
  }finally{if(!win.isDestroyed())win.destroy();}
}
module.exports={renderDossierPdf};
