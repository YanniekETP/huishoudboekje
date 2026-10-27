/** Backend voor Google Apps Script. Koppel dit script aan de financiële spreadsheet. */
const REPO_HTML = 'https://raw.githubusercontent.com/YanniekETP/huishoudboekje/main/index.html';
function doGet() {
 const r=UrlFetchApp.fetch(REPO_HTML,{muteHttpExceptions:true});
 if(r.getResponseCode()!==200) throw new Error('GitHub HTML ophalen mislukt: HTTP '+r.getResponseCode());
 return HtmlService.createHtmlOutput(r.getContentText()).setTitle('Financiën Mandy & Yanniek').addMetaTag('viewport','width=device-width, initial-scale=1');
}
function sheet_(name) {
 const ss=SpreadsheetApp.getActiveSpreadsheet();
 if(!ss) throw new Error('Koppel dit Apps Script-project aan de juiste Google Spreadsheet.');
 const sh=ss.getSheetByName(name);
 if(!sh) throw new Error('Tabblad ontbreekt: '+name);
 return sh;
}
function amount_(v) {
 if(typeof v==='number') return Number.isFinite(v)?v:NaN;
 let s=String(v??'').trim().replace(/[€\s]/g,'');
 if(s.includes(',') && s.includes('.')) s=s.replace(/\./g,'').replace(',','.');
 else if(s.includes(',')) s=s.replace(',','.');
 return Number(s);
}
function round_(v){return Math.round((v+Number.EPSILON)*100)/100;}
function readBudget_(name) {
 const rows=sheet_(name).getDataRange().getValues();
 const posten=[];let inkomen=0;
 rows.forEach((r,i)=>{
  const item=String(r[0]??'').trim(),bedrag=amount_(r[1]);
  if(!item||!Number.isFinite(bedrag)) return;
  if(/^(inkomen|salaris|loon|netto inkomen)$/i.test(item)) inkomen+=bedrag;
  else posten.push({item,bedrag});
 });
 const vasteLastenTotaal=round_(posten.reduce((s,p)=>s+p.bedrag,0));
 return {inkomen:round_(inkomen),vasteLastenTotaal,vrij:round_(inkomen-vasteLastenTotaal),posten};
}
function getCompleteFinancialData() {
 const mandy=readBudget_('Mandy'),yanniek=readBudget_('Yanniek');
 const huis=readBudget_('Gezamenlijk');
 const sh=sheet_('Uitgaven'),rows=sh.getDataRange().getValues();
 const uitgaven=[];
 rows.slice(1).forEach((r,i)=>{
  const item=String(r[0]??'').trim(),kosten=amount_(r[1]);
  if(!item||!Number.isFinite(kosten))return;
  const m=r[2]===''?round_(kosten/2):amount_(r[2]);
  const y=r[3]===''?round_(kosten-m):amount_(r[3]);
  uitgaven.push({row:i+2,item,kosten,aandeelMandy:m,aandeelYanniek:y});
 });
 const vasteLastenHuis=huis.posten.filter(p=>!/^(sparen|boodschappen|eten|vakantie)$/i.test(p.item));
 return {overview:{mandy,yanniek,huis:{vasteLasten:round_(vasteLastenHuis.reduce((s,p)=>s+p.bedrag,0)),totaalInbreng:round_(huis.vasteLastenTotaal)}},vasteLastenHuis,gezamenlijkePosten:huis.posten,uitgaven,totaalInboedelKosten:round_(uitgaven.reduce((s,p)=>s+p.kosten,0)),aandeelMandy:round_(uitgaven.reduce((s,p)=>s+p.aandeelMandy,0)),aandeelYanniek:round_(uitgaven.reduce((s,p)=>s+p.aandeelYanniek,0))};
}
function addInboedelUitgave(form) {
 const item=String(form.item??'').trim(),kosten=amount_(form.kosten);
 if(!item||item.length>200||!Number.isFinite(kosten)||kosten<=0||kosten>1000000)throw new Error('Vul een geldige omschrijving en positief bedrag in.');
 const lock=LockService.getDocumentLock()||LockService.getScriptLock();
 lock.waitLock(10000);
 try {
  const m=round_(kosten/2),y=round_(kosten-m);
  sheet_('Uitgaven').appendRow([item,round_(kosten),m,y]);
  SpreadsheetApp.flush();
 } finally {lock.releaseLock();}
 return getCompleteFinancialData();
}
function deleteInboedelUitgave(row) {
 const n=Number(row);if(!Number.isInteger(n)||n<2)throw new Error('Ongeldige rij.');
 const lock=LockService.getDocumentLock()||LockService.getScriptLock();lock.waitLock(10000);
 try {const sh=sheet_('Uitgaven');if(n>sh.getLastRow())throw new Error('Rij bestaat niet.');sh.deleteRow(n);SpreadsheetApp.flush();}
 finally{lock.releaseLock();}
 return getCompleteFinancialData();
}