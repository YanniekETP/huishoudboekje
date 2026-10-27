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
function personal_(name) {
 const rows=sheet_(name).getRange('A1:B40').getValues(),posten=[];let inkomen=0;
 rows.forEach(r=>{const item=String(r[0]||'').trim(),bedrag=amount_(r[1]);if(!item||!Number.isFinite(bedrag))return;
 if(/^netto inkomen$/i.test(item)) inkomen=bedrag;
 else if(!/^(ingevulde maandlasten|vrij te besteden|voorlopig over|nog in te vullen)$/i.test(item))posten.push({item,bedrag});
 });
 const vasteLastenTotaal=round_(posten.reduce((s,p)=>s+p.bedrag,0));
 return {inkomen,vasteLastenTotaal,vrij:round_(inkomen-vasteLastenTotaal),posten};
}
function getCompleteFinancialData() {
 const mandy=personal_('Vaste lasten Mandy'),yanniek=personal_('Vaste lasten Yanniek');
 const houseRows=sheet_('Vaste lasten koophuis').getRange('F7:G30').getValues();
 const vasteLastenHuis=houseRows.filter(r=>r[0]&&Number.isFinite(amount_(r[1]))&&!/^(totaal samen|per persoon|eigenwoningforfait|ozb)/i.test(String(r[0]))).slice(0,10).map(r=>({item:String(r[0]),bedrag:amount_(r[1])}));
 const vasteLasten=round_(vasteLastenHuis.reduce((s,p)=>s+p.bedrag,0));
 const gezamenlijkePosten=[{item:'Vaste lasten huis',bedrag:vasteLasten},{item:'Sparen',bedrag:600},{item:'Eten & Boodschappen',bedrag:640},{item:'Vakantie',bedrag:200}];
 const sh=sheet_('Uitgaven'),rows=sh.getDataRange().getValues(),uitgaven=[];
 rows.forEach((r,i)=>{
  if(i<5)return;
  const item=String(r[0]||'').trim(),kosten=amount_(r[1]);
  if(!item||!Number.isFinite(kosten)||kosten<=0)return;
  const m=r[2]===''?round_(kosten/2):amount_(r[2]);
  const y=r[3]===''?round_(kosten-m):amount_(r[3]);
  uitgaven.push({row:i+1,item,kosten,aandeelMandy:m,aandeelYanniek:y});
 });
 return {overview:{mandy,yanniek,huis:{vasteLasten,totaalInbreng:round_(gezamenlijkePosten.reduce((s,p)=>s+p.bedrag,0))}},vasteLastenHuis,gezamenlijkePosten,uitgaven,totaalInboedelKosten:round_(uitgaven.reduce((s,p)=>s+p.kosten,0)),aandeelMandy:round_(uitgaven.reduce((s,p)=>s+p.aandeelMandy,0)),aandeelYanniek:round_(uitgaven.reduce((s,p)=>s+p.aandeelYanniek,0))};
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
 const n=Number(row);if(!Number.isInteger(n)||n<6)throw new Error('Ongeldige rij.');
 const lock=LockService.getDocumentLock()||LockService.getScriptLock();lock.waitLock(10000);
 try {const sh=sheet_('Uitgaven');if(n>sh.getLastRow())throw new Error('Rij bestaat niet.');sh.deleteRow(n);SpreadsheetApp.flush();}
 finally{lock.releaseLock();}
 return getCompleteFinancialData();
}