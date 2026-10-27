/** Backend voor Google Apps Script. Koppel dit script aan de financiële spreadsheet. */
const REPO_HTML = 'https://raw.githubusercontent.com/YanniekETP/huishoudboekje/main/index.html';
function doGet() {
 const r=UrlFetchApp.fetch(REPO_HTML,{muteHttpExceptions:true});
 if(r.getResponseCode()!==200) throw new Error('GitHub HTML ophalen mislukt: HTTP '+r.getResponseCode());
 return HtmlService.createHtmlOutput(r.getContentText()).setTitle('Financiën Mandy & Yanniek').addMetaTag('viewport','width=device-width, initial-scale=1');
}
function sheet_(name) {
 const ss=SpreadsheetApp.openById('1VM1BM8eK0nZLpciFIab5DAgR3iHcGgyIXFjL8xWmLdk');
 if(!ss) throw new Error('Koppel dit Apps Script-project aan de juiste Google Spreadsheet.');
 const sh=ss.getSheetByName(name);
 if(!sh) throw new Error('Tabblad ontbreekt: '+name);
 return sh;
}
function amount_(v) {
 if(typeof v==='number') return Number.isFinite(v)?v:NaN;
 let s=String(v??'').trim().replace(/[€\s]/g,'');
 if(!s)return NaN;
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
 const houseRows=sheet_('Vaste lasten koophuis').getRange('F7:G16').getValues();
 const vasteLastenHuis=houseRows.filter(r=>r[0]&&Number.isFinite(amount_(r[1]))&&!/^(totaal samen|per persoon|eigenwoningforfait|ozb)/i.test(String(r[0]))).slice(0,10).map(r=>({item:String(r[0]),bedrag:amount_(r[1])}));
 const vasteLasten=round_(vasteLastenHuis.reduce((s,p)=>s+p.bedrag,0));
 const budgetRows=sheet_('Vaste lasten koophuis').getRange('J7:K16').getValues();
 const gezamenlijkePosten=budgetRows.filter(r=>String(r[0]||'').trim()&&Number.isFinite(amount_(r[1]))).map(r=>({item:String(r[0]).trim(),bedrag:amount_(r[1])}));
 const inbreng=round_(gezamenlijkePosten.reduce((sum,p)=>sum+p.bedrag,0));
 const sh=sheet_('Uitgaven'),totalRow=findInboedelTotalRow_(sh),rows=sh.getRange(1,1,totalRow-1,4).getValues(),uitgaven=[];
 rows.forEach((r,i)=>{
  if(i<5)return;
  const item=String(r[0]||'').trim(),kosten=amount_(r[1]);
  if(!item||!Number.isFinite(kosten)||kosten<=0)return;
  const m=r[2]===''?round_(kosten/2):amount_(r[2]);
  const y=r[3]===''?round_(kosten-m):amount_(r[3]);
  uitgaven.push({row:i+1,item,kosten,aandeelMandy:m,aandeelYanniek:y});
 });
 return {overview:{mandy,yanniek,huis:{vasteLasten,totaalInbreng:inbreng}},vasteLastenHuis,gezamenlijkePosten,uitgaven,totaalInboedelKosten:round_(uitgaven.reduce((s,p)=>s+p.kosten,0)),aandeelMandy:round_(uitgaven.reduce((s,p)=>s+p.aandeelMandy,0)),aandeelYanniek:round_(uitgaven.reduce((s,p)=>s+p.aandeelYanniek,0))};
}
/**
 * Rij met de bestaande totaalformules. Deze rij mag nooit worden overschreven
 * of verwijderd. De positie kan veranderen wanneer extra aankooprijen ontstaan.
 */
function findInboedelTotalRow_(sh) {
 const last=sh.getLastRow();
 const labels=sh.getRange(1,1,last,1).getDisplayValues();
 const found=[];
 labels.forEach((r,i)=>{if(String(r[0]).trim()==='Totaal bekende aankopen')found.push(i+1);});
 if(found.length!==1||found[0]<7)throw new Error('Totaalregel in Uitgaven ontbreekt of is dubbel; wijzigingen veilig gestopt.');
 return found[0];
}
function validateInboedel_(form) {
 if(!form||typeof form!=='object')throw new Error('Ongeldige invoer.');
 const item=String(form.item??'').trim(),kosten=amount_(form.kosten);
 if(!item||item.length>200||!Number.isFinite(kosten)||kosten<=0||kosten>1000000)throw new Error('Vul een geldige omschrijving en positief bedrag in.');
 return {item,kosten:round_(kosten)};
}
function inboedelLock_(){return LockService.getScriptLock();}
function addInboedelUitgave(form) {
 const value=validateInboedel_(form);
 const lock=inboedelLock_();lock.waitLock(10000);
 try {
  const sh=sheet_('Uitgaven');
  let totalRow=findInboedelTotalRow_(sh);
  const cells=sh.getRange(6,1,totalRow-6,2).getValues();
  let row=0;
  for(let i=0;i<cells.length;i++){
   if(cells[i][0]===''&&cells[i][1]===''){row=i+6;break;}
  }
  if(!row){
   sh.insertRowBefore(totalRow);
   row=totalRow;
   // Formules in C:E worden relatief gekopieerd vanuit de vorige rij.
   sh.getRange(totalRow-1,3,1,3).copyTo(sh.getRange(row,3,1,3),SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false);
   totalRow=findInboedelTotalRow_(sh);
   // Bij invoegen direct boven de totaalregel breidt Sheets SUM(B6:B49)
   // niet altijd automatisch uit. Stel de sommen expliciet opnieuw in.
   sh.getRange(totalRow,2,1,3).setFormulas([[
    '=SUM(B6:B'+(totalRow-1)+')',
    '=SUM(C6:C'+(totalRow-1)+')',
    '=SUM(D6:D'+(totalRow-1)+')'
   ]]);
  }
  const target=sh.getRange(row,1,1,2);
  if(target.getValues()[0].some(v=>v!==''))throw new Error('Doelrij is niet leeg; opslaan gestopt.');
  target.setValues([[value.item,value.kosten]]);
  SpreadsheetApp.flush();
 } finally {lock.releaseLock();}
 return getCompleteFinancialData();
}
function deleteInboedelUitgave(row) {
 const n=Number(row);
 if(!Number.isInteger(n))throw new Error('Ongeldige rij.');
 const lock=inboedelLock_();lock.waitLock(10000);
 try {
  const sh=sheet_('Uitgaven'),totalRow=findInboedelTotalRow_(sh);
  if(n<6||n>=totalRow)throw new Error('Deze rij mag niet verwijderd worden.');
  const target=sh.getRange(n,1,1,2);
  const values=target.getValues()[0];
  if(!String(values[0]||'').trim()||!Number.isFinite(amount_(values[1])))throw new Error('Aankoop bestaat niet meer. Vernieuw de pagina.');
  // Wis alleen invoercellen A:B; laat C:E-formules en de totaalregel intact.
  target.clearContent();
  SpreadsheetApp.flush();
 } finally {lock.releaseLock();}
 return getCompleteFinancialData();
}
