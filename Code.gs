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
 const rows=sheet_(name).getRange('A4:B15').getValues(),posten=[];
 rows.forEach((r,i)=>{
  const item=String(r[0]||'').trim(),bedrag=amount_(r[1]);
  if(item&&Number.isFinite(bedrag))posten.push({row:i+4,item,bedrag});
 });
 const income=sheet_(name).getRange('B17').getValue(),inkomen=Number.isFinite(amount_(income))?amount_(income):0;
 const vasteLastenTotaal=round_(posten.reduce((sum,p)=>sum+p.bedrag,0));
 return {inkomen,vasteLastenTotaal,vrij:round_(inkomen-vasteLastenTotaal),posten};
}
function getCompleteFinancialData() {
 const mandy=personal_('Vaste lasten Mandy'),yanniek=personal_('Vaste lasten Yanniek');
 const houseRows=sheet_('Vaste lasten koophuis').getRange('F7:G16').getValues();
 const vasteLastenHuis=houseRows.map((r,i)=>({row:i+7,item:String(r[0]||'').trim(),bedrag:amount_(r[1])})).filter(r=>r.item&&Number.isFinite(r.bedrag));
 const vasteLasten=round_(vasteLastenHuis.reduce((s,p)=>s+p.bedrag,0));
 const budgetRows=sheet_('Vaste lasten koophuis').getRange('J7:K16').getValues();
 const gezamenlijkePosten=budgetRows.map((r,i)=>({row:i+7,item:String(r[0]||'').trim(),bedrag:amount_(r[1])})).filter(r=>r.item&&Number.isFinite(r.bedrag));
 const inbreng=round_(gezamenlijkePosten.reduce((sum,p)=>sum+p.bedrag,0));
 const sh=sheet_('Uitgaven'),totalRow=findInboedelTotalRow_(sh),rows=sh.getRange(1,1,totalRow-1,9).getValues(),uitgaven=[];
 rows.forEach((r,i)=>{
  if(i<5)return;
  const item=String(r[0]||'').trim(),kosten=amount_(r[1]);
  if(!item||!Number.isFinite(kosten)||kosten<=0)return;
  const m=r[2]===''?round_(kosten/2):amount_(r[2]);
  const y=r[3]===''?round_(kosten-m):amount_(r[3]);
  const betaaldDoor=String(r[5]||'').trim()||'onbekend';
  const percentageMandy=r[8]===''?50:amount_(r[8]);
  uitgaven.push({row:i+1,item,kosten,aandeelMandy:m,aandeelYanniek:y,betaaldDoor,percentageMandy:Number.isFinite(percentageMandy)?percentageMandy:50});
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
 const betaaldDoor=String(form.betaaldDoor||'onbekend').trim();
 if(!['enof','mandy','yanniek','onbekend'].includes(betaaldDoor))throw new Error('Ongeldige betaalrekening.');
 const percentageMandy=form.percentageMandy===undefined?50:Number(form.percentageMandy);
 if(!Number.isFinite(percentageMandy)||percentageMandy<0||percentageMandy>100)throw new Error('Verdeling moet tussen 0 en 100% liggen.');
 return {item,kosten:round_(kosten),betaaldDoor,percentageMandy};
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
  writeInboedelVerdeling_(sh,row,value);
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
  sh.getRange(n,6).clearContent();
  sh.getRange(n,9).clearContent();
  restoreInboedelFormulas_(sh,n);
  SpreadsheetApp.flush();
 } finally {lock.releaseLock();}
 return getCompleteFinancialData();
}

/** Editable fixed-cost ranges. Summary rows and formulas are never writable. */
const FIXED_COST_AREAS_={
 mandy:{sheet:'Vaste lasten Mandy',first:4,last:15,nameCol:1,valueCol:2},
 yanniek:{sheet:'Vaste lasten Yanniek',first:4,last:15,nameCol:1,valueCol:2},
 huis:{sheet:'Vaste lasten koophuis',first:7,last:16,nameCol:6,valueCol:7},
 gezamenlijk:{sheet:'Vaste lasten koophuis',first:8,last:16,nameCol:10,valueCol:11}
};
function fixedArea_(area) {
 if(!Object.prototype.hasOwnProperty.call(FIXED_COST_AREAS_,area))throw new Error('Onbekende categorie vaste lasten.');
 return FIXED_COST_AREAS_[area];
}
function fixedValue_(form) {
 if(!form||typeof form!=='object')throw new Error('Ongeldige invoer.');
 const item=String(form.item??'').trim(),raw=form.bedrag,bedrag=amount_(raw);
 if(!item||item.length>150||/^[=+@]/.test(item)||/[\r\n]/.test(item))throw new Error('Vul een geldige omschrijving in.');
 if(raw===''||raw===null||raw===undefined||!Number.isFinite(bedrag)||bedrag<0||bedrag>1000000)throw new Error('Vul een geldig maandbedrag in.');
 return {item,bedrag:round_(bedrag)};
}
function fixedCostMutation_(action,area,row,form,expected) {
 const config=fixedArea_(area);
 const value=action==='delete'?null:fixedValue_(form);
 const n=Number(row),lock=LockService.getScriptLock();
 lock.waitLock(10000);
 try {
  const sh=sheet_(config.sheet);
  const range=sh.getRange(config.first,config.nameCol,config.last-config.first+1,config.valueCol-config.nameCol+1);
  const all=range.getValues();
  let index;
  if(action==='add'){
   index=all.findIndex(r=>r[0]===''&&r[config.valueCol-config.nameCol]==='');
   if(index<0)throw new Error('Geen vrije regels meer in deze categorie. Breid eerst de spreadsheetstructuur uit.');
  }else{
   if(!Number.isInteger(n)||n<config.first||n>config.last)throw new Error('Ongeldige of beschermde rij.');
   index=n-config.first;
   const current=all[index];
   if(!String(current[0]||'').trim()||!Number.isFinite(amount_(current[config.valueCol-config.nameCol])))throw new Error('Post bestaat niet meer; ververs de pagina.');
   if(!expected||String(current[0])!==String(expected.item)||round_(amount_(current[config.valueCol-config.nameCol]))!==round_(amount_(expected.bedrag)))throw new Error('Deze post is ondertussen gewijzigd. Ververs de pagina voordat je opnieuw opslaat.');
  }
  const targetRow=config.first+index;
  const nameCell=sh.getRange(targetRow,config.nameCol),amountCell=sh.getRange(targetRow,config.valueCol);
  if(nameCell.getFormula()||amountCell.getFormula())throw new Error('Deze rij bevat een formule en is beschermd.');
  if(action==='delete'){nameCell.clearContent();amountCell.clearContent();}
  else{nameCell.setValue(value.item);amountCell.setValue(value.bedrag);}
  SpreadsheetApp.flush();
 }finally{lock.releaseLock();}
 return getCompleteFinancialData();
}
function addFixedCost(area,form){return fixedCostMutation_('add',area,null,form,null);}
function updateFixedCost(area,row,form,expected){return fixedCostMutation_('update',area,row,form,expected);}
function deleteFixedCost(area,row,expected){return fixedCostMutation_('delete',area,row,null,expected);}

function restoreInboedelFormulas_(sh,row){
 sh.getRange(row,3).setFormula('=IF(AND(A'+row+'<>"";ISNUMBER(B'+row+'));ROUND(B'+row+'*IF(ISNUMBER(I'+row+');I'+row+';50)/100;2);"")');
 sh.getRange(row,4).setFormula('=IF(ISNUMBER(C'+row+');B'+row+'-C'+row+';"")');
}
function writeInboedelVerdeling_(sh,row,value){
 sh.getRange(row,6).setValue(value.betaaldDoor);
 sh.getRange(row,9).setValue(value.percentageMandy);
 restoreInboedelFormulas_(sh,row);
}
function updateInboedelVerdeling(row,form,expected){
 const n=Number(row);
 if(!Number.isInteger(n))throw new Error('Ongeldige aankooprij.');
 const value=validateInboedel_(form),lock=inboedelLock_();
 lock.waitLock(10000);
 try{
  const sh=sheet_('Uitgaven'),totalRow=findInboedelTotalRow_(sh);
  if(n<6||n>=totalRow)throw new Error('Deze aankooprij is beschermd.');
  const old=sh.getRange(n,1,1,2).getValues()[0];
  if(!expected||String(old[0]).trim()!==String(expected.item).trim()||round_(amount_(old[1]))!==round_(amount_(expected.kosten)))throw new Error('Aankoop is gewijzigd; vernieuw eerst de pagina.');
  sh.getRange(n,1,1,2).setValues([[value.item,value.kosten]]);
  writeInboedelVerdeling_(sh,n,value);
  SpreadsheetApp.flush();
 }finally{lock.releaseLock();}
 return getCompleteFinancialData();
}
