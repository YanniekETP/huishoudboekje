// Run locally: node tests/backend.test.cjs
const assert=require('node:assert/strict');
const vm=require('node:vm');
const fs=require('node:fs');
const src=fs.readFileSync('Code.gs','utf8');
function makeSheet(){
 const cells=Array.from({length:50},()=>Array(5).fill(''));
 cells[4][0]='Item';cells[5]=['Bed',100,50,50,''];
 cells[49]=['Totaal bekende aankopen','=SUM(B6:B49)','=SUM(C6:C49)','=SUM(D6:D49)',''];
 return {
  cells,
  getLastRow(){return this.cells.length;},
  getDisplayValues(){return this.cells.map(r=>[r[0]]);},
  getRange(row,col,numRows=1,numCols=1){
   const sh=this;
   return {
    getDisplayValues(){return sh.cells.slice(row-1,row-1+numRows).map(r=>r.slice(col-1,col-1+numCols));},
    getValues(){return sh.cells.slice(row-1,row-1+numRows).map(r=>r.slice(col-1,col-1+numCols));},
    setValues(v){v.forEach((r,i)=>r.forEach((x,j)=>sh.cells[row-1+i][col-1+j]=x));},
    setFormulas(v){this.setValues(v);},
    setFormula(v){this.setValues([[v]]);},
    setValue(v){this.setValues([[v]]);},
    clearContent(){for(let i=0;i<numRows;i++)for(let j=0;j<numCols;j++)sh.cells[row-1+i][col-1+j]='';},
    copyTo(dest){const vals=this.getValues();dest.setValues(vals);},
   };
  },
  insertRowBefore(n){this.cells.splice(n-1,0,Array(5).fill(''));}
 };
}
const sh=makeSheet(),lock={waitLock(){},releaseLock(){}};
const ctx={SpreadsheetApp:{openById(){return {getSheetByName(){return sh}}},flush(){},CopyPasteType:{PASTE_FORMULA:'PASTE_FORMULA'}},LockService:{getScriptLock(){return lock}}};
vm.createContext(ctx);vm.runInContext(src,ctx);
// Isolate mutation tests from the unrelated financial dashboard reads.
ctx.getCompleteFinancialData=()=>({});
assert.equal(ctx.findInboedelTotalRow_(sh),50);
assert.equal(ctx.validateInboedel_({item:'Lamp',kosten:'€ 25,50'}).kosten,25.5);
for(const bad of [{item:'',kosten:10},{item:'Lamp',kosten:0},{item:'Lamp',kosten:'abc'}])assert.throws(()=>ctx.validateInboedel_(bad));
ctx.addInboedelUitgave({item:'Lamp',kosten:25.5});
assert.deepEqual(sh.cells[6].slice(0,2),['Lamp',25.5]);
assert.equal(ctx.findInboedelTotalRow_(sh),50);
ctx.deleteInboedelUitgave(7);
assert.deepEqual(sh.cells[6].slice(0,2),['','']);
assert.equal(sh.cells[49][0],'Totaal bekende aankopen');
assert.throws(()=>ctx.deleteInboedelUitgave(50));
assert.throws(()=>ctx.deleteInboedelUitgave(5));
assert.equal(sh.cells[5][0],'Bed');
// Exhaust all available purchase rows and ensure a new row is inserted safely.
for(let row=7;row<=49;row++){sh.cells[row-1][0]='Test '+row;sh.cells[row-1][1]=row;}
ctx.addInboedelUitgave({item:'Overflow',kosten:12.34});
assert.equal(ctx.findInboedelTotalRow_(sh),51);
assert.equal(sh.cells[49][0],'Overflow');
assert.equal(sh.cells[49][1],12.34);
assert.equal(sh.cells[50][0],'Totaal bekende aankopen');
assert.equal(sh.cells[50][1],'=SUM(B6:B50)');
assert.equal(sh.cells[50][2],'=SUM(C6:C50)');
assert.equal(sh.cells[50][3],'=SUM(D6:D50)');
ctx.deleteInboedelUitgave(50);
assert.equal(sh.cells[49][0],'');
assert.equal(sh.cells[50][0],'Totaal bekende aankopen');
assert.throws(()=>ctx.deleteInboedelUitgave(51));
assert.equal(sh.cells[49][5],'');
assert.equal(sh.cells[49][8],'');
assert.throws(()=>ctx.validateInboedel_({item:'Lamp',kosten:10,betaaldDoor:'enof',percentageMandy:120}));
assert.equal(ctx.validateInboedel_({item:'Lamp',kosten:10,betaaldDoor:'mandy',percentageMandy:70}).percentageMandy,70);
console.log('Backend tests passed: validation, add, delete, overflow, formula preservation, totals');
