const fs=require('node:fs');
const vm=require('node:vm');
const assert=require('node:assert/strict');
const html=fs.readFileSync('index.html','utf8');
const scripts=[...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)];
assert.ok(scripts.length>0,'Geen inline JavaScript gevonden');
for(const [,source] of scripts)new vm.Script(source,{filename:'index.html'});
for(const id of ['loading','postenMandy','postenYanniek','totaalInbreng','inbrengPP','startupError']){
 assert.ok(html.includes('id="'+id+'"'),'HTML element ontbreekt: '+id);
}
console.log('Frontend JavaScript syntaxis en kritieke elementen OK');
