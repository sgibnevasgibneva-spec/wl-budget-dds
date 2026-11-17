const fs=require('fs'),vm=require('vm'),assert=require('assert');
const src=fs.readFileSync('Code.js','utf8');const fn=src.slice(src.indexOf('function wlBddsFetchFact('),src.indexOf('\n}',src.indexOf('function wlBddsFetchFact('))+2);
const ctx={wlKey_:s=>s,Utilities:{formatDate:()=> '2026-11'},Session:{getScriptTimeZone:()=> 'Europe/Moscow'}};vm.createContext(ctx);vm.runInContext(fn,ctx);
assert.throws(()=>ctx.wlBddsFetchFact('2026','2026-11'),/ещё не завершился/);
assert.throws(()=>ctx.wlBddsFetchFact('2026','2027-01'),/месяц выбранного года/);
assert(!fn.includes('/transaction'));assert(fn.includes("report(requestedMonth,'category')"));
console.log('Unfinished months rejected; source is the read-only report, not transaction plans');
