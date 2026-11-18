const fs=require('fs'),vm=require('vm'),assert=require('assert');const s=fs.readFileSync('BddsFactStore.js','utf8');const ctx={wlMonthCellKey_:v=>v};vm.createContext(ctx);vm.runInContext(s,ctx);
const report=(month,value)=>({year:'2026',month,articles:[{id:'income|a',byMonth:{[month]:value}}]});
const old=JSON.stringify(report('2026-10',10)),fresh=JSON.stringify(report('2026-10',20)),sep=Math.floor(fresh.length/2);
const rows=[['2026-10','a',1,1,old,''],['2026-10','b',1,2,fresh.slice(0,sep),''],['2026-10','b',2,2,fresh.slice(sep),''],['2026-10','incomplete',1,2,'{',''],['2025-10','old',1,1,old,'']];
const out=ctx.wlBddsReadFactSnapshots_(rows,'2026','Europe/Moscow');assert.equal(out.length,1);assert.equal(out[0].articles[0].byMonth['2026-10'],20);
console.log('Restores latest complete monthly snapshot; ignores partial snapshots and other years');
