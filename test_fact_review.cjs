const fs=require('fs'),vm=require('vm'),assert=require('assert');
const code=fs.readFileSync('Code.js','utf8');
const fn=code.slice(code.indexOf('function wlBddsFetchFact('),code.indexOf('\n}',code.indexOf('function wlBddsFetchFact('))+2);
const rows=[
 {id:1,date:'10.10.2026',group:'income',moneybagId:1,value:100,categoryId:11},
 {id:2,date:'11.10.2026',group:'outcome',moneybagId:1,value:25,comment:'Проверить'},
 {id:3,date:'12.10.2026',group:'outcome',moneybagId:1,value:40,categoryId:99},
 {id:4,date:'12.10.2026',group:'transfer',moneybagId:1,value:50},
 {id:5,date:'12.11.2026',group:'income',moneybagId:1,value:500,categoryId:11}
];
const ctx={Utilities:{formatDate:()=> '2026-11'},Session:{getScriptTimeZone:()=> 'Europe/Moscow'},wlKey_:x=>x,wlStore_:()=>({}),wlAdmin_:()=>({}),wlRequireTab_:()=>{},PropertiesService:{getScriptProperties:()=>({getProperty:()=> 'token'})},wlFinTabloFetch_:()=>[{id:11,name:'Продажи'}],wlBddsGrouping_:()=>({}),wlBddsNormalize_:x=>x,wlBddsMeta_:()=>({}),UrlFetchApp:{fetch:url=>({getResponseCode:()=>200,getContentText:()=>JSON.stringify({items:url.includes('/moneybag')?[{id:1,currency:'RUB'}]:rows})})}};
vm.createContext(ctx);vm.runInContext(fn,ctx);const result=ctx.wlBddsFetchFact('2026','2026-10');
assert.equal(result.articles[0].byMonth['2026-10'],100);assert.equal(result.unclassified.length,2);assert.equal(result.unclassified[0].amount,25);assert.equal(result.unclassified[1].categoryId,'99');assert.equal(result.warnings.ignoredTransfer,1);assert.throws(()=>ctx.wlBddsFetchFact('2026','2026-11'),/не завершился/);
console.log('Fact month, read-only unclassified diagnostics, transfers and unfinished month passed');
