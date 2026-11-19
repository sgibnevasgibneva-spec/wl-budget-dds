const fs=require('fs'),vm=require('vm'),assert=require('assert');
const data=new Map();let admin=true;
function sheet(name){if(!data.has(name))data.set(name,[]);const rows=data.get(name);return {
 getLastRow:()=>rows.length+1,
 getRange:(r,c,n=1,w=1)=>({getValue:()=>rows[r-2]?.[c-1],getValues:()=>rows.slice(r-2,r-2+n).map(row=>row.slice(c-1,c-1+w)),setValues:values=>values.forEach((row,i)=>{rows[r-2+i]??=[];row.forEach((v,j)=>rows[r-2+i][c-1+j]=v);})}),
 appendRow:row=>rows.push(row)
};}
const ss={getSpreadsheetTimeZone:()=> 'Europe/Moscow',getSheetByName:n=>data.has(n)?sheet(n):null};
const ctx={Session:{getScriptTimeZone:()=> 'Europe/Moscow'}};vm.createContext(ctx);
vm.runInContext(fs.readFileSync('Code.js','utf8')+'\n'+fs.readFileSync('BddsHints.js','utf8'),ctx);
Object.assign(ctx,{wlStore_:()=>ss,wlNewSheet_:(_ss,n)=>sheet(n),wlLock_:fn=>fn(),wlKey_:x=>x,wlMonthCellKey_:x=>x,wlRequireTab_:()=>{},wlAdmin_:()=>{if(!admin)throw Error('admin only');return {name:'Admin',email:'admin@example.test'};}});
const html=fs.readFileSync('Index.html','utf8');
vm.runInContext(html.slice(html.indexOf('function renderBddsHistory(){'),html.indexOf('function bddsCombineTables(){')),ctx);
const key='history|поступления за it-услуги и сервисы|1',id='2026-10|'+key,text='Страховка за пострадавший товар на озоне';
const els={bddsHistoryBody:{},bddsHistoryHead:{}};
Object.assign(ctx,{$:x=>els[x],boot:{user:{admin:true},grouping:{options:{}}},bddsHistoryData:{months:['2026-09','2026-10'],rows:[{label:'Поступления за IT-услуги и сервисы',cells:[['150','','','',text],['50','','','',text]]}]},bddsYear:'2026',bddsExpandedMonths:new Set(['2026-10']),bddsFactData:null,bddsData:{comments:{}},money:String,esc:String,monthsRu:Array(12).fill('month')});
function reload(){ctx.bddsData.comments=ctx.wlBddsComments_(ss,'2026');ctx.renderBddsHistory();return ctx.bddsCommentsForCell('2026-10',key);}
assert.equal(reload()[0].text,text);assert(els.bddsHistoryBody.innerHTML.includes(text));
ctx.wlBddsSetHint_('2026','2026-10',key,'');assert.equal(reload()[0].text,text);
// Actual regression: the old delete left an empty stored thread, but the source reappeared.
data.set('WL_БДДС_Комментарии',[['2026','2026-10',key,'[]']]);
assert.equal(reload().length,0);assert(!els.bddsHistoryBody.innerHTML.includes(text));
assert.equal(ctx.bddsCommentsForCell('2026-09',key)[0].text,text);
// Adding a new comment or dismissing a hint must not resurrect the removed source.
ctx.wlBddsPostComment_('2026','2026-10',key,'Новый комментарий');
assert.deepEqual(Array.from(reload(),x=>x.text),['Новый комментарий']);
ctx.wlBddsSetHint_('2026','2026-10',key,'');
assert.deepEqual(Array.from(reload(),x=>x.text),['Новый комментарий']);
const normal=reload()[0];ctx.wlBddsDeleteComment_('2026','2026-10',key,normal);
assert.equal(reload().length,0);
// Source comments can be edited, deleted and reread, without touching the source workbook.
ctx.wlBddsSetSourceComment_('2026','2026-09',key,null,'Исправлено');
assert.equal(ctx.wlBddsComments_(ss,'2026')['2026-09|'+key][0].text,'Исправлено');
assert.throws(()=>ctx.wlBddsSetSourceComment_('2026','2026-09',key,null,'Старое окно'));
const expected=ctx.wlBddsComments_(ss,'2026')['2026-09|'+key][0];
ctx.wlBddsSetSourceComment_('2026','2026-09',key,expected,'');reload();
assert.equal(ctx.bddsCommentsForCell('2026-09',key).length,0);
assert.equal(ctx.bddsHistoryData.rows[0].cells[0][4],text);
// Deleting an unrelated user comment preserves the existing source comment.
data.set('WL_БДДС_Комментарии',[]);
const added=ctx.wlBddsPostComment_('2026','2026-10',key,'Другой текст').messages[0];
ctx.wlBddsDeleteComment_('2026','2026-10',key,added);assert.equal(reload()[0].text,text);
// Long histories retain deletion markers and hint dismissals.
const inherited=ctx.wlBddsComments_(ss,'2026')[id].find(x=>x.kind==='history-source');
ctx.wlBddsSetSourceComment_('2026','2026-10',key,inherited,'');ctx.wlBddsSetHint_('2026','2026-10',key,'');
for(let i=0;i<35;i++)ctx.wlBddsPostComment_('2026','2026-10',key,'Комментарий '+i);
assert.equal(reload().length,30);assert(!els.bddsHistoryBody.innerHTML.includes(text));
assert.equal(ctx.bddsData.comments[id].filter(x=>x.kind==='history-source').length,1);
assert.equal(ctx.bddsData.comments[id].filter(x=>x.kind==='hint').length,1);
admin=false;assert.throws(()=>ctx.wlBddsSetSourceComment_('2026','2026-10',key,null,''));
console.log('Historical comment deletion survives refresh/new notes/limits; edit, conflict, month isolation and admin checks passed');
