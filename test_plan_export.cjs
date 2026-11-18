const fs=require('fs'),vm=require('vm'),assert=require('assert');
const props=new Map([['WL_DDS_FINTABLO_API_TOKEN','test-token']]);
const records=[{id:9,categoryId:10,date:'12.2026',value:50,directionId:1},{id:10,categoryId:10,date:'12.2026',value:50,directionId:2}];
let posts=0;
const c={Utilities:{DigestAlgorithm:{SHA_256:1},Charset:{UTF_8:1},computeDigest:(_,s)=>[...Buffer.from(s)].slice(0,32),base64EncodeWebSafe:x=>Buffer.from(x).toString('base64url')},PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k)||null,setProperty:(k,v)=>props.set(k,v),deleteProperty:k=>props.delete(k)})},
 wlKey_:x=>x,wlStore_:()=>({}),wlAdmin_:()=>({email:'admin@example.test'}),wlRequireTab_:()=>{},wlFindMonth_:()=>({}),wlLock_:f=>f(),wlDictionaries_:()=>({incomeCategories:[],categories:['Аренда','Связь']}),wlBddsNormalize_:x=>String(x).toLowerCase(),
 wlOps_:()=>({}),wlOpsRead_:()=>[{row:{category:'Аренда',department:'Отдел',cashFlowType:'expense',amounts:{1:100}}},{row:{category:'Связь',department:'Отдел',cashFlowType:'expense',amounts:{1:25}}}],
 UrlFetchApp:{fetch:(url,opts)=>{let data;if(opts.method==='get'){const id=new URL(url).searchParams.get('categoryId');data={items:id?records.filter(x=>x.categoryId===Number(id)):records.slice()};}else{posts++;const item=JSON.parse(opts.payload);item.id=99;records.push(item);data={items:[item]};}return {getResponseCode:()=>200,getContentText:()=>JSON.stringify(data)};}},
 URL,console};
vm.createContext(c);vm.runInContext(fs.readFileSync('FinTabloSync.js','utf8'),c);
c.wlFinTabloFetch_=()=>[];
c.wlFinTabloNormalize_=()=>({candidates:[{id:'10',name:'Аренда',group:'outcome'},{id:'11',name:'Связь',group:'outcome'}],warnings:['Совпадающие названия разных статей. Требуется отдельное сопоставление: курсовая разница']});
let p=c.wlFinTabloBudgetPreview_('2026-12');assert.equal(p.warnings.length,0);assert.equal(p.rows.find(x=>x.name==='Аренда').action,'skip');assert.equal(p.rows.find(x=>x.name==='Связь').action,'create');
let sent=c.wlFinTabloBudgetSend_('2026-12',p.hash);assert.equal(sent.created,1);assert.equal(sent.skipped,1);assert.equal(posts,1);
p=c.wlFinTabloBudgetPreview_('2026-12');assert(p.rows.every(x=>x.action==='skip'));sent=c.wlFinTabloBudgetSend_('2026-12',p.hash);assert.equal(sent.created,0);assert.equal(posts,1);
records[0].value=51;p=c.wlFinTabloBudgetPreview_('2026-12');assert.equal(p.rows.find(x=>x.name==='Аренда').action,'blocked');assert.throws(()=>c.wlFinTabloBudgetSend_('2026-12',p.hash),/требующие проверки/);
console.log('Plan export skips equal existing totals, creates missing once and blocks conflicting totals');
