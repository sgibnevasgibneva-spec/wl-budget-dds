const fs=require('fs'),vm=require('vm'),assert=require('assert');
const html=fs.readFileSync('Index.html','utf8');
async function scenario(fail){
 let resolve,reject,requests=0,remembered=0,stored=0;const request=new Promise((a,b)=>{resolve=a;reject=b;}),elements={};
 const element=id=>elements[id]||(elements[id]={inert:false,innerHTML:'',textContent:'',classList:{add(){},remove(){},toggle(){}},appendChild(){},style:{setProperty(){}}});
 const data={boot:{current:'2026-11',months:[{key:'2026-11',status:'Открыт'}],dict:{},user:{admin:false,view:['HR'],tabs:['Отделы'],tabEdit:['Отделы'],edit:['HR']}},month:{key:'2026-11',month:{status:'Открыт'},rows:[{department:'HR'}]},timing:{serverMs:100}};
 const c={Date,performance,console:{info(){}},setTimeout(){},window:{addEventListener(){}},document:{querySelector:element,querySelectorAll:()=>[],documentElement:{style:{setProperty(){}}},createElement:element},$:element,VIEWS:['deptView'],TABS:['Отделы'],activeTab:'Отделы',dept:'',dict:{},authGet:()=> 'session',call:name=>{assert.equal(name,'wlStartup');requests++;return request;},status(){},hideError(){},showError(){},renderYearOptions(){},selectMonths(){},tabs(){},renderRows(){},rememberMonth(){remembered++},views(){},esc:String,localStorage:{removeItem(){}}};
 vm.createContext(c);vm.runInContext(html.slice(html.indexOf('let startupRun='),html.indexOf('function refreshBulkUi()')),c);
 c.startupCacheRead=async()=>data;c.startupCacheWrite=()=>stored++;
 c.startApp();await Promise.resolve();await Promise.resolve();assert.equal(requests,1);assert.equal(c.loading,true);assert.equal(elements.deptView.inert,true);assert.equal(remembered,0);
 if(fail)reject(Error('network'));else resolve(data);await Promise.resolve();await Promise.resolve();
 assert.equal(c.loading,fail);assert.equal(elements.deptView.inert,fail);assert.equal(remembered,fail?0:1);assert.equal(stored,fail?0:1);
}
(async()=>{await scenario(false);await scenario(true);console.log('One initial RPC; cached preview is inert; only fresh response enables editing; network errors keep editing blocked');})().catch(e=>{console.error(e);process.exitCode=1});
