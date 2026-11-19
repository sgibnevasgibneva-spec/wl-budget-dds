const fs=require('fs'),vm=require('vm'),assert=require('assert'),crypto=require('crypto');
const propsData=new Map(),cached=new Map(),rules=[{email:'admin@test.local',role:'admin',enabled:true,view:[]},{email:'a@test.local',name:'Alice',role:'user',enabled:true,view:['HR']},{email:'b@test.local',name:'Bob',role:'user',enabled:true,view:['IT']}];let admin=true,reads=0;
const props={getProperties:()=>{reads++;return Object.fromEntries(propsData);},getProperty:k=>propsData.get(k)||null,setProperty:(k,v)=>propsData.set(k,v),deleteProperty:k=>propsData.delete(k)},cache={get:k=>cached.get(k)||null,put:(k,v)=>cached.set(k,v),remove:k=>cached.delete(k)};
const c={Utilities:{DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},getUuid:()=>crypto.randomUUID(),computeDigest:(_a,s)=>[...crypto.createHash('sha256').update(String(s)).digest()],base64EncodeWebSafe:a=>Buffer.from(a).toString('base64url')},PropertiesService:{getScriptProperties:()=>props},CacheService:{getScriptCache:()=>cache},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},wlStore_:()=>({getUrl:()=>''}),wlRules_:()=>rules,wlLock_:fn=>fn(),wlAdmin_:()=>{if(!admin)throw Error('Admin only');},wlDictionaries_:()=>({departments:[]}),WL_TAB_NAMES:[]};vm.createContext(c);
vm.runInContext(fs.readFileSync('LoginSecurity.js','utf8')+'\n'+fs.readFileSync('LoginCodes.js','utf8'),c);const source=fs.readFileSync('Code.js','utf8');vm.runInContext(source.slice(source.indexOf('function wlGetSettings_(){'),source.indexOf('function wlSaveAccess_(')),c);
const issued=c.wlIssueEmployeeCodes_(['a@test.local','b@test.local'],crypto.randomUUID(),{}).items,original=issued.map(x=>propsData.get(c.wlAuthCredentialKey_(x.email)));
const session=c.wlVerifyLogin(issued[0].email,issued[0].code).token;
// Previous release only cached the displayed code. Preserve the same value before expiry.
issued.forEach(x=>propsData.delete(c.wlEmployeeCodeDisplayKey_(x.email)));const migrated=c.wlGetSettings_();assert.equal(migrated.rules[1].loginCode,issued[0].code);
cached.clear();assert.equal(c.wlGetSettings_().rules[1].loginCode,issued[0].code);assert.equal(c.wlAuthSession_(session),issued[0].email);
assert.deepEqual(issued.map(x=>propsData.get(c.wlAuthCredentialKey_(x.email))),original);
assert.equal(c.wlGetSettings_().rules[0].loginCode,'');rules[2].enabled=false;assert.equal(c.wlGetSettings_().rules[2].loginCode,'');rules[2].enabled=true;
admin=false;const before=reads;assert.throws(()=>c.wlGetSettings_());assert.equal(reads,before);assert.throws(()=>c.wlRememberEmployeeCodes_(issued));admin=true;
// Restore a saved list after the temporary cache expired; validate all items before writing.
issued.forEach(x=>propsData.delete(c.wlEmployeeCodeDisplayKey_(x.email)));
assert.throws(()=>c.wlRememberEmployeeCodes_([{email:issued[0].email,code:issued[0].code},{email:issued[1].email,code:issued[0].code}]));assert(!propsData.has(c.wlEmployeeCodeDisplayKey_(issued[0].email)));
c.wlRememberEmployeeCodes_(issued);assert.equal(c.wlGetSettings_().rules[2].loginCode,issued[1].code);assert.equal(c.wlAuthSession_(session),issued[0].email);assert.deepEqual(issued.map(x=>propsData.get(c.wlAuthCredentialKey_(x.email))),original);
c.wlGenerateLoginCode_(issued[1].email);assert.equal(c.wlGetSettings_().rules[2].loginCode,'');assert(!propsData.has(c.wlEmployeeCodeDisplayKey_(issued[1].email)));
c.wlAuthEraseCode_(issued[0].email);assert(!propsData.has(c.wlEmployeeCodeDisplayKey_(issued[0].email)));
console.log('Display codes survive reload/cache expiry, are admin-only, preserve credentials/sessions, and clear on reset/revocation');

(async()=>{const html=fs.readFileSync('Index.html','utf8'),part=html.slice(html.indexOf('function employeeLoginCodeCell('),html.indexOf('function rememberEmployeeCodes('));let copied='';const node=tag=>({tag,children:[],appendChild(n){this.children.push(n);},setAttribute(k,v){this[k]=v;},select(){this.selected=true;}}),status={};const u={document:{createElement:node},navigator:{clipboard:{writeText:async text=>{copied=text;}}},$:()=>status};vm.createContext(u);vm.runInContext(part,u);
const td=u.employeeLoginCodeCell({email:'a@test.local',name:'Alice',role:'user',enabled:true,loginCode:issued[0].code}),input=td.children[0];assert.equal(input.readOnly,true);assert.equal(input.value,issued[0].code);await input.onclick();assert.equal(copied,issued[0].code);assert(input.selected);
assert.equal(u.employeeLoginCodeCell({role:'admin'}).children.length,0);assert.equal(u.employeeLoginCodeCell({role:'user',enabled:false,loginCode:issued[0].code}).children.length,0);
assert(!html.includes('function renderIssuedLoginCodes('));assert(!html.includes('Показать код'));assert(html.includes('<th>Личный код</th>'));console.log('Code column copies the existing value; redundant buttons and lower list removed');})().catch(e=>{console.error(e);process.exitCode=1});
