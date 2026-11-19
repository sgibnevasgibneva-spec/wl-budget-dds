const fs=require('fs'),vm=require('vm'),assert=require('assert'),crypto=require('crypto');
const values=new Map(),entries=new Map(),rules=[{email:'owner@test.local',role:'admin',enabled:true},{email:'a@test.local',name:'A',role:'user',enabled:true,view:['HR']},{email:'b@test.local',name:'B',role:'user',enabled:true,view:['IT']},{email:'off@test.local',role:'user',enabled:false}];let admin=true;
const props={getProperty:k=>values.get(k)||null,setProperty:(k,v)=>values.set(k,v),deleteProperty:k=>values.delete(k)},cache={get:k=>entries.get(k)||null,put:(k,v)=>entries.set(k,v),remove:k=>entries.delete(k)};
const c={Utilities:{DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},getUuid:()=>crypto.randomUUID(),computeDigest:(_a,s)=>[...crypto.createHash('sha256').update(String(s)).digest()],base64EncodeWebSafe:a=>Buffer.from(a).toString('base64url')},PropertiesService:{getScriptProperties:()=>props},CacheService:{getScriptCache:()=>cache},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},wlStore_:()=>({}),wlRules_:()=>rules,wlLock_:f=>f(),wlAdmin_:()=>{if(!admin)throw Error('admin only');}};
vm.createContext(c);vm.runInContext(fs.readFileSync('LoginSecurity.js','utf8')+'\n'+fs.readFileSync('LoginCodes.js','utf8'),c);
const ownerCode=c.wlGenerateLoginCode_('owner@test.local').code,ownerSession=c.wlVerifyLogin('owner@test.local',ownerCode).token,ownerCredential=values.get(c.wlAuthCredentialKey_('owner@test.local'));
const request=crypto.randomUUID(),emails=['a@test.local','b@test.local'];
assert.throws(()=>c.wlIssueEmployeeCodes_([...emails,'owner@test.local'],request,{}));assert(!values.has(c.wlAuthCredentialKey_(emails[0])));
assert.throws(()=>c.wlIssueEmployeeCodes_([...emails,'off@test.local'],request,{}));assert(!values.has(c.wlAuthCredentialKey_(emails[0])));
admin=false;assert.throws(()=>c.wlIssueEmployeeCodes_(emails,request,{}));admin=true;
const first=c.wlIssueEmployeeCodes_(emails,request,{}).items;assert.equal(first.length,2);assert.notEqual(first[0].code,first[1].code);first.forEach(x=>{assert.match(x.code,/^[a-f0-9]{20}$/);assert(!values.get(c.wlAuthCredentialKey_(x.email)).includes(x.code));});
const employeeSession=c.wlVerifyLogin(first[0].email,first[0].code).token,epoch=c.wlAuthEpoch_(first[0].email);
const repeated=c.wlIssueEmployeeCodes_(emails,request,{}).items;assert.deepEqual(Array.from(repeated,x=>x.code),Array.from(first,x=>x.code));assert.equal(c.wlAuthEpoch_(first[0].email),epoch);assert.equal(c.wlAuthSession_(employeeSession),first[0].email);
const another=c.wlIssueEmployeeCodes_(emails,crypto.randomUUID(),{}).items;assert(another.every(x=>x.status==='existing'&&!x.code));assert.equal(c.wlAuthSession_(employeeSession),first[0].email);
entries.clear();const expired=c.wlIssueEmployeeCodes_(emails,request,{}).items;assert.deepEqual(Array.from(expired,x=>x.code),Array.from(first,x=>x.code));assert.equal(c.wlAuthSession_(employeeSession),first[0].email);
const replaceRequest=crypto.randomUUID(),replacement={[first[0].email]:first[0].codeVersion},changed=c.wlIssueEmployeeCodes_([first[0].email],replaceRequest,replacement).items[0];assert.notEqual(changed.code,first[0].code);assert.throws(()=>c.wlAuthSession_(employeeSession));assert.throws(()=>c.wlVerifyLogin(first[0].email,first[0].code));assert(c.wlVerifyLogin(first[0].email,changed.code).token);
assert.equal(c.wlIssueEmployeeCodes_([first[0].email],replaceRequest,replacement).items[0].code,changed.code);
assert.throws(()=>c.wlIssueEmployeeCodes_([first[0].email],crypto.randomUUID(),replacement));
assert.equal(values.get(c.wlAuthCredentialKey_('owner@test.local')),ownerCredential);assert.equal(c.wlAuthSession_(ownerSession),'owner@test.local');
console.log('Bulk employee codes: replay-safe, existing codes preserved, guarded replacement, admin excluded, inactive recipients blocked');

(async()=>{
 const html=fs.readFileSync('Index.html','utf8'),source=html.slice(html.indexOf('async function issueLoginCodes('),html.indexOf('function employeeLoginCodeCell('));let resolve,requests=0,renders=0;const status={};
 const u={loginCodesBusy:false,loginCodesRequest:null,settings:{rules:[{email:emails[0],codeVersion:''}]},genId:()=>crypto.randomUUID(),document:{querySelectorAll:()=>[{disabled:false}]},$:()=>status,renderSettings:()=>renders++,call:()=>{requests++;return new Promise(r=>resolve=r);}};vm.createContext(u);vm.runInContext(source,u);
 const one=u.issueLoginCodes([emails[0]],{}),two=u.issueLoginCodes([emails[0]],{});assert.equal(requests,1);resolve({items:[first[0]]});await Promise.all([one,two]);assert.equal(renders,1);assert.equal(u.settings.rules[0].loginCode,first[0].code);assert.equal(u.loginCodesBusy,false);assert.equal(u.settings.rules[0].codeVersion,first[0].codeVersion);
 let rejectedId;u.call=(_n,_e,id)=>{rejectedId=id;return Promise.reject(Error('connection'));};await u.issueLoginCodes([emails[0]],{});assert.equal(u.loginCodesBusy,false);u.call=(_n,_e,id)=>{assert.equal(id,rejectedId);return Promise.resolve({items:[first[0]]});};await u.issueLoginCodes([emails[0]],{});assert.equal(u.loginCodesRequest,null);
 console.log('UI double-click sends one request; a failed response retries with the same issuance ID');
})().catch(e=>{console.error(e);process.exitCode=1});
