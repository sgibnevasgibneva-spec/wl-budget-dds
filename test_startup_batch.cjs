const fs=require('fs'),vm=require('vm'),assert=require('assert'),crypto=require('crypto');
const token='a'.repeat(64),email='employee@example.test',values=new Map(),cache=new Map();let calls=0,http=200;
const headers=[['Email','Роль','Видит','Меняет','Активен','Вкладки','Правка','ФИО'],['Отделы','Как','Откуда','Расход','Приход'],['Месяц','Статус','Показывать','Создан','Из','Изменён'],['Месяц','Изменено','JSON'],['Месяц','Отдел','Порядок'],['Тип','Статья','Деятельность','Группа']];
const month='2026-11',serial=(Date.UTC(2026,10,1)-Date.UTC(1899,11,30))/86400000;
let tables=[
 [headers[0],[email,'user','HR','HR','true','Отделы','Отделы','Employee']],
 [headers[1],['HR','счет','ООО','Зарплата','Продажа'],['IT']],
 [headers[2],[serial,'Открыт',true]],
 [headers[3],[serial,'',JSON.stringify({id:'a',department:'HR',amounts:{1:25}})],[month,'',JSON.stringify({id:'b',department:'IT',amounts:{1:999}})],[month,'',JSON.stringify({id:'c',department:'HR',amounts:{2:30}})]],
 [headers[4],[serial,'HR',JSON.stringify(['c','a'])]],
 [headers[5]]
];
const props={getProperty:k=>values.get(k)||null,getProperties:()=>Object.fromEntries(values)};
const c={console:{log(){}},Date,Map,Set,Utilities:{DigestAlgorithm:{SHA_256:1},Charset:{UTF_8:1},computeDigest:(_,s)=>[...crypto.createHash('sha256').update(String(s)).digest()],base64EncodeWebSafe:b=>Buffer.from(b).toString('base64url'),formatDate:(_date,_zone,pattern)=>pattern==='yyyy'?'2026':month},Session:{getScriptTimeZone:()=> 'Europe/Moscow',getActiveUser:()=>({getEmail:()=>email})},PropertiesService:{getScriptProperties:()=>props},CacheService:{getScriptCache:()=>({get:k=>cache.get(k)||null,put:(k,v)=>cache.set(k,v),remove:k=>cache.delete(k)})},SpreadsheetApp:{openById(){throw Error('Unexpected sequential sheet read');}},Sheets:{Spreadsheets:{Values:{batchGet:(id,options)=>{calls++;assert.equal(id,'test-sheet');assert.equal(options.ranges.length,6);assert.equal(options.valueRenderOption,'UNFORMATTED_VALUE');if(http!==200)throw Error('permission denied 403');return JSON.parse(JSON.stringify({valueRanges:tables.map(values=>({values}))}));}}}}};
assert(JSON.parse(fs.readFileSync('appsscript.json','utf8')).dependencies.enabledAdvancedServices.some(s=>s.userSymbol==='Sheets'&&s.serviceId==='sheets'&&s.version==='v4'));
vm.createContext(c);for(const file of ['Code.js','LoginSecurity.js','Startup.js','StartupBatch.js'])vm.runInContext(fs.readFileSync(file,'utf8'),c);
values.set('WL_DDS_SPREADSHEET_ID','test-sheet');values.set(c.wlAuthSessionKey_(c.wlAuthDigest_(token)),JSON.stringify({email,epoch:'0'}));values.set(c.wlSkeletonMark_(month),String(Date.now()));
const original=c.wlStore_,read=()=>c.wlTryStartupBatch_(token,()=>{const actual=c.wlAuthSession_(token);assert.equal(actual,email);vm.runInContext('WL_AUTH_EMAIL='+JSON.stringify(actual),c);try{return c.wlStartup_();}finally{vm.runInContext("WL_AUTH_EMAIL=''",c);}});
let result=read();assert.equal(calls,1);assert.equal(result.timing.readMode,'batch');assert.equal(result.boot.current,month);assert.deepEqual(Array.from(result.month.rows,r=>r.id),['c','a']);assert.deepEqual(Array.from(result.boot.user.view),['HR']);assert.equal(result.boot.dict.departments.length,1);assert.equal(c.wlStore_,original);assert.equal(vm.runInContext('WL_STARTUP_READ_PROPERTIES',c),null);
tables[2][1][1]='Закрыт';result=read();assert.equal(result.month.month.status,'Закрыт');
tables[0][1][2]='IT';result=read();assert.deepEqual(Array.from(result.month.rows,r=>r.id),['b']);assert.equal(result.boot.dict.departments[0].name,'IT');
tables[0][1][4]='false';assert.throws(read,/Сеанс/);assert.equal(c.wlStore_,original);assert.equal(vm.runInContext('WL_STARTUP_READ_PROPERTIES',c),null);tables[0][1][4]='true';
values.set('WL_AUTH_EPOCH_'+c.wlAuthDigest_(email),'revoked');let before=calls;assert.equal(read(),null);assert.equal(calls,before);values.delete('WL_AUTH_EPOCH_'+c.wlAuthDigest_(email));
values.delete(c.wlSkeletonMark_(month));assert.equal(read(),null);assert.equal(c.wlStore_,original);values.set(c.wlSkeletonMark_(month),String(Date.now()));
http=403;assert.equal(read(),null);assert.equal(vm.runInContext('WL_STARTUP_BATCH_FALLBACK',c),'BATCH_PERMISSION_DENIED');http=200;
tables[0]=[['Email'],[email]];assert.equal(read(),null); // Unmigrated ACL needs the writable regular service.
assert.equal(c.wlTryStartupBatch_('bad',()=>assert.fail()),null);
console.log('Batch startup: one fresh Sheets read, department isolation, fresh permission/closure checks, epoch revocation, serial dates, row order and safe fallback.');
