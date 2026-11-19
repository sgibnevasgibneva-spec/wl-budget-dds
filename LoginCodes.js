/** Admin-only issuance. A repeated request never replaces an issued employee code. */
function wlIssueEmployeeCodes_(emails,requestId,replacements){
 if(!Array.isArray(emails)||!emails.length||emails.length>100||typeof requestId!=='string'||!/^[a-zA-Z0-9_-]{16,100}$/.test(requestId))throw Error('Некорректный запрос кодов');
 emails=[...new Set(emails.map(x=>String(x).trim().toLowerCase()))];
 replacements=replacements&&typeof replacements==='object'?replacements:{};
 return wlLock_(()=>{
  const ss=wlStore_();wlAdmin_(ss);
  const rules=wlRules_(ss),props=PropertiesService.getScriptProperties(),cache=CacheService.getScriptCache();
  // Validate every recipient before issuing any code. Administrators are excluded.
  const targets=emails.map(email=>{
   const found=rules.filter(r=>r.email===email&&r.enabled&&r.role!=='admin');
   if(found.length!==1)throw Error('Проверьте активный доступ сотрудника: '+email);
   const raw=props.getProperty(wlAuthCredentialKey_(email))||'',version=raw?wlAuthDigest_(raw):'';
   let credential=null;if(raw){try{credential=JSON.parse(raw);}catch(_){throw Error('Не удалось прочитать код сотрудника: '+email);}}
   const retry=credential&&credential.requestId===requestId;
   if(Object.prototype.hasOwnProperty.call(replacements,email)&&!retry&&replacements[email]!==version)throw Error('Код уже изменён. Обновите список сотрудников: '+email);
   return {rule:found[0],raw,version,credential,retry};
  });
  return {items:targets.map(({rule,raw,version,credential,retry})=>{
   const email=rule.email,key='issued:'+wlAuthDigest_(email+'|'+requestId),base={email,name:rule.name||'',departments:rule.view||[]};
   if(retry){const code=wlEmployeeCodeDisplay_(email,raw)||cache.get(key);if(code&&wlAuthDigest_(credential.salt+'|'+email+'|'+code)===credential.hash)return {...base,code,codeVersion:version,status:'reused'};return {...base,codeVersion:version,status:'existing'};}
   if(raw&&!Object.prototype.hasOwnProperty.call(replacements,email))return {...base,codeVersion:version,status:'existing'};
   const code=wlAuthToken_().slice(0,20),salt=wlAuthToken_(),record=JSON.stringify({salt,hash:wlAuthDigest_(salt+'|'+email+'|'+code),requestId});
   // Login verifies the hash; only the admin settings may read the display copy.
   cache.put(key,code,600);
   wlAuthRevoke_(email);props.setProperty(wlAuthCredentialKey_(email),record);
   props.setProperty(wlEmployeeCodeDisplayKey_(email),JSON.stringify({code,codeVersion:wlAuthDigest_(record)}));
   const digest=wlAuthDigest_(email);cache.remove('blocked:'+digest);cache.remove('fail:'+digest);
   return {...base,code,codeVersion:wlAuthDigest_(record),status:'created'};
  })};
 });
}

function wlEmployeeCodeDisplayKey_(email){return 'WL_AUTH_DISPLAY_'+wlAuthDigest_(email);}
/** Called only after admin authorization; credentials remain unchanged during migration. */
function wlEmployeeCodeDisplay_(email,raw){
 if(!raw)return '';const props=PropertiesService.getScriptProperties();let credential,display;
 try{credential=JSON.parse(raw);display=JSON.parse(props.getProperty(wlEmployeeCodeDisplayKey_(email))||'null');}catch(_){return '';}
 const version=wlAuthDigest_(raw),valid=code=>typeof code==='string'&&/^[a-f0-9]{20}$/.test(code)&&wlAuthDigest_(credential.salt+'|'+email+'|'+code)===credential.hash;
 if(display&&display.codeVersion===version&&valid(display.code))return display.code;
 const code=credential.requestId?CacheService.getScriptCache().get('issued:'+wlAuthDigest_(email+'|'+credential.requestId)):'';
 if(!valid(code))return '';
 props.setProperty(wlEmployeeCodeDisplayKey_(email),JSON.stringify({code,codeVersion:version}));return code;
}
/** Preserve already-issued codes, never issue replacements or invalidate sessions. */
function wlRememberEmployeeCodes_(items){
 if(!Array.isArray(items)||!items.length||items.length>100)throw Error('Некорректный список кодов');
 return wlLock_(()=>{
  const ss=wlStore_();wlAdmin_(ss);const rules=wlRules_(ss),props=PropertiesService.getScriptProperties(),seen=new Set();
  const checked=items.map(item=>{
   const email=String(item.email||'').trim().toLowerCase(),code=String(item.code||'').trim().toLowerCase();
   if(seen.has(email)||!/^[a-f0-9]{20}$/.test(code)||rules.filter(r=>r.email===email&&r.enabled&&r.role!=='admin').length!==1)throw Error('Проверьте почту и код сотрудника');seen.add(email);
   const raw=props.getProperty(wlAuthCredentialKey_(email));let credential;try{credential=JSON.parse(raw);}catch(_){}
   if(!credential||wlAuthDigest_(credential.salt+'|'+email+'|'+code)!==credential.hash)throw Error('Код не соответствует текущему доступу: '+email);
   return {email,code,codeVersion:wlAuthDigest_(raw)};
  });
  checked.forEach(x=>props.setProperty(wlEmployeeCodeDisplayKey_(x.email),JSON.stringify({code:x.code,codeVersion:x.codeVersion})));
  return wlGetSettings_();
 });
}
