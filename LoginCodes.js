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
   if(retry){const code=cache.get(key);if(code&&wlAuthDigest_(credential.salt+'|'+email+'|'+code)===credential.hash)return {...base,code,codeVersion:version,status:'reused'};return {...base,codeVersion:version,status:'existing'};}
   if(raw&&!Object.prototype.hasOwnProperty.call(replacements,email))return {...base,codeVersion:version,status:'existing'};
   const code=wlAuthToken_().slice(0,20),salt=wlAuthToken_(),record=JSON.stringify({salt,hash:wlAuthDigest_(salt+'|'+email+'|'+code),requestId});
   // Persist only the hash. The response is available for a short retry window.
   cache.put(key,code,600);
   wlAuthRevoke_(email);props.setProperty(wlAuthCredentialKey_(email),record);
   const digest=wlAuthDigest_(email);cache.remove('blocked:'+digest);cache.remove('fail:'+digest);
   return {...base,code,codeVersion:wlAuthDigest_(record),status:'created'};
  })};
 });
}
