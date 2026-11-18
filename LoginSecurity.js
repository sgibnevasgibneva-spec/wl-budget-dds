/** Prototype: administrator-issued personal codes. */
let WL_AUTH_EMAIL='';
function wlAuthDigest_(s){return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8)).replace(/=+$/,'');}
function wlAuthToken_(){return Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');}
function wlAuthEpoch_(email){return PropertiesService.getScriptProperties().getProperty('WL_AUTH_EPOCH_'+wlAuthDigest_(email))||'0';}
function wlAuthSessionIndexKey_(email){return 'WL_AUTH_SESSIONS_'+wlAuthDigest_(email);}
function wlAuthSessionKey_(hash){return 'WL_AUTH_SESSION_'+hash;}
function wlAuthSessionIndex_(props,email){try{const ids=JSON.parse(props.getProperty(wlAuthSessionIndexKey_(email))||'[]');return Array.isArray(ids)?ids.filter(x=>typeof x==='string'):[];}catch(_){return [];}}
function wlAuthRevoke_(email){const props=PropertiesService.getScriptProperties();for(const id of wlAuthSessionIndex_(props,email))props.deleteProperty(wlAuthSessionKey_(id));props.deleteProperty(wlAuthSessionIndexKey_(email));props.setProperty('WL_AUTH_EPOCH_'+wlAuthDigest_(email),wlAuthToken_());}
function wlAuthRemember_(token,session){const props=PropertiesService.getScriptProperties(),hash=wlAuthDigest_(token),ids=wlAuthSessionIndex_(props,session.email).filter(id=>id!==hash);ids.push(hash);while(ids.length>12)props.deleteProperty(wlAuthSessionKey_(ids.shift()));props.setProperty(wlAuthSessionKey_(hash),JSON.stringify({email:session.email,epoch:session.epoch}));props.setProperty(wlAuthSessionIndexKey_(session.email),JSON.stringify(ids));}
function wlAuthRules_(email){return wlRules_(wlStore_()).some(r=>r.enabled&&r.email===email);}
function wlAuthCredentialKey_(email){return 'WL_AUTH_CODE_'+wlAuthDigest_(email);}
function wlAuthEraseCode_(email){PropertiesService.getScriptProperties().deleteProperty(wlAuthCredentialKey_(email));wlAuthRevoke_(email);}
function wlGenerateLoginCode_(email){
 return wlLock_(()=>{
  const ss=wlStore_();wlAdmin_(ss);email=String(email||'').trim().toLowerCase();
  if(!wlAuthRules_(email))throw Error('Сначала сохраните активный доступ сотрудника');
  const code=wlAuthToken_().slice(0,20),salt=wlAuthToken_();
  PropertiesService.getScriptProperties().setProperty(wlAuthCredentialKey_(email),JSON.stringify({salt,hash:wlAuthDigest_(salt+'|'+email+'|'+code)}));
  wlAuthRevoke_(email);const cache=CacheService.getScriptCache(),key=wlAuthDigest_(email);cache.remove('blocked:'+key);cache.remove('fail:'+key);return {code};
 });
}
/** Run only in the Apps Script editor, by the project owner, before changing deployment access. */
function wlSetupOwnerCode_(){
 const props=PropertiesService.getScriptProperties(),owner=String(props.getProperty(WL_PREFIX+'OWNER_EMAIL')||'').trim().toLowerCase();
 if(!owner||String(Session.getActiveUser().getEmail()||'').trim().toLowerCase()!==owner)throw Error('Начальную настройку выполняет владелец проекта');
 const result=wlGenerateLoginCode_(owner);Logger.log('Личный код владельца: '+result.code);return 'Готово';
}
function wlVerifyLogin(email,code){
 email=String(email||'').trim().toLowerCase();code=String(code||'').trim().toLowerCase();
 if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)||!/^[a-f0-9]{20}$/.test(code))throw Error('Неверная почта или код');
 const cache=CacheService.getScriptCache(),key=wlAuthDigest_(email),lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  if(cache.get('blocked:'+key))throw Error('Слишком много попыток. Попробуйте позже');
  const raw=PropertiesService.getScriptProperties().getProperty(wlAuthCredentialKey_(email));let credential;try{credential=JSON.parse(raw);}catch(_){}
  if(!credential||!wlAuthRules_(email)||wlAuthDigest_(credential.salt+'|'+email+'|'+code)!==credential.hash){
   const failures=Number(cache.get('fail:'+key)||0)+1;cache.put('fail:'+key,String(failures),900);if(failures>=5){cache.put('blocked:'+key,'1',900);cache.remove('fail:'+key);}throw Error('Неверная почта или код');
  }
  cache.remove('fail:'+key);const token=wlAuthToken_();wlAuthRemember_(token,{email,epoch:wlAuthEpoch_(email)});return {token};
 }finally{lock.releaseLock();}
}
function wlAuthSession_(token){
 if(!/^[a-f0-9]{64}$/i.test(String(token||'')))throw Error('Войдите в приложение');
 const hash=wlAuthDigest_(token),props=PropertiesService.getScriptProperties(),cache=CacheService.getScriptCache(),stored=props.getProperty(wlAuthSessionKey_(hash));
 let session,legacy=false;try{session=JSON.parse(stored||cache.get('session:'+hash));legacy=!stored&&!!session;}catch(_){}
 if(!session||!session.email||(legacy&&(!Number.isFinite(session.expires)||session.expires<=Date.now()))||session.epoch!==wlAuthEpoch_(session.email)||!wlAuthRules_(session.email))throw Error('Сеанс истёк или доступ отозван. Войдите снова');
 if(legacy){const lock=LockService.getScriptLock();lock.waitLock(30000);try{if(session.epoch!==wlAuthEpoch_(session.email)||!wlAuthRules_(session.email))throw Error('Сеанс истёк или доступ отозван. Войдите снова');wlAuthRemember_(token,session);cache.remove('session:'+hash);}finally{lock.releaseLock();}}
 return session.email;
}
function wlLogout(token){
 if(!/^[a-f0-9]{64}$/i.test(String(token||'')))return {ok:true};
 const hash=wlAuthDigest_(token),props=PropertiesService.getScriptProperties(),lock=LockService.getScriptLock();lock.waitLock(30000);
 try{let session;try{session=JSON.parse(props.getProperty(wlAuthSessionKey_(hash))||'null');}catch(_){}props.deleteProperty(wlAuthSessionKey_(hash));CacheService.getScriptCache().remove('session:'+hash);if(session&&session.email){const ids=wlAuthSessionIndex_(props,session.email).filter(id=>id!==hash);props.setProperty(wlAuthSessionIndexKey_(session.email),JSON.stringify(ids));}}finally{lock.releaseLock();}
 return {ok:true};
}
