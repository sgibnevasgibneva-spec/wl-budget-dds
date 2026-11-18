/** A read snapshot lives for this startup request only. Permissions are read afresh on every RPC. */
function wlReadSnapshot_(read){
 const originals={store:wlStore_,rules:wlRules_,dict:wlDictionaries_,months:wlMonthList_,user:wlUser_};
 const memo=fn=>{let ready=false,value;return (...args)=>{if(!ready){value=fn(...args);ready=!Array.isArray(value)||value.length>0;}return value;};};
 wlStore_=memo(originals.store);wlRules_=memo(originals.rules);wlMonthList_=memo(originals.months);wlUser_=memo(originals.user);
 const dictionaries=memo(originals.dict);wlDictionaries_=(...args)=>JSON.parse(JSON.stringify(dictionaries(...args)));
 try{return read();
 }finally{
  wlStore_=originals.store;wlRules_=originals.rules;wlDictionaries_=originals.dict;wlMonthList_=originals.months;wlUser_=originals.user;
 }
}

function wlStartup_(){
 const started=Date.now(),boot=wlBoot_(),bootDone=Date.now();
 const month=boot.user.tabs.some(t=>t==='Отделы'||t==='ЦФО')?wlGetMonth_(boot.current):null;
 return {boot,month,timing:{bootMs:bootDone-started,monthMs:Date.now()-bootDone,serverMs:Date.now()-started}};
}
