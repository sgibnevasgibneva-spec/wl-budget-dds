/** Fresh Sheets values for one startup request. Never caches access rules across requests. */
let WL_STARTUP_READ_PROPERTIES=null;

function wlStartupBatchStore_(id){
 const specs=[['Настройки доступа',8],['Справочники',5],['Периоды',6],['Операции',3],['WL_Порядок_строк',3],['WL_Группы_статей',4]];
 const ranges=specs.map(([name,width])=>"'"+name+"'!A:"+String.fromCharCode(64+width));
 const url='https://sheets.googleapis.com/v4/spreadsheets/'+encodeURIComponent(id)+'/values:batchGet?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER&'+ranges.map(r=>'ranges='+encodeURIComponent(r)).join('&');
 const response=UrlFetchApp.fetch(url,{method:'get',muteHttpExceptions:true,followRedirects:false,headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}});
 if(response.getResponseCode()!==200)throw Error('BATCH_HTTP_'+response.getResponseCode());
 const result=JSON.parse(response.getContentText());
 if(!Array.isArray(result.valueRanges)||result.valueRanges.length!==specs.length)throw Error('BATCH_SHAPE');
 const sheets=new Map();
 specs.forEach(([name,width],i)=>{
  const values=result.valueRanges[i].values||[];
  // Sheets serial dates contain local calendar time, not a UTC instant.
  if(['Периоды','Операции','WL_Порядок_строк'].includes(name))values.slice(1).forEach(row=>{
   if(typeof row[0]==='number'&&Number.isFinite(row[0]))row[0]=new Date(Date.UTC(1899,11,30)+Math.floor(row[0])*86400000).toISOString().slice(0,7);
  });
  const columns=Math.max(0,...values.map(row=>row.length));
  const sheet={getLastRow:()=>values.length,getLastColumn:()=>columns,getRange:(row,col,rows=1,cols=1)=>{
   if(row<1||col<1||rows<1||cols<1||col+cols-1>width)throw Error('BATCH_RANGE');
   const read=()=>Array.from({length:rows},(_,r)=>Array.from({length:cols},(_,c)=>values[row+r-1]?.[col+c-1]??''));
   return {getValues:read,getDisplayValues:()=>read().map(r=>r.map(String))};
  }};
  sheets.set(name,sheet);
 });
 if(specs.slice(0,4).some(([name,width])=>!sheets.get(name).getLastRow()||sheets.get(name).getLastColumn()<width))throw Error('BATCH_UNINITIALIZED');
 return {getSheetByName:name=>sheets.get(name)||null,getSpreadsheetTimeZone:()=>Session.getScriptTimeZone()};
}

/** Validate persistent session before reading the sheet. Legacy sessions use the normal migration. */
function wlTryStartupBatch_(token,dispatch){
 if(!/^[a-f0-9]{64}$/i.test(String(token||'')))return null;
 const started=Date.now(),all=PropertiesService.getScriptProperties().getProperties();
 let session;try{session=JSON.parse(all[wlAuthSessionKey_(wlAuthDigest_(token))]||'null');}catch(_){}
 if(!session||!session.email||session.epoch!==(all['WL_AUTH_EPOCH_'+wlAuthDigest_(session.email)]||'0'))return null;
 const id=all[WL_PREFIX+'SPREADSHEET_ID'];if(!id)return null;
 let ss;
 try{ss=wlStartupBatchStore_(id);}catch(e){console.log('WL_STARTUP_BATCH fallback '+String(e.message||'').match(/^BATCH_[A-Z0-9_]+$/)?.[0]);return null;}
 const store=wlStore_,dict=wlDictionaries_,previous=WL_STARTUP_READ_PROPERTIES;
 WL_STARTUP_READ_PROPERTIES={getProperty:key=>all[key]??null,getProperties:()=>({...all})};
 wlStore_=()=>ss;wlDictionaries_=sheet=>dict(sheet,true);
 try{
  const result=wlReadSnapshot_(dispatch);
  result.timing={...result.timing,batchMs:Date.now()-started,requestServerMs:Date.now()-started,readMode:'batch'};
  return result;
 }catch(e){
  // Initialization or template reconciliation needs the regular writable Sheets service.
  if(String(e.message||'')==='STARTUP_NEEDS_RECONCILE')return null;
  throw e;
 }finally{wlStore_=store;wlDictionaries_=dict;WL_STARTUP_READ_PROPERTIES=previous;}
}
