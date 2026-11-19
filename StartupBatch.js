/** Fresh Sheets values for one startup request. Never caches access rules across requests. */
let WL_STARTUP_READ_PROPERTIES=null;
let WL_STARTUP_BATCH_FALLBACK='';

function wlStartupBatchStore_(id){
 const specs=[['Настройки доступа',8],['Справочники',5],['Периоды',6],['Операции',3],['WL_Порядок_строк',3],['WL_Группы_статей',4]];
 const ranges=specs.map(([name,width])=>"'"+name+"'!A:"+String.fromCharCode(64+width));
 // The manifest enables Sheets v4, including its API in the default Cloud project.
 // Apps Script supplies authorization; no extra OAuth token or HTTP request is needed.
 const result=Sheets.Spreadsheets.Values.batchGet(id,{ranges,valueRenderOption:'UNFORMATTED_VALUE',dateTimeRenderOption:'SERIAL_NUMBER'});
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
 WL_STARTUP_BATCH_FALLBACK='';
 if(!/^[a-f0-9]{64}$/i.test(String(token||'')))return null;
 const started=Date.now(),all=PropertiesService.getScriptProperties().getProperties();
 let session;try{session=JSON.parse(all[wlAuthSessionKey_(wlAuthDigest_(token))]||'null');}catch(_){}
 if(!session||!session.email||session.epoch!==(all['WL_AUTH_EPOCH_'+wlAuthDigest_(session.email)]||'0'))return null;
 const id=all[WL_PREFIX+'SPREADSHEET_ID'];if(!id)return null;
 let ss;
 try{ss=wlStartupBatchStore_(id);}catch(e){
  const message=String(e.message||''),known=message.match(/^BATCH_[A-Z0-9_]+$/);
  WL_STARTUP_BATCH_FALLBACK=known?known[0]:/not been used|disabled|accessNotConfigured/i.test(message)?'BATCH_API_DISABLED':/permission|forbidden|403/i.test(message)?'BATCH_PERMISSION_DENIED':/range|sheet.*not found/i.test(message)?'BATCH_RANGE_MISSING':'BATCH_READ_FAILED';
  console.log('WL_STARTUP_BATCH fallback '+WL_STARTUP_BATCH_FALLBACK);return null;
 }
 const store=wlStore_,dict=wlDictionaries_,previous=WL_STARTUP_READ_PROPERTIES;
 WL_STARTUP_READ_PROPERTIES={getProperty:key=>all[key]??null,getProperties:()=>({...all})};
 wlStore_=()=>ss;wlDictionaries_=sheet=>dict(sheet,true);
 try{
  const result=wlReadSnapshot_(dispatch);
  result.timing={...result.timing,batchMs:Date.now()-started,requestServerMs:Date.now()-started,readMode:'batch'};
  return result;
 }catch(e){
  // Initialization or template reconciliation needs the regular writable Sheets service.
  if(String(e.message||'')==='STARTUP_NEEDS_RECONCILE'){WL_STARTUP_BATCH_FALLBACK='TEMPLATE_RECONCILE';return null;}
  throw e;
 }finally{wlStore_=store;wlDictionaries_=dict;WL_STARTUP_READ_PROPERTIES=previous;}
}
