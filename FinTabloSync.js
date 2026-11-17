/**
 * ФинТабло (ТОЛЬКО ЧТЕНИЕ): получаем справочник через официальный API.
 * В этом файле НЕТ POST/PATCH/PUT/DELETE запросов к ФинТабло.
 * Только администратор может сохранять API-ключ и применить изменения в НАШЕМ Sheets.
 * Ключ хранится в ScriptProperties; не выводится в интерфейс и не попадает в ZIP.
 */
const WL_FT_TOKEN_PROP_='WL_DDS_FINTABLO_API_TOKEN';
const WL_FT_CLOSED_YEARS_PROP_='WL_DDS_CLOSED_YEARS';
const WL_FT_ENDPOINT_='https://api.fintablo.ru/v1/category';

function wlFinTabloSetToken(token){
  const ss=wlStore_();wlAdmin_(ss);
  const value=String(token||'').trim();
  if(!/^[A-Za-z0-9_.-]{12,240}$/.test(value))throw Error('Некорректный формат API-ключа');
  PropertiesService.getScriptProperties().setProperty(WL_FT_TOKEN_PROP_,value);
  return {saved:true};
}
function wlFinTabloFetch_(){
  const token=PropertiesService.getScriptProperties().getProperty(WL_FT_TOKEN_PROP_);
  if(!token)throw Error('Сначала сохраните ключ ФинТабло в разделе «Справочники»');
  // ВАЖНО: только GET; никогда не пересылаем ключ браузеру, в логах нет токена.
  const res=UrlFetchApp.fetch(WL_FT_ENDPOINT_,{
    method:'get',muteHttpExceptions:true,followRedirects:false,
    headers:{Authorization:'Bearer '+token,Accept:'application/json'}
  });
  const code=res.getResponseCode();
  if(code!==200)throw Error('ФинТабло не вернул справочник (HTTP '+code+'). Проверьте ключ и разрешение UrlFetchApp. Никаких данных не изменено.');
  let body;try{body=JSON.parse(res.getContentText());}catch(_){throw Error('ФинТабло вернул некорректный JSON');}
  let list=Array.isArray(body)?body:(Array.isArray(body.data)?body.data:(Array.isArray(body.items)?body.items:(Array.isArray(body.categories)?body.categories:null)));
  if(!list||!list.length)throw Error('API вернул пустой или неизвестный формат справочника. Применение отменено.');
  return list;
}
function wlFinTabloField_(o,keys){
  for(const k of keys)if(o&&o[k]!==undefined&&o[k]!==null&&o[k]!=='')return o[k];
  return '';
}
function wlFinTabloNorm_(value){return String(value||'').trim().toLocaleLowerCase('ru').replace(/\s+/g,' ');}
function wlFinTabloType_(raw){
  const val=wlFinTabloNorm_(raw);
  if(['income','incomes','receipt','receipts','incoming','revenue','приход','поступление','поступления'].includes(val))return 'income';
  if(['outcome','outcomes','expense','expenses','spending','outgoing','расход','расходы','списание','списания'].includes(val))return 'outcome';
  if(['transfer','transfers','перевод','переводы'].includes(val))return 'transfer';
  return '';
}
/** Strict, fail-closed parsing. Unknown classification is never guessed or applied. */
function wlFinTabloNormalize_(source){
  const errors=[],flat=[];
  function push(raw,derivedParent,depth){
    if(depth>5)throw Error('Слишком глубокая вложенность статей ФинТабло');
    if(!raw||typeof raw!=='object'||Array.isArray(raw))throw Error('Некорректная статья из API');
    const name=String(wlFinTabloField_(raw,['name','title','label'])||'').trim();
    if(!name)throw Error('Статья API без названия — изменения не применены');
    const rawId=wlFinTabloField_(raw,['id','uuid','categoryId']);
    const id=String(rawId||'').trim()||'unnamed:'+name;
    const parentObj=wlFinTabloField_(raw,['parent','parentCategory']);
    const parentId=String(wlFinTabloField_(raw,['parentId','parentCategoryId','parentCategoryID','parent_id','subCategoryForId']) ||
       (parentObj&&typeof parentObj==='object'?parentObj.id:'') ||derivedParent||'').trim();
    const parentName=String(wlFinTabloField_(raw,['parentName','subCategoryFor','subcategoryFor','subCategoryOf']) ||
       (parentObj&&typeof parentObj==='object'?parentObj.name:typeof parentObj==='string'?parentObj:'')||'').trim();
    let rawGroup=wlFinTabloField_(raw,['group','operationGroup','operationType','transactionType','cashFlowType','type','categoryType']);
    if(!rawGroup&&raw.isIncome===true)rawGroup='income';
    if(!rawGroup&&raw.isExpense===true)rawGroup='outcome';
    const group=wlFinTabloType_(rawGroup);
    const archived=raw.archived===true||raw.isArchived===true||wlFinTabloNorm_(raw.status)==='archived';
    flat.push({id,name,parentId,parentName,group,archived});
    const children=Array.isArray(raw.children)?raw.children:(Array.isArray(raw.subcategories)?raw.subcategories:[]);
    children.forEach(child=>push(child,id,depth+1));
  }
  source.forEach(raw=>push(raw,'',0));
  if(flat.length>2500)throw Error('ФинТабло вернул неожиданно много статей — проверка остановлена');
  const byId=new Map(),byName=new Map();
  for(const item of flat){
    if(byId.has(item.id))errors.push('Повтор ID: '+item.id);
    byId.set(item.id,item);byName.set(wlFinTabloNorm_(item.name),item);
  }
  const parentOf=item=>byId.get(item.parentId)||byName.get(wlFinTabloNorm_(item.parentName));
  const containsArchive=(item,depth)=>{
    if(depth>6)throw Error('Цикл во вложенности справочника');
    if(item.archived||wlFinTabloNorm_(item.name)==='архив')return true;
    const parent=parentOf(item);return parent?containsArchive(parent,depth+1):wlFinTabloNorm_(item.parentName)==='архив';
  };
  let archivedCount=0,parentCount=0;const candidates=[];
  for(const item of flat){
    if(containsArchive(item,0)){archivedCount++;continue;}
    if(flat.some(ch=>ch!==item&&(ch.parentId===item.id||wlFinTabloNorm_(ch.parentName)===wlFinTabloNorm_(item.name)))){parentCount++;continue;}
    let group=item.group,parent=parentOf(item),steps=0;
    while(!group&&parent&&steps++<6){group=parent.group;parent=parentOf(parent);}
    if(!group){errors.push('Не определён тип поступления/списания: '+item.name);continue;}
    if(group==='transfer')continue;
    if(!item.id||item.id.startsWith('unnamed:'))errors.push('Не указан ID статьи: '+item.name);
    candidates.push({id:item.id,name:item.name,group,parentId:item.parentId});
  }
  const duplicated=new Set(),counts=new Map();
  candidates.forEach(item=>{const key=wlFinTabloNorm_(item.name);counts.set(key,(counts.get(key)||0)+1);});
  counts.forEach((count,name)=>{if(count>1)duplicated.add(name);});
  if(duplicated.size)errors.push('Совпадающие названия разных статей. Требуется отдельное сопоставление: '+[...duplicated].join(', '));
  const expenses=[...new Set(candidates.filter(x=>x.group==='outcome').map(x=>x.name))].sort((a,b)=>a.localeCompare(b,'ru'));
  const incomes=[...new Set(candidates.filter(x=>x.group==='income').map(x=>x.name))].sort((a,b)=>a.localeCompare(b,'ru'));
  if(!expenses.length&&!incomes.length)errors.push('Не обнаружены статьи поступлений и списаний');
  return {expenses,incomes,candidates,archivedCount,parentCount,warnings:errors};
}
function wlFinTabloDigest_(data){
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(data),Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/,'');
}
/** Read-only comparison. Do not infer ownership of pre-existing plans. */
function wlFinTabloBudgetPreview(month){
 month=wlKey_(month);const ss=wlStore_(),user=wlAdmin_(ss);wlRequireTab_(user,'ФинТабло',false);wlFindMonth_(ss,month,false);
 const token=PropertiesService.getScriptProperties().getProperty(WL_FT_TOKEN_PROP_);if(!token)throw Error('Сначала сохраните ключ ФинТабло');
 const cat=wlFinTabloNormalize_(wlFinTabloFetch_()),local=new Map(),warnings=cat.warnings.slice(),dict=wlDictionaries_(ss),income=new Set(dict.incomeCategories.map(wlBddsNormalize_)),expense=new Set(dict.categories.map(wlBddsNormalize_));
 for(const e of wlOpsRead_(wlOps_(ss),month)){const r=e.row,name=String(r.category||'').trim(),amount=Object.values(r.amounts||{}).reduce((s,v)=>s+(Number(v)||0),0);if(!amount)continue;if(!name){warnings.push('Есть сумма без статьи в отделе '+r.department);continue;}
  const normalized=wlBddsNormalize_(name),type=r.cashFlowType==='income'?'income':r.cashFlowType==='expense'?'expense':income.has(normalized)&&!expense.has(normalized)?'income':'expense',key=type+'|'+normalized,old=local.get(key)||{type,name,amount:0};old.amount+=amount;local.set(key,old);
 }
 const date=month.slice(5)+'.'+month.slice(0,4),response=UrlFetchApp.fetch('https://api.fintablo.ru/v1/bdds?date='+encodeURIComponent(date),{method:'get',muteHttpExceptions:true,followRedirects:false,headers:{Authorization:'Bearer '+token,Accept:'application/json'}}),code=response.getResponseCode();
 if(code!==200)throw Error('Не удалось прочитать план ФинТабло (HTTP '+code+'). Данные не изменены.');let body;try{body=JSON.parse(response.getContentText());}catch(_){throw Error('Некорректный ответ ФинТабло');}
 if(!body||!Array.isArray(body.items))throw Error('Неизвестный формат ответа плана ФинТабло. Сверка остановлена.');
 const remote=body.items;if(remote.length>=1000)warnings.push('Слишком много записей ФинТабло: полнота ответа не подтверждена');
 const result=[];for(const item of local.values()){
  const matches=cat.candidates.filter(c=>wlBddsNormalize_(c.name)===wlBddsNormalize_(item.name)&&c.group===(item.type==='income'?'income':'outcome'));
  if(item.amount<0){result.push({...item,status:'Отрицательная сумма — проверить'});continue;}
  if(matches.length!==1||!/^\d+$/.test(matches[0].id)){result.push({...item,status:'Статья не сопоставлена однозначно'});continue;}
  const id=Number(matches[0].id),entries=remote.filter(r=>Number(r.categoryId)===id&&String(r.date)===date),unknown=entries.some(r=>!Number.isFinite(Number(r.value))||r.value==null||r.value==='');
  if(unknown){result.push({...item,status:'Некорректные данные ФинТабло'});continue;}
  const amount=entries.reduce((s,r)=>s+Number(r.value),0),status=entries.length===0?'Нет плана в ФинТабло':entries.length>1?'Несколько записей — проверить':entries[0].directionId!=null?'Есть направление — проверить':Math.abs(amount-item.amount)<0.005?'Совпадает':'Сумма отличается';
  result.push({...item,remoteAmount:entries.length?amount:null,remoteCount:entries.length,status});
 }
 return {month,rows:result.sort((a,b)=>a.name.localeCompare(b.name,'ru')),warnings,remoteCount:remote.length};
}
function wlFinTabloLocal_(ss){const d=wlDictionaries_(ss);return {expenses:d.categories.slice().sort(),incomes:d.incomeCategories.slice().sort()};}
function wlFinTabloPreview(){
  const ss=wlStore_();wlAdmin_(ss);
  const fresh=wlFinTabloNormalize_(wlFinTabloFetch_());
  const local=wlFinTabloLocal_(ss),current=new Set([...local.expenses,...local.incomes]);
  const remote=new Set([...fresh.expenses,...fresh.incomes]);
  const added=[...remote].filter(v=>!current.has(v)).sort();
  const removed=[...current].filter(v=>!remote.has(v)).sort();
  // Used articles are never erased from historic budget rows; only removed from NEW choice lists.
  const used=new Set(wlOpsRead_(wlOps_(ss)).map(entry=>String(entry.row.category||'')));
  const changes=[];
  for(const [type,field] of [['outcome','expenses'],['income','incomes']]){
    const oldSet=new Set(local[field]),newSet=new Set(fresh[field]);
    fresh[field].filter(name=>!oldSet.has(name)).forEach(name=>changes.push({key:type+'|add|'+name,type,action:'add',name,used:false}));
    local[field].filter(name=>!newSet.has(name)).forEach(name=>changes.push({key:type+'|remove|'+name,type,action:'remove',name,used:used.has(name)}));
  }
  return {remoteHash:wlFinTabloDigest_([fresh.expenses,fresh.incomes]),localHash:wlFinTabloDigest_(local),changes,
    added,removed,usedRemoved:removed.filter(x=>used.has(x)),
    outcomeCount:fresh.expenses.length,incomeCount:fresh.incomes.length,
    archivedCount:fresh.archivedCount,parentCount:fresh.parentCount,warnings:fresh.warnings};
}
function wlFinTabloApply(remoteHash,localHash){
  // Fetch outside lock so ten employees are never waiting for a network request.
  const ss=wlStore_(),user=wlAdmin_(ss);
  const fresh=wlFinTabloNormalize_(wlFinTabloFetch_());
  if(fresh.warnings.length)throw Error('В ответе API есть неоднозначности. Сначала проверьте справочник ФинТабло; изменения отменены.');
  if(wlFinTabloDigest_([fresh.expenses,fresh.incomes])!==remoteHash)throw Error('Список ФинТабло изменился. Запустите проверку ещё раз.');
  return wlLock_(()=>{
    wlAdmin_(ss);
    const local=wlFinTabloLocal_(ss);
    if(wlFinTabloDigest_(local)!==localHash)throw Error('Локальный справочник изменился после проверки. Запустите проверку повторно.');
    // Snapshot before editing the dictionary. This does NOT change budget operations.
    const b=wlNewSheet_(ss,'Архив справочников ФинТабло',['Время','Администратор','Списания JSON','Поступления JSON','Хэш ФинТабло']);
    b.appendRow([new Date(),user.email,JSON.stringify(local.expenses),JSON.stringify(local.incomes),remoteHash]);
    const sh=wlDictSheet_(ss),old=Math.max(sh.getLastRow()-1,0),length=Math.max(old,fresh.expenses.length,fresh.incomes.length);
    sh.getRange(1,4,1,2).setValues([['Статьи списаний','Статьи поступлений']]);
    if(length){
      const values=Array.from({length},(_,i)=>[fresh.expenses[i]||'',fresh.incomes[i]||'']);
      sh.getRange(2,4,length,2).setValues(values);
    }
    const ref=wlNewSheet_(ss,'Статьи ФинТабло — ID',['ID','Название','Тип','ID родителя','Дата проверки']);
    const max=ref.getLastRow()-1;if(max>0)ref.getRange(2,1,max,5).clearContent();
    if(fresh.candidates.length)ref.getRange(2,1,fresh.candidates.length,5).setValues(fresh.candidates.map(c=>[c.id,c.name,c.group,c.parentId,new Date()]));
    wlInvalidateDictionaries_();return {ok:true,expenses:fresh.expenses.length,incomes:fresh.incomes.length};
  });
}

/** Year lock: present even if a month status is manually changed in the periods sheet. */
function wlYearClosed_(year){
  const closed=JSON.parse(PropertiesService.getScriptProperties().getProperty(WL_FT_CLOSED_YEARS_PROP_)||'[]');
  return closed.includes(String(year));
}
function wlYearManage(action,year){
  year=String(year||'');if(!/^(20|21)\d{2}$/.test(year))throw Error('Неверный год');
  if(action!=='close'&&action!=='open')throw Error('Недопустимое действие с годом');
  return wlLock_(()=>{
    const ss=wlStore_(),user=wlAdmin_(ss),props=PropertiesService.getScriptProperties();
    let closed=JSON.parse(props.getProperty(WL_FT_CLOSED_YEARS_PROP_)||'[]');
    const sh=wlMonthSheet_(ss),list=wlMonthList_(ss).filter(m=>m.key.slice(0,4)===year);
    if(!list.length)throw Error('В выбранном году нет месяцев');
    const already=closed.includes(year);
    if((action==='close'&&already)||(action==='open'&&!already))throw Error('Год уже в этом состоянии');
    const saveKey='WL_DDS_YEAR_PREVIOUS_STATUS_'+year;
    if(action==='close'){
      const original=list.map(m=>({key:m.key,status:m.status}));
      props.setProperty(saveKey,JSON.stringify(original));
      list.forEach(m=>sh.getRange(m.row,2).setValue('Закрыт'));
      closed.push(year);
    }else{
      const old=JSON.parse(props.getProperty(saveKey)||'[]');
      const oldMap=new Map(old.map(e=>[e.key,e.status]));
      list.forEach(m=>sh.getRange(m.row,2).setValue(oldMap.get(m.key)||'Открыт'));
      closed=closed.filter(v=>v!==year);
    }
    props.setProperty(WL_FT_CLOSED_YEARS_PROP_,JSON.stringify(closed));
    const audit=wlNewSheet_(ss,'Журнал закрытия годов',['Время','Администратор','Год','Действие','Месяцев']);
    audit.appendRow([new Date(),user.email,year,action==='close'?'Закрыт':'Открыт',list.length]);
    return {ok:true,months:wlMonthList_(ss).filter(m=>m.active).map(m=>({key:m.key,status:m.status})),year,closed:action==='close'};
  });
}

/** Selective, explicitly approved import into OUR dictionary; never writes to FinTablo. */
function wlFinTabloApplySelected(remoteHash,localHash,selectedKeys){
  if(!Array.isArray(selectedKeys)||!selectedKeys.length||selectedKeys.length>500)throw Error('Выберите статьи для применения');
  const ss=wlStore_(),admin=wlAdmin_(ss);
  // No outbound mutation: GET the source of truth, then apply approved local diffs only.
  const fresh=wlFinTabloNormalize_(wlFinTabloFetch_());
  const blockers=fresh.warnings.filter(w=>!w.startsWith('Совпадающие названия разных статей.'));
  if(blockers.length)throw Error('ФинТабло вернул неоднозначные данные: '+blockers.join('; '));
  if(wlFinTabloDigest_([fresh.expenses,fresh.incomes])!==remoteHash)throw Error('Статьи ФинТабло изменились. Повторите проверку');
  return wlLock_(()=>{
    wlAdmin_(ss);
    const local=wlFinTabloLocal_(ss);
    if(wlFinTabloDigest_(local)!==localHash)throw Error('Наш справочник изменился после просмотра. Повторите проверку');
    const valid=new Map();
    [['outcome','expenses'],['income','incomes']].forEach(([type,field])=>{
      const oldSet=new Set(local[field]),newSet=new Set(fresh[field]);
      fresh[field].filter(n=>!oldSet.has(n)).forEach(n=>valid.set(type+'|add|'+n,{type,field,action:'add',name:n}));
      local[field].filter(n=>!newSet.has(n)).forEach(n=>valid.set(type+'|remove|'+n,{type,field,action:'remove',name:n}));
    });
    if(new Set(selectedKeys).size!==selectedKeys.length)throw Error('Повтор статьи в выборе');
    const selected=selectedKeys.map(k=>valid.get(String(k)));
    if(selected.some(x=>!x))throw Error('Список изменился. Повторите проверку, прежде чем применять');
    const next={expenses:local.expenses.slice(),incomes:local.incomes.slice()};
    selected.forEach(item=>{
      if(item.action==='add')next[item.field].push(item.name);
      else next[item.field]=next[item.field].filter(x=>x!==item.name);
    });
    next.expenses=[...new Set(next.expenses)].sort((a,b)=>a.localeCompare(b,'ru'));
    next.incomes=[...new Set(next.incomes)].sort((a,b)=>a.localeCompare(b,'ru'));
    const audit=wlNewSheet_(ss,'Архив справочников ФинТабло',['Время','Администратор','Списания JSON','Поступления JSON','Хэш ФинТабло']);
    audit.appendRow([new Date(),admin.email,JSON.stringify(local.expenses),JSON.stringify(local.incomes),remoteHash]);
    const sh=wlDictSheet_(ss),previous=Math.max(sh.getLastRow()-1,0),count=Math.max(previous,next.expenses.length,next.incomes.length);
    sh.getRange(1,4,1,2).setValues([['Статьи списаний','Статьи поступлений']]);
    if(count)sh.getRange(2,4,count,2).setValues(Array.from({length:count},(_,i)=>[next.expenses[i]||'',next.incomes[i]||'']));
    // IDs are stored ONLY for approved categories, for a future separate BDDs planning integration.
    const ref=wlNewSheet_(ss,'Статьи ФинТабло — ID',['ID','Название','Тип','ID родителя','Дата проверки']);
    const previousRefs=Math.max(ref.getLastRow()-1,0);
    if(previousRefs)ref.getRange(2,1,previousRefs,5).clearContent();
    const approved=fresh.candidates.filter(c=>(c.group==='income'?next.incomes:next.expenses).includes(c.name));
    if(approved.length){
      if(ref.getMaxRows()<approved.length+1)ref.insertRowsAfter(ref.getMaxRows(),approved.length+1-ref.getMaxRows());
      ref.getRange(2,1,approved.length,5).setValues(approved.map(c=>[c.id,c.name,c.group,c.parentId,new Date()]));
    }
    // Archive existing operations unchanged; exclusions are ONLY from new selections.
    wlInvalidateDictionaries_();
    return {ok:true,applied:selected.length,expenses:next.expenses.length,incomes:next.incomes.length};
  });
}
