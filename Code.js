/** WL | Бюджет ДДС. Исправление переноса месяцев и массовых действий. Прежние данные сохраняются. */
const WL_PREFIX='WL_DDS_';
const WL_OPERATIONS='Операции';
const WL_TRASH='Корзина удалённых строк';
const WL_MONTHS='Периоды';
const WL_DICT='Справочники';
const WL_ACCESS='Настройки доступа';
const WL_KEYS=['departments','payments','sources','categories','incomeCategories'];
const WL_HEADERS=['Отделы','Как платим','Откуда платим','Статьи расходов ДДС','Статьи поступлений ДДС'];
const WL_OLD_SHEETS=['Отделы','Как платим','Откуда платим','Статьи ДДС'];
const WL_MIN_MONTH='2026-11',WL_MAX_INIT_MONTH='2027-12';
const WL_ALLOWED_FIELDS=['payment','source','category','description'];
const WL_TAB_NAMES=['Отделы','ЦФО','БДДС','Справочники','ФинТабло','Настройки','Инструкция'];
const WL_USER_DEFAULT_TABS=['Отделы','ЦФО','Справочники','Инструкция'];
const WL_ACCESS_COLUMNS=['Email','Роль','Видит отделы (через |)','Изменяет отделы (через |)','Активен','Видит вкладки (через |)','Изменяет вкладки (через |)','ФИО'];

function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('WL | Бюджет ДДС').addMetaTag('viewport','width=device-width,initial-scale=1');}

function wlStore_(){
  const props=PropertiesService.getScriptProperties();
  const id=props.getProperty(WL_PREFIX+'SPREADSHEET_ID');
  if(id)return SpreadsheetApp.openById(id);
  const ss=SpreadsheetApp.create('WL | Бюджет ДДС — данные приложения');
  props.setProperty(WL_PREFIX+'SPREADSHEET_ID',ss.getId());
  return ss;
}
function wlNewSheet_(ss,name,headers){
  let sh=ss.getSheetByName(name);
  if(!sh)sh=ss.insertSheet(name);
  if(sh.getLastRow()===0){
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#e7eef7');
    sh.setFrozenRows(1);
    for(let c=1;c<=headers.length;c++)sh.setColumnWidth(c,c===1?260:205);
  }
  return sh;
}
function wlOps_(ss){return wlNewSheet_(ss,WL_OPERATIONS,['Месяц','Изменено','Данные строки (JSON)']);}
function wlTrashSheet_(ss){return wlNewSheet_(ss,WL_TRASH,['Месяц','Удалено','Кто удалил','Данные строки (JSON)','Статус','Восстановлено','Кто восстановил']);}
function wlAccessSheet_(ss){
  const sh=wlNewSheet_(ss,WL_ACCESS,WL_ACCESS_COLUMNS);
  // Совместимость с v4: добавляем только заголовки, не меняем существующие права.
  if(sh.getLastColumn()<7)sh.getRange(1,6,1,2).setValues([WL_ACCESS_COLUMNS.slice(5,7)]);
  if(sh.getLastColumn()<8)sh.getRange(1,8).setValue('ФИО');
  return sh;
}
function wlMonthSheet_(ss){return wlNewSheet_(ss,WL_MONTHS,['Месяц','Статус','Показывать','Создан','Скопирован из','Изменён']);}
function wlDictSheet_(ss){
  let sh=ss.getSheetByName(WL_DICT);
  if(sh){if(sh.getLastColumn()<5)sh.getRange(1,5).setValue('Статьи поступлений ДДС');return sh;}
  sh=wlNewSheet_(ss,WL_DICT,WL_HEADERS);
  // Берём существующие дополнения пользователя из листов версии 3, затем добираем справочник из Excel.
  const seeds=wlSeedDictionaries_();
  const vals=WL_KEYS.map((key,i)=>{
    let result=[];
    const old=WL_OLD_SHEETS[i]?ss.getSheetByName(WL_OLD_SHEETS[i]):null;
    if(old&&old.getLastRow()>1)result=old.getRange(2,1,old.getLastRow()-1,1).getDisplayValues().map(r=>r[0]);
    result=result.concat(seeds[key]||[]).map(v=>String(v||'').trim());
    return [...new Set(result.filter(Boolean).map(v=>key==='departments'&&v==='Отдел запусков'?'Запуски':v))];
  });
  const len=Math.max(...vals.map(v=>v.length),0);
  if(len)sh.getRange(2,1,len,WL_KEYS.length).setValues(Array.from({length:len},(_,r)=>vals.map(v=>v[r]||'')));
  return sh;
}
/** A single cache for non-secret reference lists; permissions are NEVER cached. */
function wlDictionaries_(ss){
  const cache=CacheService.getScriptCache(),cacheKey=WL_PREFIX+'DICTIONARIES_V6';
  const stored=cache.get(cacheKey);
  if(stored){try{return JSON.parse(stored);}catch(e){cache.remove(cacheKey);}}
  const sh=wlDictSheet_(ss),n=sh.getLastRow();
  const arr=n>1?sh.getRange(2,1,n-1,WL_KEYS.length).getDisplayValues():[];
  const d={};WL_KEYS.forEach((key,i)=>d[key]=[...new Set(arr.map(row=>String(row[i]||'').trim()).filter(Boolean))]);
  d.departments=d.departments.map(v=>({name:v}));
  const json=JSON.stringify(d);if(json.length<90000)cache.put(cacheKey,json,300);
  return d;
}
function wlInvalidateDictionaries_(){CacheService.getScriptCache().remove(WL_PREFIX+'DICTIONARIES_V6');}
function wlKey_(k){if(typeof k!=='string'||!/^\d{4}-(0[1-9]|1[0-2])$/.test(k))throw Error('Неверный месяц');return k;}
function wlNormalizeDept_(d){return String(d||'').trim()==='Отдел запусков'?'Запуски':String(d||'').trim();}
/** В Google Sheets значение «2026-11» может автоматически превратиться в дату.
 * Считываем месяц независимо от типа ячейки, не трогая сами данные. */
function wlMonthCellKey_(value, tz){
  if(value instanceof Date && !isNaN(value.getTime()))
    return Utilities.formatDate(value,tz||Session.getScriptTimeZone(),'yyyy-MM');
  const s=String(value===null||value===undefined?'':value).trim();
  let m=s.match(/^(\d{4})[-./](0?[1-9]|1[0-2])(?:[-./]\d{1,2})?$/);
  if(m)return m[1]+'-'+('0'+m[2]).slice(-2);
  m=s.match(/^(0?[1-9]|1[0-2])[./-](\d{4})$/);
  if(m)return m[2]+'-'+('0'+m[1]).slice(-2);
  m=s.match(/^\d{1,2}[./-](0?[1-9]|1[0-2])[./-](\d{4})$/);
  if(m)return m[2]+'-'+('0'+m[1]).slice(-2);
  return '';
}
function wlMonthList_(ss){
  const closedYears=JSON.parse(PropertiesService.getScriptProperties().getProperty('WL_DDS_CLOSED_YEARS')||'[]');
  const sh=wlMonthSheet_(ss),v=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,6).getValues():[];
  const zone=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone();
  return v.map((r,i)=>({row:i+2,key:wlMonthCellKey_(r[0],zone),status:closedYears.includes(String(wlMonthCellKey_(r[0],zone)).slice(0,4))?'Закрыт':String(r[1]||'Открыт'),active:r[2]!==false&&String(r[2]).toLowerCase()!=='false',clonedFrom:String(r[4]||'')})).filter(m=>/^\d{4}-(0[1-9]|1[0-2])$/.test(m.key)).sort((a,b)=>a.key.localeCompare(b.key));
}
function wlFindMonth_(ss,key,includeHidden){const m=wlMonthList_(ss).find(x=>x.key===wlKey_(key));if(!m||(!includeHidden&&!m.active))throw Error('Месяц отсутствует в списке');return m;}
function wlEnsureDefaultMonths_(ss){
  const sh=wlMonthSheet_(ss);
  // Если месяцы существуют, включая закрытые и скрытые, не меняем настройки.
  if(wlMonthList_(ss).length)return;
  const vals=[];
  for(let y=2026;y<=2027;y++)for(let m=1;m<=12;m++){
    const k=y+'-'+('0'+m).slice(-2);if(k<WL_MIN_MONTH||k>WL_MAX_INIT_MONTH)continue;
    vals.push([k,'Открыт',true,new Date(),'','']);
  }
  if(vals.length){
    const first=sh.getLastRow()+1;
    sh.getRange(first,1,vals.length,1).setNumberFormat('@');
    sh.getRange(first,1,vals.length,6).setValues(vals);
  }
}
function wlCurrentEmail_(){return String(Session.getActiveUser().getEmail()||'').trim().toLowerCase();}
function wlRules_(ss){
  const sh=wlAccessSheet_(ss);if(sh.getLastRow()<2)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,8).getDisplayValues().map((r,i)=>{
    if(!r[0])return null;
    const rawRole=String(r[1]).trim();const role=rawRole==='admin'?'admin':'user'; const legacyTreasurer=rawRole==='treasurer';
    const viewTabs=String(r[5]||'').split('|').map(s=>s.trim()==='Свод'?'ЦФО':s.trim()).filter(Boolean);
    const editTabs=String(r[6]||'').split('|').map(s=>s.trim()==='Свод'?'ЦФО':s.trim()).filter(Boolean);
    return {row:i+2,email:String(r[0]).trim().toLowerCase(),name:String(r[7]||'').trim(),role,legacyTreasurer,
      view:String(r[2]||'').split('|').map(s=>s.trim()).filter(Boolean),
      edit:legacyTreasurer?[]:String(r[3]||'').split('|').map(s=>s.trim()).filter(Boolean),
      enabled:String(r[4]).toLowerCase()!=='false'&&String(r[4]).toLowerCase()!=='нет',
      tabs:viewTabs.length?viewTabs:WL_USER_DEFAULT_TABS.slice(),
      tabEdit:legacyTreasurer?[]:(editTabs.length?editTabs:(viewTabs.length?[]:['Отделы']))};
  }).filter(Boolean);
}
function wlUser_(ss){
  const email=wlCurrentEmail_();
  if(!email)throw Error('Google не передал email посетителя. Доступ закрыт. Для индивидуальных прав нужен вход Google с определяемым email (обычно Google Workspace / выполнение от пользователя).');
  const rule=wlRules_(ss).find(r=>r.email===email&&r.enabled);
  if(!rule)throw Error('Для '+email+' доступ ещё не назначен администратором.');
  const all=wlDictionaries_(ss).departments.map(d=>d.name);
  const admin=rule.role==='admin';
  const legacyReadOnly=!!rule.legacyTreasurer;
  const tabs=admin?WL_TAB_NAMES.slice():rule.tabs.filter(t=>WL_TAB_NAMES.includes(t)&&t!=='Настройки'&&t!=='ФинТабло');
  const tabEdit=admin?WL_TAB_NAMES.slice():rule.tabEdit.filter(t=>t==='Отделы'&&tabs.includes(t));
  return {email,name:rule.name||email,admin,view:admin?all:(legacyReadOnly&&!rule.view.length?all:rule.view.filter(v=>all.includes(v))),
    edit:admin?all:(legacyReadOnly?[]:rule.edit.filter(v=>all.includes(v))),tabs,tabEdit:legacyReadOnly?[]:tabEdit};
}
/** Employees see only current/future calendar years; administrators retain historic access. */
function wlCanSeeBudgetYear_(user,key){
  if(user.admin)return true;
  const currentYear=Number(Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy'));
  return Number(String(key).slice(0,4))>=currentYear;
}
function wlRequireBudgetYear_(user,key){
  if(!wlCanSeeBudgetYear_(user,key))throw Error('Прошлый финансовый год перенесён в архив и доступен только администратору');
}
function wlCanSeeTab_(u,tab){return u.tabs.includes(tab);}
function wlCanEditTab_(u,tab){return u.tabEdit.includes(tab);}
function wlRequireTab_(u,tab,edit){if(!(edit?wlCanEditTab_(u,tab):wlCanSeeTab_(u,tab)))throw Error('Нет права '+(edit?'изменять':'открывать')+' раздел «'+tab+'»');}
function wlAdmin_(ss){const u=wlUser_(ss);if(!u.admin)throw Error('Действие доступно только администратору');return u;}
function wlLock_(fn){const l=LockService.getScriptLock();l.waitLock(30000);try{return fn();}finally{l.releaseLock();}}
/** For interactive saves: bounded lock wait. The browser automatically retries SERVER_BUSY_RETRY. */
function wlSaveLock_(fn){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(4500))throw Error('SERVER_BUSY_RETRY');
  try{return fn();}finally{lock.releaseLock();}
}
/** Reconcile month templates once, not on every tab/department change.
    Any edit to a template field makes subsequent months dirty immediately. */
function wlSkeletonMark_(month){return WL_PREFIX+'SKELETON_OPT_20261107_'+month;}
function wlSkeletonSourceKey_(month){return WL_PREFIX+'SKELETON_SOURCE_20261107_'+month;}
/** One durable revision write instead of clearing a dozen month markers on every edit. */
function wlInvalidateFollowingMonths_(ss,key,includeCurrent){
  const props=PropertiesService.getScriptProperties(),name=wlSkeletonSourceKey_(key);
  const next=Math.max(Date.now()+1,Number(props.getProperty(name)||0)+1);
  props.setProperty(name,String(next));
  if(includeCurrent)props.deleteProperty(wlSkeletonMark_(key));
}
function wlSkeletonIsCurrent_(props,key){
  const all=props.getProperties(),stamp=Number(all[wlSkeletonMark_(key)]||0);
  if(!stamp||Date.now()-stamp>=24*60*60*1000)return false;
  const prefix=WL_PREFIX+'SKELETON_SOURCE_20261107_';
  return !Object.keys(all).some(name=>name.startsWith(prefix)&&name.slice(prefix.length)<key&&Number(all[name])>stamp);
}
/** Index IDs to sheet row numbers, with verification before every edit. */
function wlIndexCacheKey_(key){return WL_PREFIX+'OPS_INDEX_OPT_'+key;}
function wlInvalidateOpsIndex_(key){CacheService.getScriptCache().remove(wlIndexCacheKey_(key));}
function wlReadMonthIndex_(ops,key){
  const c=CacheService.getScriptCache(),ck=wlIndexCacheKey_(key),data=c.get(ck);
  if(data){try{const parsed=JSON.parse(data);if(parsed&&typeof parsed==='object')return parsed;}catch(e){}}
  const index=Object.create(null);
  wlOpsRead_(ops,key).forEach(e=>{index[String(e.row.id)]=e.line;});
  wlStoreMonthIndex_(key,index);
  return index;
}
function wlStoreMonthIndex_(key,index){
  const json=JSON.stringify(index),cache=CacheService.getScriptCache(),ck=wlIndexCacheKey_(key);
  if(json.length<90000)cache.put(ck,json,300);else cache.remove(ck);
}
function wlReadIndexedRow_(ops,key,id,index){
  const line=Object.prototype.hasOwnProperty.call(index,id)?Number(index[id]):0;
  if(!line)return null;
  const value=ops.getRange(line,1,1,3).getValues()[0];
  try{
    const found=JSON.parse(value[2]);
    if((wlMonthCellKey_(value[0])||String(value[0]||''))!==key||String(found.id)!==id)throw Error('index moved');
    found.department=wlNormalizeDept_(found.department);
    found.amounts=(found.amounts&&typeof found.amounts==='object')?found.amounts:{};
    found.notes=(found.notes&&typeof found.notes==='object'&&!Array.isArray(found.notes))?found.notes:{};
    return {line,key,row:found};
  }catch(e){return false;} // stale index: caller rebuilds it once
}
/** Never recreate a row that was removed by another employee or an administrator. */
function wlDeletedIdsForMonth_(ss,key){
  const sh=wlTrashSheet_(ss),n=sh.getLastRow(),out=new Set();
  if(n<2)return out;
  const items=sh.getRange(2,1,n-1,5).getValues();
  items.forEach(v=>{
    if((wlMonthCellKey_(v[0])||String(v[0]||''))!==key||String(v[4])!=='Удалена')return;
    try{const row=JSON.parse(v[3]);if(row&&row.id)out.add(String(row.id));}catch(e){}
  });
  return out;
}
/** Merge only neighboring edited sheet rows into a single Sheets write. */
function wlWriteExistingRowChanges_(ops,entries,stamp){
  const sorted=entries.filter(e=>e.line).sort((a,b)=>a.line-b.line);
  for(let i=0;i<sorted.length;){
    const start=i;while(i+1<sorted.length&&sorted[i+1].line===sorted[i].line+1)i++;
    const group=sorted.slice(start,i+1);
    ops.getRange(group[0].line,2,group.length,2).setValues(group.map(e=>[stamp,JSON.stringify(e.row)]));
    i++;
  }
}
function wlOpsRead_(sh,key){
  const v=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,3).getValues():[];
  return v.map((r,i)=>{
    const month=wlMonthCellKey_(r[0])||String(r[0]||'');
    if(key&&month!==key)return null;
    try{const o=JSON.parse(r[2]);if(!o||!o.id)return null;
      o.department=wlNormalizeDept_(o.department);
      o.amounts=(o.amounts&&typeof o.amounts==='object')?o.amounts:{};
      o.notes=(o.notes&&typeof o.notes==='object'&&!Array.isArray(o.notes))?o.notes:{};
      return {line:i+2,key:month,row:o};
    }catch(e){return null;}
  }).filter(Boolean);
}
/** Уникализируем старые/дублированные ID только в рамках конкретного месяца.
 * Суммы, статьи, описания и данные других месяцев не меняются. */
function wlRepairRowIds_(ops, key) {
  const n=ops.getLastRow();if(n<2)return 0;
  const vals=ops.getRange(2,1,n-1,3).getValues(),seen=new Set(),changes=[];
  vals.forEach((v,i)=>{
    if(wlMonthCellKey_(v[0])!==key)return;
    let data;try{data=JSON.parse(v[2]);}catch(_){return;}
    if(!data||typeof data!=='object'||Array.isArray(data))return;
    let id=String(data.id||'').trim();
    if(!id||seen.has(id)){
      const oldId=id;
      do{id='repair-'+Utilities.getUuid().replace(/-/g,'');}while(seen.has(id));
      data.id=id;
      if(data.templateRoot===oldId)data.templateRoot=id;
      changes.push({line:i+2,json:JSON.stringify(data)});
    }
    seen.add(id);
  });
  const stamp=new Date();
  changes.forEach(x=>ops.getRange(x.line,2,1,2).setValues([[stamp,x.json]]));
  if(changes.length)wlInvalidateOpsIndex_(key);
  return changes.length;
}
function wlTemplateRoot_(row){return String(row.templateRoot||row.id||'').trim();}
function wlTemplateFields_(r){return {department:wlNormalizeDept_(r.department),payment:String(r.payment||''),source:String(r.source||''),category:String(r.category||''),description:String(r.description||''),cashFlowType:String(r.cashFlowType||'')};}
function wlTemplateSignature_(r){const f=wlTemplateFields_(r);return JSON.stringify([f.department,f.payment,f.source,f.category,f.description]);}
/** Синхронизирует недостающие строки из предыдущего месяца даже после импорта.
 * Уже существующие суммы текущего месяца, индивидуальные правки и удалённые строки сохраняются.
 * При первом обращении к дальнему месяцу последовательно достраиваются промежуточные месяцы. */
function wlEnsureSkeleton_(ss,key){
  const props=PropertiesService.getScriptProperties();
  // Once a day re-check external sheet edits, or immediately after source-month edits.
  if(wlSkeletonIsCurrent_(props,key))return;
  return wlLock_(()=>{
    if(wlSkeletonIsCurrent_(props,key))return;
    const periods=wlMonthList_(ss).filter(x=>x.key<=key),ops=wlOps_(ss);
    const syncedMonths={};
    const trash=wlTrashSheet_(ss),tn=trash.getLastRow();
    const deleted=new Map();
    if(tn>1)trash.getRange(2,1,tn-1,5).getValues().forEach(v=>{
      if(String(v[4])!=='Удалена')return;
      const m=wlMonthCellKey_(v[0]);if(!m||m>key)return;
      try{const r=JSON.parse(v[3]);if(!r)return;
        if(!deleted.has(m))deleted.set(m,new Set());
        deleted.get(m).add(wlTemplateRoot_(r));
      }catch(_){}
    });
    let previousKey='',previous=[];const cancelledRoots=new Set();
    for(const period of periods){
      for(const root of (deleted.get(period.key)||[]))cancelledRoots.add(root);
      if(period.status==='Открыт')wlRepairRowIds_(ops,period.key);
      const current=wlOpsRead_(ops,period.key);
      // Закрытый месяц остаётся полностью неизменным, в том числе при автопереносе.
      if(period.status==='Открыт'&&previousKey&&previous.length){
        const existingRoots=new Set(current.map(e=>wlTemplateRoot_(e.row)));
        const blocked=deleted.get(period.key)||new Set();
        const oldCopies=new Map();
        current.forEach(e=>{
          if(e.row.templateRoot||!String(e.row.id).startsWith('copy-'+period.key+'-'))return;
          const sig=wlTemplateSignature_(e.row);
          if(!oldCopies.has(sig))oldCopies.set(sig,[]);
          oldCopies.get(sig).push(e);
        });
        const added=[],now=new Date();
        previous.forEach(src=>{
          const root=wlTemplateRoot_(src),fields=wlTemplateFields_(src);
          if(!root||blocked.has(root)||cancelledRoots.has(root))return;
          if(existingRoots.has(root)){
            // Для новых шаблонных строк подтягиваем правки источника только
            // в полях, которые в целевом месяце ещё не изменяли вручную.
            const ent=current.find(e=>wlTemplateRoot_(e.row)===root&&e.row.templateBaseline);
            if(ent){
              const base=ent.row.templateBaseline;let changed=false;
              for(const f of ['payment','source','category','description','cashFlowType']){
                const before=String(base[f]||''),value=String(ent.row[f]||''),next=String(fields[f]||'');
                if(value===before&&next!==before){ent.row[f]=next;changed=true;}
                if(next!==before){base[f]=next;changed=true;}
              }
              if(changed)ops.getRange(ent.line,2,1,2).setValues([[now,JSON.stringify(ent.row)]]);
            }
            return;
          }
          // Переподключаем старые строки, скопированные предыдущей версией,
          // только при полном совпадении реквизитов. Суммы не меняются.
          const sig=wlTemplateSignature_(src),candidates=oldCopies.get(sig)||[];
          const old=candidates.shift();
          if(old){
            old.row.templateRoot=root;
            // Legacy copies have no reliable history of manual edits; never auto-delete them.
            old.row.templateLegacyUnverified=true;
            old.row.templateFrom=previousKey;
            old.row.templateBaseline=Object.assign({},fields);
            ops.getRange(old.line,2,1,2).setValues([[now,JSON.stringify(old.row)]]);
            existingRoots.add(root);
            return;
          }
          const row={id:'carry-'+period.key+'-'+Utilities.getUuid().replace(/-/g,''),
            templateRoot:root,templateFrom:previousKey,templateBaseline:Object.assign({},fields),
            ...(src.templateLegacyUnverified?{templateLegacyUnverified:true}:{}),
            department:fields.department,payment:fields.payment,source:fields.source,
            category:fields.category,description:fields.description,cashFlowType:fields.cashFlowType,amounts:{}};
          added.push([period.key,now,JSON.stringify(row)]);existingRoots.add(root);
        });
        if(added.length){
          const first=ops.getLastRow()+1;
          ops.getRange(first,1,added.length,1).setNumberFormat('@');
          ops.getRange(first,1,added.length,3).setValues(added);
          wlInvalidateOpsIndex_(period.key);
        }
        if(added.length||!period.clonedFrom){
          wlMonthSheet_(ss).getRange(period.row,5,1,2).setValues([[previousKey,now]]);
        }
      }
      previousKey=period.key;
      previous=wlOpsRead_(ops,period.key).map(e=>e.row);
      syncedMonths[wlSkeletonMark_(period.key)]=String(Date.now());
    }
    if(Object.keys(syncedMonths).length)props.setProperties(syncedMonths,false);
  });
}
/** Выполняется владельцем ОДИН РАЗ из редактора Apps Script. */
function wlInitializeV4(){
  return wlLock_(()=>{
    const ss=wlStore_();wlOps_(ss);wlDictSheet_(ss);wlEnsureDefaultMonths_(ss);
    const sh=wlAccessSheet_(ss),rules=wlRules_(ss);
    if(!rules.length){
      const email=wlCurrentEmail_();
      const effective=String(Session.getEffectiveUser().getEmail()||'').toLowerCase();
      if(!email||email!==effective)throw Error('Первичную настройку должен запустить владелец из редактора Apps Script');
      sh.appendRow([email,'admin','Все','Все',true]);
      PropertiesService.getScriptProperties().setProperty(WL_PREFIX+'OWNER_EMAIL',email);
    }
    return {ok:true,spreadsheetUrl:ss.getUrl(),admin:wlCurrentEmail_()};
  });
}
function wlBoot(){
  const start=Date.now(),ss=wlStore_(),user=wlUser_(ss);wlEnsureDefaultMonths_(ss);
  const dict=wlDictionaries_(ss);dict.departments=dict.departments.filter(x=>user.view.includes(x.name));
  // Incoming articles may be read for summaries; only administrators can edit or select them.
  const active=wlMonthList_(ss).filter(m=>m.active&&wlCanSeeBudgetYear_(user,m.key));
  if(!active.length)throw Error('В таблице «Периоды» нет открытых для выбора месяцев. Проверьте столбец «Показывать».');
  console.log('WL_PERF boot '+(Date.now()-start)+'ms');
  const nowKey=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM');
  const current=active.some(m=>m.key===nowKey)?nowKey:(active.filter(m=>m.key<=nowKey).slice(-1)[0]||active[0]).key;
  const props=PropertiesService.getScriptProperties(),archivedYears=[...new Set(active.map(m=>m.key.slice(0,4)))].filter(y=>props.getProperty('WL_DDS_BDDS_ARCHIVED_'+y)==='true');
  return {user,dict,months:active.map(m=>({key:m.key,status:m.status})),current,archivedYears,grouping:wlBddsGrouping_(ss)};
}
function wlRowOrder_(ss,key,department,rows){
 const sh=ss.getSheetByName('WL_Порядок_строк');if(!sh||sh.getLastRow()<2)return rows;
 const zone=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone(),records=sh.getRange(2,1,sh.getLastRow()-1,3).getValues(),found=records.find(r=>wlMonthCellKey_(r[0],zone)===key&&String(r[1])===department);let ids=[];if(found)try{ids=JSON.parse(String(found[2]));}catch(_){}if(!Array.isArray(ids))return rows;
 const rank=new Map(ids.map((id,i)=>[String(id),i]));return rows.map((row,i)=>({row,i})).sort((a,b)=>(rank.has(a.row.id)?rank.get(a.row.id):ids.length+a.i)-(rank.has(b.row.id)?rank.get(b.row.id):ids.length+b.i)).map(x=>x.row);
}
function wlMoveBudgetRow(key,department,id,direction){
 key=wlKey_(key);department=wlNormalizeDept_(department);id=String(id||'');if(!['up','down','top','bottom'].includes(direction))throw Error('Недопустимое направление');
 return wlLock_(()=>{const ss=wlStore_(),user=wlUser_(ss),month=wlFindMonth_(ss,key);wlRequireBudgetYear_(user,key);wlRequireTab_(user,'Отделы',true);if(month.status!=='Открыт'||!user.edit.includes(department))throw Error('Нет права изменять строки этого отдела или месяц закрыт');
  const rows=wlRowOrder_(ss,key,department,wlOpsRead_(wlOps_(ss),key).map(x=>x.row).filter(r=>r.department===department)),index=rows.findIndex(r=>r.id===id);if(index<0)throw Error('Строка не найдена. Обновите страницу');
  const target=direction==='top'?0:direction==='bottom'?rows.length-1:direction==='up'?Math.max(0,index-1):Math.min(rows.length-1,index+1);const [moved]=rows.splice(index,1);rows.splice(target,0,moved);
  const sh=wlNewSheet_(ss,'WL_Порядок_строк',['Месяц','Отдел','Порядок ID','Изменено','Кто изменил']),zone=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone(),old=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,2).getValues():[],i=old.findIndex(r=>wlMonthCellKey_(r[0],zone)===key&&String(r[1])===department),line=i<0?sh.getLastRow()+1:i+2;
  sh.getRange(line,1,1,5).setValues([[key,department,JSON.stringify(rows.map(r=>r.id)),new Date(),user.email]]);return rows.map(r=>r.id);
 });
}
function wlGetMonth(key){
  const start=Date.now();
  key=wlKey_(key);const ss=wlStore_(),user=wlUser_(ss);wlRequireBudgetYear_(user,key);let month=wlFindMonth_(ss,key);
  if(!wlCanSeeTab_(user,'Отделы')&&!wlCanSeeTab_(user,'ЦФО'))throw Error('Нет доступа к детальным данным бюджета');
  wlEnsureSkeleton_(ss,key);month=wlFindMonth_(ss,key);
  const visible=wlOpsRead_(wlOps_(ss),key).map(x=>x.row).filter(r=>user.view.includes(r.department));
  const rows=user.view.flatMap(department=>wlRowOrder_(ss,key,department,visible.filter(r=>r.department===department)));
  console.log('WL_PERF getMonth '+key+' '+(Date.now()-start)+'ms');
  return {key,month:{key:month.key,status:month.status},rows};
}
function wlPatchChanges(key,updates){
  const start=Date.now();key=wlKey_(key);
  if(!Array.isArray(updates)||!updates.length||updates.length>150)throw Error('Неверный набор изменений');
  return wlSaveLock_(()=>{
    const ss=wlStore_(),user=wlUser_(ss),month=wlFindMonth_(ss,key);
    wlRequireBudgetYear_(user,key);
    wlRequireTab_(user,'Отделы',true);
    if(month.status!=='Открыт')throw Error('Месяц закрыт. Изменения запрещены');
    const dict=wlDictionaries_(ss),ops=wlOps_(ss);
    let index=wlReadMonthIndex_(ops,key);
    const byId=new Map(),touched=new Map();
    let changesTemplate=false,deletedIds=null;
    updates.forEach(item=>{
      const id=String(item.rowId||'').trim(),dept=wlNormalizeDept_(item.department),field=String(item.field||'');
      if(!id||id.length>100||!/^[A-Za-z0-9._:-]+$/.test(id))throw Error('Неверный идентификатор строки');
      if(!user.edit.includes(dept))throw Error('Нет права изменять отдел '+dept);
      if(!dict.departments.some(d=>d.name===dept))throw Error('Неизвестный отдел '+dept);
      let ent=byId.get(id);
      if(!ent){
        ent=wlReadIndexedRow_(ops,key,id,index);
        if(ent===false){
          wlInvalidateOpsIndex_(key);
          index=wlReadMonthIndex_(ops,key);
          ent=wlReadIndexedRow_(ops,key,id,index);
          if(ent===false)throw Error('Строка перемещена. Обновите таблицу и повторите действие');
        }
        if(!ent){
          if(!deletedIds)deletedIds=wlDeletedIdsForMonth_(ss,key);
          if(deletedIds.has(id))throw Error('Строка уже удалена. Обновите страницу перед повторным вводом');
          ent={line:0,key,row:{id,department:dept,payment:'',source:'',category:'',description:'',amounts:{},notes:{}}};changesTemplate=true;
        }
        byId.set(id,ent);
      }
      if(ent.row.department!==dept)throw Error('Нельзя редактировать другой отдел');
      if(field==='amount'){
        const day=Number(item.day),[y,m]=key.split('-').map(Number);
        if(!Number.isInteger(day)||day<1||day>new Date(y,m,0).getDate())throw Error('Неверная дата');
        if(item.value===''||item.value==null)delete ent.row.amounts[day];
        else{const n=Number(item.value);if(!Number.isFinite(n)||Math.abs(n)>1e12)throw Error('Недопустимая сумма');ent.row.amounts[day]=Math.round(n*100)/100;}
      }else if(field==='formula'){
        const day=Number(item.day),[y,m]=key.split('-').map(Number);
        if(!Number.isInteger(day)||day<1||day>new Date(y,m,0).getDate())throw Error('Неверная дата формулы');
        const text=String(item.value==null?'':item.value).trim();
        if(text&&(text.length>160||!/^=[0-9+\-*\/().,\s\u00a0\u202f]+$/.test(text)))throw Error('Недопустимый формат формулы');
        if(!ent.row.formulas||typeof ent.row.formulas!=='object'||Array.isArray(ent.row.formulas))ent.row.formulas={};
        if(text)ent.row.formulas[day]=text;else delete ent.row.formulas[day];
      }else if(field==='note'){
        const day=Number(item.day),[y,m]=key.split('-').map(Number);
        if(!Number.isInteger(day)||day<1||day>new Date(y,m,0).getDate())throw Error('Неверная дата заметки');
        const note=String(item.value==null?'':item.value).trim();
        if(note.length>1000)throw Error('Заметка слишком длинная (максимум 1000 символов)');
        if(!ent.row.notes||typeof ent.row.notes!=='object'||Array.isArray(ent.row.notes))ent.row.notes={};
        // Store a threaded comment with server-verified identity. Historical plain-text notes remain readable.
        const old=ent.row.notes[day];
        const messages=Array.isArray(old&&old.messages)?old.messages.slice():(typeof old==='string'&&old?[{text:old,by:'Ранее добавлено',email:'',at:''}]:[]);
        if(item.noteAction==='delete'){
          if(!user.admin&&messages.some(m=>String(m.email||'').toLowerCase()!==user.email))throw Error('Удалять чужие комментарии может только администратор');
          delete ent.row.notes[day];
        }else if(note){
          messages.push({text:note,by:user.name||user.email,email:user.email,at:new Date().toISOString()});
          ent.row.notes[day]={messages};
        }else throw Error('Пустой комментарий не добавлен');
      }else{
        if(!WL_ALLOWED_FIELDS.includes(field))throw Error('Недопустимое поле');
        if(field==='category'&&!user.admin)throw Error('Статью ДДС может изменять только администратор');
        let val=String(item.value==null?'':item.value).trim();if(val.length>(field==='description'?2000:350))throw Error('Слишком длинный текст');
        const dictKey={payment:'payments',source:'sources',category:'categories'}[field];
        // Старые, не сопоставленные значения сохраняются неизменными; новые только из справочника.
        if(dictKey&&val&&!dict[dictKey].includes(val)&&!(field==='category'&&user.admin&&dict.incomeCategories.includes(val))&&val!==String(ent.row[field]||''))throw Error('Выбери значение из справочника: '+field);
        if(field==='description'&&val!==String(ent.row.description||'')){
          // Администратор видит предупреждение, если описание изменилось или заполнено впервые.
          if(!ent.row.reviewPending){
            ent.row.reviewPrevious=String(ent.row.description||'');
            ent.row.reviewChangedAt=new Date().toISOString();
            ent.row.reviewChangedBy=user.email;
            ent.row.reviewCategory=String(ent.row.category||'');
          }
          ent.row.reviewPending=true;
        }
        if(String(ent.row[field]||'')!==val)changesTemplate=true;
        ent.row[field]=val;
        if(field==='category'){
          const requested=String(item.categoryType||'');
          if(requested&&requested!=='income'&&requested!=='expense')throw Error('Недопустимый тип статьи ДДС');
          if(requested==='income'&&!dict.incomeCategories.includes(val))throw Error('Статья отсутствует в поступлениях');
          if(requested==='expense'&&!dict.categories.includes(val))throw Error('Статья отсутствует в списаниях');
          ent.row.cashFlowType=requested|| (dict.incomeCategories.includes(val)&&!dict.categories.includes(val)?'income':'expense');
        }
      }
      touched.set(id,ent);
    });
    const edits=[...touched.values()],now=new Date(),news=[];
    wlWriteExistingRowChanges_(ops,edits,now);
    edits.forEach(e=>{if(!e.line)news.push([key,now,JSON.stringify(e.row)]);});
    if(news.length){
      const first=ops.getLastRow()+1;
      ops.getRange(first,1,news.length,1).setNumberFormat('@');
      ops.getRange(first,1,news.length,3).setValues(news);
      news.forEach((n,i)=>{const id=JSON.parse(n[2]).id;index[id]=first+i;});
      wlStoreMonthIndex_(key,index);
    }
    if(changesTemplate)wlInvalidateFollowingMonths_(ss,key,false);
    console.log('WL_PERF save '+key+' changes='+updates.length+' '+(Date.now()-start)+'ms');
    return {ok:true,count:updates.length};
  });
}
/** Preserve edited and closed future months when deleting inherited budget rows. */
function wlIsUntouchedCarry_(row){
 if(!row||!row.templateRoot||!row.templateBaseline||row.templateLegacyUnverified||String(row.id||'').startsWith('copy-'))return false;
 if(Object.values(row.amounts||{}).some(x=>Number(x)!==0))return false;
 if(Object.keys(row.notes||{}).length||Object.keys(row.formulas||{}).length)return false;
 const fields=wlTemplateFields_(row),base=row.templateBaseline;
 return ['department','payment','source','category','description','cashFlowType'].every(k=>String(fields[k]||'')===String(base[k]||''));
}
function wlCascadeDeletedTemplates_(ss,key,removed,u){
 const roots=new Set(removed.map(wlTemplateRoot_).filter(Boolean));
 if(!roots.size)return {propagated:0,keptEdited:0,keptClosed:0};
 const states=new Map(wlMonthList_(ss).filter(m=>m.key>key).map(m=>[m.key,m.status]));
 const ops=wlOps_(ss),items=wlOpsRead_(ops).filter(v=>v.key>key&&v.row.templateRoot&&roots.has(wlTemplateRoot_(v.row)));
 const drop=[],edited=[];let keptClosed=0;
 for(const item of items){
  if(states.get(item.key)!=='Открыт'){keptClosed++;continue;}
  if(wlIsUntouchedCarry_(item.row))drop.push(item);
  else {item.row.templateRoot=item.row.id;delete item.row.templateFrom;delete item.row.templateBaseline;delete item.row.templateLegacyUnverified;edited.push(item);}
 }
 const now=new Date();
 if(drop.length){const trash=wlTrashSheet_(ss),from=trash.getLastRow()+1;
  trash.getRange(from,1,drop.length,7).setValues(drop.map(item=>[item.key,now,u.email,JSON.stringify(item.row),'Удалена','','']));}
 for(const item of edited)ops.getRange(item.line,2,1,2).setValues([[now,JSON.stringify(item.row)]]);
 drop.sort((a,b)=>b.line-a.line).forEach(item=>ops.deleteRow(item.line));
 for(const period of new Set([...drop,...edited].map(item=>item.key))){wlInvalidateOpsIndex_(period);wlInvalidateFollowingMonths_(ss,period,false);}
 return {propagated:drop.length,keptEdited:edited.length,keptClosed};
}
function wlDeleteRow(key,id){
  return wlLock_(()=>{
    key=wlKey_(key);id=String(id||'').trim();if(!id)throw Error('Не указан идентификатор строки');
    const ss=wlStore_(),u=wlUser_(ss),month=wlFindMonth_(ss,key);
    wlRequireBudgetYear_(u,key);
    wlRequireTab_(u,'Отделы',true);
    if(month.status!=='Открыт')throw Error('Закрытый месяц нельзя изменять');
    const ops=wlOps_(ss),entry=wlOpsRead_(ops,key).find(v=>String(v.row.id)===id);
    // Кнопка может быть нажата до записи новой строки или повторно.
    if(!entry)return {ok:true,alreadyMissing:true};
    if(!u.edit.includes(entry.row.department))throw Error('Нет доступа');
    wlTrashSheet_(ss).appendRow([key,new Date(),u.email,JSON.stringify(entry.row),'Удалена','','']);
    ops.deleteRow(entry.line);
    const cascade=wlCascadeDeletedTemplates_(ss,key,[entry.row],u);
    wlInvalidateOpsIndex_(key);wlInvalidateFollowingMonths_(ss,key,false);
    return {ok:true,...cascade};
  });
}
/** Администратор видит удалённые строки и может восстановить любую из них в открытом месяце. */
function wlTrashList(){
  const ss=wlStore_();wlAdmin_(ss);
  const sh=wlTrashSheet_(ss),n=sh.getLastRow();if(n<2)return [];
  return sh.getRange(2,1,n-1,7).getValues().map((r,i)=>{
    let row=null;try{row=JSON.parse(r[3]);}catch(_){return null;}
    if(!row||!row.id)return null;
    return {entry:i+2,month:wlMonthCellKey_(r[0])||String(r[0]),deletedAt:r[1] instanceof Date?r[1].toISOString():String(r[1]),deletedBy:String(r[2]||''),department:String(row.department||''),category:String(row.category||''),description:String(row.description||''),amount:Math.round(Object.values(row.amounts||{}).reduce((sum,n)=>sum+(Number(n)||0),0)*100)/100,restored:String(r[4])==='Восстановлена'};
  }).filter(Boolean).reverse().slice(0,300);
}
function wlTrashRestore(entryNumber){
  return wlLock_(()=>{
    const ss=wlStore_(),admin=wlAdmin_(ss),sh=wlTrashSheet_(ss);
    const n=Number(entryNumber);
    if(!Number.isInteger(n)||n<2||n>sh.getLastRow())throw Error('Запись корзины не найдена');
    const r=sh.getRange(n,1,1,7).getValues()[0];
    if(String(r[4])==='Восстановлена')throw Error('Строка уже восстановлена');
    const key=wlKey_(String(r[0]));
    if(wlFindMonth_(ss,key,true).status!=='Открыт')throw Error('Для восстановления сначала откройте месяц');
    let row;try{row=JSON.parse(r[3]);}catch(_){throw Error('Снимок повреждён');}
    if(!row||!row.id)throw Error('Строка повреждена');
    const ops=wlOps_(ss);
    if(wlOpsRead_(ops,key).some(x=>x.row.id===row.id))throw Error('Такая строка уже есть в бюджете');
    const next=ops.getLastRow()+1;ops.getRange(next,1).setNumberFormat('@');ops.appendRow([key,new Date(),JSON.stringify(row)]);
    sh.getRange(n,5,1,3).setValues([['Восстановлена',new Date(),admin.email]]);
    wlInvalidateOpsIndex_(key);wlInvalidateFollowingMonths_(ss,key,false);
    return {ok:true,month:key};
  });
}
function wlMonthManage(action,key){
  return wlLock_(()=>{
    key=wlKey_(key);const ss=wlStore_();wlAdmin_(ss);const sh=wlMonthSheet_(ss),list=wlMonthList_(ss),p=list.find(x=>x.key===key);
    if(action==='add'){
      if(wlYearClosed_(key.slice(0,4)))throw Error('Год закрыт. Добавление месяца запрещено');
      if(p){sh.getRange(p.row,3).setValue(true);sh.getRange(p.row,6).setValue(new Date());}
      else{const row=sh.getLastRow()+1;sh.getRange(row,1).setNumberFormat('@');sh.getRange(row,1,1,6).setValues([[key,'Открыт',true,new Date(),'','']]);}
    }else{
      if(!p)throw Error('Нет такого месяца');
      if(action==='close'||action==='open'){if(action==='open'&&wlYearClosed_(key.slice(0,4)))throw Error('Год закрыт. Сначала откройте весь год в настройках');sh.getRange(p.row,2).setValue(action==='close'?'Закрыт':'Открыт');}
      else if(action==='hide')throw Error('Функция скрытия месяцев отключена');
      else throw Error('Недопустимая операция');
      sh.getRange(p.row,6).setValue(new Date());
    }
    wlInvalidateFollowingMonths_(ss,key,true);
    return wlMonthList_(ss).filter(v=>v.active).map(v=>({key:v.key,status:v.status}));
  });
}
function wlDictionaryAdd(kind,value){
  return wlLock_(()=>{
    const ss=wlStore_();wlAdmin_(ss);const i=WL_KEYS.indexOf(kind);if(i<0)throw Error('Неизвестный справочник');
    value=String(value||'').trim();if(!value||value.length>300)throw Error('Недопустимое значение');
    const dict=wlDictionaries_(ss),existing=kind==='departments'?dict.departments.map(v=>v.name):dict[kind];
    if(existing.some(v=>v.toLocaleLowerCase('ru')===value.toLocaleLowerCase('ru')))throw Error('Такая запись уже есть');
    const sh=wlDictSheet_(ss),all=sh.getLastRow()>1?sh.getRange(2,i+1,sh.getLastRow()-1,1).getDisplayValues():[];
    const gap=all.findIndex(r=>!r[0]);sh.getRange(gap<0?sh.getLastRow()+1:gap+2,i+1).setValue(value);
    wlInvalidateDictionaries_();return wlDictionaries_(ss);
  });
}

/** Rename payment sources or methods in all stored budgets and restorable rows. */
function wlDictionaryRename(kind,oldValue,newValue){
 return wlLock_(()=>{
  const ss=wlStore_(),user=wlAdmin_(ss),fields={sources:'source',payments:'payment'},field=fields[kind];
  if(!field)throw Error('Переименование этого справочника пока не поддерживается');
  oldValue=String(oldValue||'').trim();newValue=String(newValue||'').trim();
  if(!newValue||newValue.length>300||newValue.startsWith('='))throw Error('Недопустимое название');
  if(oldValue===newValue)return wlDictionaries_(ss);
  const index=WL_KEYS.indexOf(kind),dict=wlDictSheet_(ss),values=dict.getLastRow()>1?dict.getRange(2,index+1,dict.getLastRow()-1,1).getDisplayValues():[],hits=values.map((r,i)=>r[0]===oldValue?i:-1).filter(i=>i>=0);
  if(!hits.length)throw Error('Название не найдено. Обновите справочник и повторите попытку');
  if(values.some(r=>String(r[0]).toLocaleLowerCase('ru')===newValue.toLocaleLowerCase('ru')))throw Error('Такое название уже есть');
  const changes=[],periods=new Set();
  for(const [sheet,col] of [[wlOps_(ss),3],[wlTrashSheet_(ss),4]]){
   if(sheet.getLastRow()<2)continue;const data=sheet.getRange(2,1,sheet.getLastRow()-1,col).getValues();
   data.forEach((row,i)=>{if(!row[col-1])return;let item;try{item=JSON.parse(row[col-1]);}catch(e){throw Error('В данных найдена повреждённая строка. Переименование отменено');}if(item&&item[field]===oldValue){const before=String(row[col-1]);item[field]=newValue;changes.push({sheet,line:i+2,col,before,after:JSON.stringify(item)});const month=wlMonthCellKey_(row[0]);if(month)periods.add(month);}});
  }
  const audit=wlNewSheet_(ss,'WL_Переименования',['Время','Администратор','Справочник','Было','Стало','Лист','Строка','Данные до изменения']);
  const snapshot=changes.length?changes.map(c=>[new Date(),user.email,kind,oldValue,newValue,c.sheet.getName(),c.line,c.before]):[[new Date(),user.email,kind,oldValue,newValue,'', '', '']];
  if(audit.getMaxRows()<audit.getLastRow()+snapshot.length)audit.insertRowsAfter(audit.getMaxRows(),audit.getLastRow()+snapshot.length-audit.getMaxRows());
  audit.getRange(audit.getLastRow()+1,1,snapshot.length,8).setValues(snapshot);
  changes.forEach(c=>c.sheet.getRange(c.line,c.col).setValue(c.after));hits.forEach(i=>dict.getRange(i+2,index+1).setValue(newValue));
  periods.forEach(month=>{wlInvalidateOpsIndex_(month);wlInvalidateFollowingMonths_(ss,month,false);});wlInvalidateDictionaries_();return wlDictionaries_(ss);
 });
}
function wlDictionaryRemove(kind,value){
  return wlLock_(()=>{
    const ss=wlStore_();wlAdmin_(ss);const i=WL_KEYS.indexOf(kind);if(i<0)throw Error('Неизвестный справочник');
    const sh=wlDictSheet_(ss),v=sh.getLastRow()>1?sh.getRange(2,i+1,sh.getLastRow()-1,1).getDisplayValues():[];
    const pos=v.findIndex(r=>r[0]===value);if(pos<0)throw Error('Не найдено');
    const field={departments:'department',payments:'payment',sources:'source',categories:'category',incomeCategories:'category'}[kind];
    if(wlOpsRead_(wlOps_(ss)).some(x=>x.row[field]===value))throw Error('Пункт используется в бюджете. Сначала замените его в операциях');
    sh.getRange(pos+2,i+1).clearContent();wlInvalidateDictionaries_();return wlDictionaries_(ss);
  });
}
function wlGetSettings(){const ss=wlStore_();wlAdmin_(ss);return {rules:wlRules_(ss).map(r=>({email:r.email,name:r.name||'',role:r.role,view:r.view,edit:r.edit,tabs:r.role==='admin'?WL_TAB_NAMES:r.tabs,tabEdit:r.role==='admin'?WL_TAB_NAMES:r.tabEdit,enabled:r.enabled})),departments:wlDictionaries_(ss).departments.map(d=>d.name),tabs:WL_TAB_NAMES,spreadsheetUrl:ss.getUrl()};}
function wlSaveAccess(item){
  return wlLock_(()=>{
    const ss=wlStore_();wlAdmin_(ss);
    const email=String(item.email||'').trim().toLowerCase();if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw Error('Неверный email');
    const name=String(item.name||'').trim().replace(/\s+/g,' ');if(name.length>150)throw Error('ФИО слишком длинное');
    const owner=PropertiesService.getScriptProperties().getProperty(WL_PREFIX+'OWNER_EMAIL');
    if(email===owner&&item.role!=='admin')throw Error('Нельзя лишить владельца прав администратора');
    const deps=wlDictionaries_(ss).departments.map(d=>d.name);
    const role=item.role==='admin'?'admin':'user';
    const view=role==='admin'?deps:(Array.isArray(item.view)?item.view:[]).filter(d=>deps.includes(d));
    const edit=role==='admin'?deps:(Array.isArray(item.edit)?item.edit:[]).filter(d=>view.includes(d));
    const enabled=!!item.enabled;
    const tabs=role==='admin'?WL_TAB_NAMES:(Array.isArray(item.tabs)?item.tabs:WL_USER_DEFAULT_TABS).filter(t=>WL_TAB_NAMES.includes(t)&&t!=='Настройки'&&t!=='ФинТабло');
    const tabEdit=role==='admin'?WL_TAB_NAMES:(Array.isArray(item.tabEdit)?item.tabEdit:[]).filter(t=>t==='Отделы'&&tabs.includes(t));
    if(!tabs.length)throw Error('Выберите хотя бы одну доступную вкладку');
    if(email===owner&&!enabled)throw Error('Нельзя отключить учётную запись владельца');
    const sh=wlAccessSheet_(ss),rule=wlRules_(ss).find(r=>r.email===email);
    const row=[email,role,view.join('|'),edit.join('|'),enabled,tabs.join('|'),tabEdit.join('|'),name||((rule&&rule.name)||'')];
    if(rule)sh.getRange(rule.row,1,1,8).setValues([row]);else sh.appendRow(row);
    return wlGetSettings();
  });
}
function wlRemoveAccess(email){
  return wlLock_(()=>{
    const ss=wlStore_();wlAdmin_(ss);email=String(email||'').toLowerCase();
    if(email===PropertiesService.getScriptProperties().getProperty(WL_PREFIX+'OWNER_EMAIL'))throw Error('Владельца удалять нельзя');
    const rule=wlRules_(ss).find(r=>r.email===email);if(!rule)throw Error('Не найден');wlAccessSheet_(ss).deleteRow(rule.row);return wlGetSettings();
  });
}
/** Импорт повторяемый: существующие строки по ID не трогаются. Ручные строки тоже сохраняются. */

/** Проверки изменённых описаний: уведомления сохраняются между сеансами. */
function wlReviewList(){
  const ss=wlStore_();wlAdmin_(ss);
  const closed=new Set(wlMonthList_(ss).filter(m=>m.status==='Закрыт').map(m=>m.key));
  return wlOpsRead_(wlOps_(ss)).filter(e=>e.row.reviewPending).map(e=>({
    key:e.key,id:String(e.row.id),department:e.row.department,
    category:String(e.row.category||''),oldCategory:String(e.row.reviewCategory||''),
    description:String(e.row.description||''),previous:String(e.row.reviewPrevious||''),
    by:String(e.row.reviewChangedBy||''),at:String(e.row.reviewChangedAt||''),
    monthClosed:closed.has(e.key),cashFlowType:String(e.row.cashFlowType||'')
  })).sort((a,b)=>b.at.localeCompare(a.at)).slice(0,500);
}
function wlReviewResolve(key,id,category,categoryType){
  return wlLock_(()=>{
    key=wlKey_(key);id=String(id||'').trim();category=String(category||'').trim();
    const ss=wlStore_(),admin=wlAdmin_(ss),month=wlFindMonth_(ss,key,true);
    if(month.status!=='Открыт')throw Error('Сначала откройте месяц, чтобы подтвердить статью ДДС');
    const choices=wlDictionaries_(ss);if(!choices.categories.includes(category)&&!choices.incomeCategories.includes(category))throw Error('Выберите статью из утверждённого справочника');
    const ops=wlOps_(ss),entry=wlOpsRead_(ops,key).find(e=>String(e.row.id)===id);
    if(!entry)throw Error('Строка не найдена — обновите список');
    categoryType=String(categoryType||'');
    if(categoryType&&categoryType!=='income'&&categoryType!=='expense')throw Error('Недопустимый тип статьи');
    if(categoryType==='income'&&!choices.incomeCategories.includes(category))throw Error('Такой статьи нет в поступлениях');
    if(categoryType==='expense'&&!choices.categories.includes(category))throw Error('Такой статьи нет в списаниях');
    entry.row.category=category;entry.row.cashFlowType=categoryType||(choices.incomeCategories.includes(category)&&!choices.categories.includes(category)?'income':'expense');entry.row.reviewPending=false;
    entry.row.reviewResolvedAt=new Date().toISOString();entry.row.reviewResolvedBy=admin.email;
    ops.getRange(entry.line,2,1,2).setValues([[new Date(),JSON.stringify(entry.row)]]);
    return {ok:true,month:key};
  });
}

/** Проверка месяцев после исправления — безопасно, без изменения операций. */
function wlCheckMonthsV41(){
  const ss=wlStore_();
  wlEnsureDefaultMonths_(ss);
  const months=wlMonthList_(ss);
  return {ok:months.length>0,spreadsheetUrl:ss.getUrl(),count:months.length,active:months.filter(m=>m.active).length,first:months.length?months[0].key:'',last:months.length?months[months.length-1].key:''};
}


/** Массовое удаление только выбранных уникальных строк, с записью в корзину. */
function wlDeleteRows(key, ids) {
  key=wlKey_(key);
  if(!Array.isArray(ids)||!ids.length||ids.length>400)throw Error('Отметьте от 1 до 400 строк');
  const wanted=new Set(ids.map(x=>String(x||'').trim()).filter(Boolean));
  if(wanted.size!==ids.length)throw Error('Список строк устарел. Обновите страницу бюджета и повторите выбор');
  return wlLock_(()=>{
    const ss=wlStore_(),u=wlUser_(ss),month=wlFindMonth_(ss,key);
    wlRequireBudgetYear_(u,key);
    wlRequireTab_(u,'Отделы',true);
    if(month.status!=='Открыт')throw Error('Месяц закрыт');
    const ops=wlOps_(ss),repaired=wlRepairRowIds_(ops,key);
    if(repaired)throw Error('Исправлены повторяющиеся идентификаторы. Обновите страницу и отметьте строки снова');
    const entries=wlOpsRead_(ops,key).filter(v=>wanted.has(String(v.row.id)));
    if(entries.length!==wanted.size)throw Error('Часть строк была изменена или удалена другим пользователем. Обновите страницу');
    entries.forEach(e=>{if(!u.edit.includes(e.row.department))throw Error('Нет права удалять строки отдела '+e.row.department);});
    // Отдельные записи корзины сохраняем до фактического удаления из операций.
    const trash=wlTrashSheet_(ss),now=new Date();
    const snapshots=entries.map(e=>[key,now,u.email,JSON.stringify(e.row),'Удалена','','']);
    trash.getRange(trash.getLastRow()+1,1,snapshots.length,7).setValues(snapshots);
    const nums=entries.map(e=>e.line).sort((a,b)=>a-b),runs=[];
    for(let i=0;i<nums.length;i++){
      const last=runs[runs.length-1];
      if(last&&nums[i]===last.start+last.count)last.count++;
      else runs.push({start:nums[i],count:1});
    }
    runs.reverse().forEach(r=>ops.deleteRows(r.start,r.count));
    const cascade=wlCascadeDeletedTemplates_(ss,key,entries.map(e=>e.row),u);
    wlInvalidateOpsIndex_(key);
    wlInvalidateFollowingMonths_(ss,key,false);
    return {ok:true,deleted:entries.length,missing:0,...cascade};
  });
}

/** v4.3: восстановление нескольких строк из корзины, с проверкой открытых месяцев. */
function wlTrashRestoreMany(entries){
  if(!Array.isArray(entries)||!entries.length||entries.length>300)throw Error('Выберите до 300 строк');
  return wlLock_(()=>{
    const ss=wlStore_(),u=wlAdmin_(ss),trash=wlTrashSheet_(ss),ops=wlOps_(ss);
    const uniq=[...new Set(entries.map(Number))];
    const activeIds=new Set(wlOpsRead_(ops).map(x=>x.key+'|'+x.row.id));
    const records=[];
    uniq.forEach(entry=>{
      if(!Number.isInteger(entry)||entry<2||entry>trash.getLastRow())throw Error('Строка корзины не найдена');
      const r=trash.getRange(entry,1,1,7).getValues()[0];
      if(String(r[4])!=='Удалена')throw Error('Запись уже обработана');
      const key=wlMonthCellKey_(r[0])||String(r[0]||'');
      if(wlFindMonth_(ss,key,true).status!=='Открыт')throw Error('Для восстановления откройте '+key);
      let row;try{row=JSON.parse(r[3]);}catch(_){throw Error('Снимок повреждён');}
      if(!row||!row.id||activeIds.has(key+'|'+row.id))throw Error('Одна из строк уже существует в бюджете');
      activeIds.add(key+'|'+row.id);records.push({entry,key,row});
    });
    const now=new Date();
    const start=ops.getLastRow()+1;
    ops.getRange(start,1,records.length,1).setNumberFormat('@');
    ops.getRange(start,1,records.length,3).setValues(records.map(x=>[x.key,now,JSON.stringify(x.row)]));
    records.forEach(x=>trash.getRange(x.entry,5,1,3).setValues([['Восстановлена',now,u.email]]));
    [...new Set(records.map(x=>x.key))].forEach(k=>{wlInvalidateOpsIndex_(k);wlInvalidateFollowingMonths_(ss,k,false);});
    return {ok:true,restored:records.length};
  });
}

/** v4.3: очистка корзины. Данные безвозвратно удаляются из корзины; минимальный журнал сохраняется. */
function wlTrashPurge(entries){
  if(!Array.isArray(entries)||!entries.length||entries.length>300)throw Error('Выберите до 300 строк');
  return wlLock_(()=>{
    const ss=wlStore_(),u=wlAdmin_(ss),trash=wlTrashSheet_(ss),n=trash.getLastRow();
    const wanted=new Set(entries.map(Number));
    if(wanted.size!==entries.length)throw Error('Повторяются номера строк');
    for(const x of wanted)if(!Number.isInteger(x)||x<2||x>n)throw Error('Запись корзины не найдена, обновите список');
    const data=n>1?trash.getRange(2,1,n-1,7).getValues():[],kept=[],audit=[];
    data.forEach((r,i)=>{
      if(!wanted.has(i+2)){kept.push(r);return;}
      let row={};try{row=JSON.parse(r[3])||{};}catch(_){}
      audit.push([new Date(),u.email,wlMonthCellKey_(r[0])||String(r[0]),String(row.id||''),String(row.department||'')]);
    });
    const log=wlNewSheet_(ss,'Журнал очистки корзины',['Время','Администратор','Месяц','ID строки','Отдел']);
    log.getRange(log.getLastRow()+1,1,audit.length,5).setValues(audit);
    trash.getRange(2,1,data.length,7).clearContent();
    if(kept.length)trash.getRange(2,1,kept.length,7).setValues(kept);
    return {ok:true,purged:audit.length};
  });
}

/** v4.3: импорт строк XLSX, разобранных в браузере. Не перезаписывает существующие строки. */
function wlImportExcelBatch(key, payload){
  key=wlKey_(key);
  if(!Array.isArray(payload)||!payload.length||payload.length>80)throw Error('Допустимо от 1 до 80 строк в пакете');
  return wlLock_(()=>{
    const ss=wlStore_();wlAdmin_(ss);
    const month=wlFindMonth_(ss,key);
    if(month.status!=='Открыт')throw Error('Месяц закрыт для изменений: '+key);
    const dict=wlDictionaries_(ss),dset=new Set(dict.departments.map(x=>x.name)),cats=new Set([...dict.categories,...dict.incomeCategories]),pays=new Set(dict.payments),srcs=new Set(dict.sources);
    const ops=wlOps_(ss),existing=wlOpsRead_(ops,key),ids=new Set(existing.map(x=>String(x.row.id)));
    const signature=r=>JSON.stringify([r.department,r.payment||'',r.source||'',r.category||'',r.description||'',r.amounts||{}]);
    const importedSignatures=new Set(existing.filter(x=>/^xl-/.test(String(x.row.id))).map(x=>signature(x.row)));
    // Удалённые ранее импортированные ID нельзя вернуть обходом корзины.
    const trash=wlTrashSheet_(ss),old=trash.getLastRow()>1?trash.getRange(2,1,trash.getLastRow()-1,5).getValues():[];
    old.forEach(v=>{if(wlMonthCellKey_(v[0])===key&&String(v[4])==='Удалена'){try{const r=JSON.parse(v[3]);ids.add(String(r.id));}catch(_){}}});
    const seen=new Set(),rows=[];
    payload.forEach((raw,i)=>{
      if(!raw||typeof raw!=='object')throw Error('Повреждённая строка '+(i+1));
      const id=String(raw.id||'').trim(),department=wlNormalizeDept_(raw.department),payment=String(raw.payment||'').trim(),source=String(raw.source||'').trim(),category=String(raw.category||'').trim(),description=String(raw.description||'').trim();
      if(!/^imp-[0-9]{6}-[a-z0-9]{8}$/.test(id))throw Error('Неверный ID импорта, строка '+(i+1));
      if(seen.has(id))throw Error('Повтор ID в файле: '+id);seen.add(id);
      if(!dset.has(department))throw Error('Отдел отсутствует в справочнике: '+department);
      if(payment&&!pays.has(payment))throw Error('Способ оплаты отсутствует в справочнике: '+payment);
      if(source&&!srcs.has(source))throw Error('Источник оплаты отсутствует в справочнике: '+source);
      if(category&&!cats.has(category))throw Error('Статья ДДС отсутствует в справочнике: '+category);
      if(description.length>2000)throw Error('Описание длиннее 2000 символов');
      if(!description&&!category&&!payment&&!source&&!Object.keys(raw.amounts||{}).length)throw Error('Пустая строка в пакете');
      const [y,m]=key.split('-').map(Number),days=new Date(y,m,0).getDate(),amounts={};
      Object.keys(raw.amounts||{}).forEach(day=>{
        const d=Number(day),value=Number(raw.amounts[day]);
        if(!Number.isInteger(d)||d<1||d>days)throw Error('Несуществующая дата: '+day+' для '+key);
        if(!Number.isFinite(value)||Math.abs(value)>1e12)throw Error('Недопустимая сумма в '+id);
        if(value!==0)amounts[d]=Math.round(value*100)/100;
      });
      const prepared={id,department,payment,source,category,description,amounts,cashFlowType:dict.incomeCategories.includes(category)&&!dict.categories.includes(category)?'income':'expense',reviewPending:!category};
      if(!ids.has(id)&&!importedSignatures.has(signature(prepared)))rows.push(prepared);
    });
    if(rows.length){
      const backupKey='WL_DDS_XLSX_BACKUP_ID';
      if(!PropertiesService.getScriptProperties().getProperty(backupKey)){
        const snap=ss.copy('Резервная копия WL ДДС перед импортом Excel — '+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm'));
        PropertiesService.getScriptProperties().setProperty(backupKey,snap.getId());
      }
      const stamp=new Date(),start=ops.getLastRow()+1;
      ops.getRange(start,1,rows.length,1).setNumberFormat('@');
      ops.getRange(start,1,rows.length,3).setValues(rows.map(r=>[key,stamp,JSON.stringify(r)]));
      wlInvalidateOpsIndex_(key);wlInvalidateFollowingMonths_(ss,key,false);
    }
    return {ok:true,added:rows.length,skipped:payload.length-rows.length};
  });
}


/** Monthly audit maintenance: archive historic purge events instead of erasing evidence.
 * Installed only after an administrator explicitly enables the schedule. */
function wlMonthlyAuditMaintenance(){
  return wlLock_(()=>{
    const ss=wlStore_();
    const log=ss.getSheetByName('Журнал очистки корзины');
    if(!log||log.getLastRow()<2)return {archived:0};
    const rows=log.getRange(2,1,log.getLastRow()-1,5).getValues();
    const cutoff=new Date();cutoff.setDate(cutoff.getDate()-35);
    const older=[],recent=[];
    for(const row of rows){
      const date=row[0] instanceof Date?row[0]:new Date(row[0]);
      if(!isNaN(date.getTime())&&date<cutoff)older.push(row);else recent.push(row);
    }
    if(!older.length)return {archived:0};
    const archive=wlNewSheet_(ss,'Архив журнала очистки',['Время','Администратор','Месяц','ID строки','Отдел']);
    archive.getRange(archive.getLastRow()+1,1,older.length,5).setValues(older);
    log.getRange(2,1,rows.length,5).clearContent();
    if(recent.length)log.getRange(2,1,recent.length,5).setValues(recent);
    return {archived:older.length,remaining:recent.length};
  });
}
/** Run once as admin. Trigger will archive 35+ day entries at the start of each month. */
function wlEnableMonthlyAuditArchive(){
  const ss=wlStore_();wlAdmin_(ss);
  const handler='wlMonthlyAuditMaintenance';
  try{
    const old=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===handler);
    if(!old.length)ScriptApp.newTrigger(handler).timeBased().onMonthDay(1).atHour(3).create();
    return {enabled:true,triggerCount:old.length||1};
  }catch(e){
    if(wlYearArchiveMissingTriggerScope_(e))throw Error('Для автоматической очистки журнала требуется разрешение Google. Выполните в редакторе Apps Script wlAuthorizeYearArchive, подтвердите доступ, затем включите очистку ещё раз.');
    throw e;
  }
}


/** Admin-only preview for moving royalty receipts out of expense categories. */
function wlRoyaltyIncomePreview(){
 const ss=wlStore_();wlAdmin_(ss);
 const labels=['Получение роялти','Поступление роялти'];
 const dict=wlDictionaries_(ss),target=dict.incomeCategories.find(v=>labels.includes(v))||'Получение роялти';
 const entries=wlOpsRead_(wlOps_(ss)).filter(e=>labels.includes(String(e.row.category||'').trim())&&(e.row.cashFlowType!=='income'||String(e.row.category||'')!==target));
 return {count:entries.length,total:Math.round(entries.reduce((sum,e)=>sum+Object.values(e.row.amounts||{}).reduce((v,n)=>v+(Number(n)||0),0),0)*100)/100};
}
/** One-time, reversible-by-audit reclassification; does not delete the budget rows. */
function wlMigrateRoyaltyIncome(){
 return wlLock_(()=>{
  const ss=wlStore_(),u=wlAdmin_(ss),dict=wlDictionaries_(ss),labels=['Получение роялти','Поступление роялти'];
  const target=dict.incomeCategories.find(v=>labels.includes(v))||'Получение роялти';
  const ops=wlOps_(ss),entries=wlOpsRead_(ops).filter(e=>labels.includes(String(e.row.category||'').trim())&&(e.row.cashFlowType!=='income'||String(e.row.category||'')!==target));
  const sh=wlDictSheet_(ss),n=sh.getLastRow();
  const all=n>1?sh.getRange(2,4,n-1,2).getDisplayValues():[];
  const already=all.some(r=>r[1]===target);
  // Prepare before-image records before touching any operation or dictionary item.
  const log=wlNewSheet_(ss,'Журнал переноса роялти',['Время','Администратор','Месяц','ID строки','Прежняя статья','Новая статья','JSON до переноса']);
  const stamp=new Date();
  if(entries.length){
   const audit=entries.map(e=>[stamp,u.email,e.key,String(e.row.id||''),String(e.row.category||''),target,JSON.stringify(e.row)]);
   log.getRange(log.getLastRow()+1,1,audit.length,7).setValues(audit);
  }
  entries.forEach(e=>{
   e.row.category=target;
   e.row.cashFlowType='income';
   if(e.row.templateBaseline&&typeof e.row.templateBaseline==='object'){
    e.row.templateBaseline.category=target;e.row.templateBaseline.cashFlowType='income';
   }
  });
  wlWriteExistingRowChanges_(ops,entries,stamp);
  if(!already){
   const gap=all.findIndex(r=>!r[1]);
   sh.getRange(gap<0?sh.getLastRow()+1:gap+2,5).setValue(target);
  }
  const purged=[];
  all.forEach((r,i)=>{if(labels.includes(String(r[0]||'').trim()))purged.push(i+2);});
  purged.forEach(row=>sh.getRange(row,4).clearContent());
  [...new Set(entries.map(e=>e.key))].forEach(key=>{
   wlInvalidateOpsIndex_(key);wlInvalidateFollowingMonths_(ss,key,false);
  });
  wlInvalidateDictionaries_();
  return {ok:true,changed:entries.length,removedExpenseEntries:purged.length,incomeCategory:target,auditSheet:'Журнал переноса роялти'};
 });
}


/** Round-trip Excel import. The browser provides exported ID and fingerprint for existing rows.
 * Never delete rows omitted from the file; never overwrite a row changed after export.
 * Admin-only, month must be open. All checks are repeated at commit time under the lock.
 */
function wlExcelFingerprint_(key,row){
  const amounts=row.amounts||{},vals=[];
  for(let day=1;day<=31;day++)vals.push(Math.round((Number(amounts[day])||0)*100)/100);
  const value=JSON.stringify([key,wlNormalizeDept_(row.department),String(row.payment||''),String(row.source||''),String(row.category||''),String(row.description||''),vals]);
  let h=2166136261;for(let i=0;i<value.length;i++){h^=value.charCodeAt(i);h=Math.imul(h,16777619);}
  return (h>>>0).toString(36).padStart(8,'0').slice(-8);
}
function wlExcelRoundTripCore_(key,payload,dryRun){
  key=wlKey_(key);
  if(!Array.isArray(payload)||!payload.length||payload.length>(dryRun?800:60))throw Error('Слишком много строк в одном пакете Excel');
  const action=()=>{
    const ss=wlStore_();wlAdmin_(ss);
    const month=wlFindMonth_(ss,key);
    if(month.status!=='Открыт')throw Error('Месяц закрыт: '+key);
    const dict=wlDictionaries_(ss),allowedDeps=new Set(dict.departments.map(d=>d.name)),pays=new Set(dict.payments),sources=new Set(dict.sources),cats=new Set([...dict.categories,...dict.incomeCategories]);
    const ops=wlOps_(ss),existing=wlOpsRead_(ops,key),byId=new Map(),deleted=wlDeletedIdsForMonth_(ss,key);
    for(const record of existing){
      const id=String(record.row.id||'');
      if(byId.has(id))throw Error('В базе повторяется ID '+id+'. Сначала исправьте дубли, импорт остановлен');
      byId.set(id,record);
    }
    const seen=new Set(),add=[],edit=[],stats={added:0,updated:0,unchanged:0,conflicts:[]};
    payload.forEach((raw,index)=>{
      if(!raw||typeof raw!=='object')throw Error('Повреждённая запись '+(index+1));
      const id=String(raw.id||'').trim(),base=String(raw.baseSig||'').trim(),dep=wlNormalizeDept_(raw.department),payment=String(raw.payment||'').trim(),source=String(raw.source||'').trim(),cat=String(raw.category||'').trim(),description=String(raw.description||'').trim();
      if(!/^[A-Za-z0-9._:-]{8,120}$/.test(id))throw Error('Некорректный ID строки: '+id);
      if(seen.has(id))throw Error('Одинаковый ID дважды в файле: '+id);seen.add(id);
      if(!allowedDeps.has(dep))throw Error('Неизвестный отдел '+dep);
      if(payment&&!pays.has(payment))throw Error('Неизвестный способ оплаты '+payment);
      if(source&&!sources.has(source))throw Error('Неизвестный источник оплаты '+source);
      if(cat&&!cats.has(cat))throw Error('Статья ДДС отсутствует в справочнике: '+cat);
      if(description.length>2000)throw Error('Описание превышает 2000 символов');
      const amounts={},max=new Date(Number(key.slice(0,4)),Number(key.slice(5,7)),0).getDate();
      Object.keys(raw.amounts||{}).forEach(day=>{
        const d=Number(day),n=Number(raw.amounts[day]);
        if(!Number.isInteger(d)||d<1||d>max||!Number.isFinite(n)||Math.abs(n)>1e12||Math.abs(n*100-Math.round(n*100))>1e-6)throw Error('Неправильная сумма, дата или более двух знаков после запятой: '+id+' / '+day);
        if(n!==0)amounts[d]=Math.round(n*100)/100;
      });
      const old=byId.get(id);
      const inIncome=dict.incomeCategories.includes(cat),inExpense=dict.categories.includes(cat);
      const requestType=String(raw.cashFlowType||'').trim();
      if(requestType&&requestType!=='income'&&requestType!=='expense')throw Error('Неизвестный тип статьи '+id);
      if(cat&&requestType==='income'&&!inIncome)throw Error('Статья не относится к поступлениям: '+cat);
      if(cat&&requestType==='expense'&&!inExpense)throw Error('Статья не относится к списаниям: '+cat);
      if(cat&&inIncome&&inExpense&&!requestType&&!(old&&old.row.cashFlowType))throw Error('Уточните тип потока для статьи '+cat);
      const nextType=requestType||(old&&old.row.category===cat&&old.row.cashFlowType)||(inIncome&&!inExpense?'income':'expense');
      const next={id,department:dep,payment,source,category:cat,description,amounts,cashFlowType:nextType};
      if(old){
        // When the same blank-template import is uploaded twice, its stable imp- ID
        // must be idempotent. Without a fingerprint it may only be a no-op;
        // never overwrite an already stored row based on an unverified template.
        if(!base&&/^imp-/.test(id)){
          if(wlExcelFingerprint_(key,old.row)===wlExcelFingerprint_(key,next)){
            stats.unchanged++;return;
          }
          stats.conflicts.push(dep+' / '+id+' — строка уже импортирована и отличается; скачайте актуальный бюджет');return;
        }
        if(!/^[a-z0-9]{8}$/.test(base))throw Error('Для существующей строки '+id+' отсутствует контрольная сумма. Скачайте свежий бюджет');
        if(wlExcelFingerprint_(key,old.row)!==base){stats.conflicts.push(dep+' / '+id+' — данные изменились после выгрузки');return;}
        if(wlExcelFingerprint_(key,next)===base){stats.unchanged++;return;}
        const merged=Object.assign({},old.row,{department:dep,payment,source,category:cat,description,amounts,cashFlowType:next.cashFlowType});
        if(merged.formulas&&typeof merged.formulas==='object'){
          merged.formulas=Object.assign({},merged.formulas);
          for(let d=1;d<=max;d++)if((Number(old.row.amounts[d])||0)!==(Number(amounts[d])||0))delete merged.formulas[d];
        }
        edit.push({line:old.line,row:merged});stats.updated++;
      }else{
        if(base){stats.conflicts.push(dep+' / '+id+' — строка не найдена (возможно удалена или перемещена)');return;}
        if(!/^imp-/.test(id))throw Error('Новая строка не может иметь произвольный ID: '+id);
        if(deleted.has(id)){stats.conflicts.push(dep+' / '+id+' — запись уже удалялась');return;}
        next.reviewPending=!cat;add.push(next);stats.added++;
      }
    });
    if(stats.conflicts.length){if(!dryRun)throw Error('Конфликты Excel: '+stats.conflicts.slice(0,5).join('; '));return stats;}
    if(dryRun)return stats;
    if(edit.length||add.length){
      const props=PropertiesService.getScriptProperties(),backupKey='WL_DDS_EXCEL_ROUNDTRIP_BACKUP';
      if(!props.getProperty(backupKey)){
        const snapshot=ss.copy('WL ДДС — резерв перед обновлением через Excel — '+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm'));
        props.setProperty(backupKey,snapshot.getId());
      }
      const stamp=new Date();
      if(edit.length)wlWriteExistingRowChanges_(ops,edit,stamp);
      if(add.length){
        const first=ops.getLastRow()+1;
        ops.getRange(first,1,add.length,1).setNumberFormat('@');
        ops.getRange(first,1,add.length,3).setValues(add.map(row=>[key,stamp,JSON.stringify(row)]));
      }
      wlInvalidateOpsIndex_(key);wlInvalidateFollowingMonths_(ss,key,false);
    }
    return stats;
  };
  return dryRun?action():wlLock_(action);
}
function wlPreviewExcelRoundTrip(key,payload){return wlExcelRoundTripCore_(key,payload,true);}
function wlApplyExcelRoundTrip(key,payload){return wlExcelRoundTripCore_(key,payload,false);}


/** ЦФО/БДДС: агрегирование по всем доступным отделам без записи в Sheets. */
const WL_BDDS_CLASSIFICATION = {"income|продажа товара lamoda":"operating","income|продажа товара ozon":"operating","income|продажа товара wb":"operating","income|прочие поступления":"operating","income|виртуальные поступления":"operating","income|возвраты от поставщиков":"operating","income|корректировка":"operating","income|курсовая разница":"operating","income|оптовые продажи":"operating","income|получение роялти":"operating","income|поступления за it-услуги и сервисы":"operating","income|приход не связанный с деятельностью мп":"operating","income|продажа товаров интернет-магазин (в т.ч. инстаграм)":"operating","income|съемка_возврат реквизита":"operating","income|торговля на мероприятиях":"operating","expense|расходы на персонал":"operating","expense|взносы в фонды":"operating","expense|заработная плата":"operating","expense|командировочные и представительские расходы":"operating","expense|корпоративные мероприятия и подарки":"operating","expense|ндфл":"operating","expense|обучение и развитие персонала":"operating","expense|подбор персонала":"operating","expense|корпоративный социальный пакет":"operating","expense|сервис для оплаты самозанятых":"operating","expense|реклама и маркетинг":"operating","expense|выкупы":"operating","expense|выплата зп контент мейкеры":"operating","expense|выплата зп креаторам":"operating","expense|контекстная реклама":"operating","expense|маркетинг прочее":"operating","expense|маркетинговые и рекламные расходы (если не относятся к коммерческим расходам)":"operating","expense|маркировка рекламы":"operating","expense|реклама на платформах":"operating","expense|реклама на мп":"operating","expense|реклама у блогеров (инфлюенс)":"operating","expense|реклама у блогеров (имиджевая)":"operating","expense|фотосессии и видеосъемки":"operating","expense|съемка_аренда помещения/студии":"operating","expense|таргетированная реклама":"operating","expense|фото/видеосъемка+монтаж":"operating","expense|комиссия сервиса оплат":"operating","expense|комиссия сервиса оплат ecom":"operating","expense|выкупы продукции конкурентов":"operating","expense|маркетинговые товары":"operating","expense|реклама бартеры (fast-sales)":"operating","expense|аффилиаты":"operating","expense|самовыкупы через подрядчика (товары)":"operating","expense|самовыкупы через подрядчика (услуги)":"operating","expense|маркетинговые мероприятия":"operating","expense|съемки для каталога":"operating","expense|сервис для оплаты самозанятых (фбк)":"operating","expense|отзывы покупателей":"operating","expense|реклама у блогеров":"operating","expense|коллаборации":"operating","expense|кураторство креаторов":"operating","expense|логистические и складские расходы":"operating","expense|дополнительная логистика фулфилмента":"operating","expense|доставка по россии":"operating","expense|международные доставки":"operating","expense|доставка товара со склада на склада мп":"operating","expense|прочая логистика":"operating","expense|прочие услуги фулфилмента (переклейка, маркировка)":"operating","expense|фулфиллмент":"operating","expense|доставка образцов по рф":"operating","expense|логистика (бренд)":"operating","expense|закупка товар":"operating","expense|закупка товаров":"operating","expense|новинка":"operating","expense|закупки план":"operating","expense|закупка упаковки":"operating","expense|комиссия по переводам на закупки":"operating","expense|запуски":"operating","expense|закупка образцов невозвратная":"operating","expense|прототипирование и тестирование":"operating","expense|разработка дизайна и эскизов":"operating","expense|разработка новых продуктов":"operating","expense|закупка и тестирование новых тканей и материалов":"operating","expense|инспекция товара":"operating","expense|маркировка товара":"operating","expense|сертификация товаров":"operating","expense|административно-хозяйственные расходы":"operating","expense|авансы по услугам":"operating","expense|аренда офиса и складских помещений":"operating","expense|ведение б/у и н/у":"operating","expense|консультация":"operating","expense|мебель, офисное оборудование, хозтовары, канцелярия":"operating","expense|обслуживание помещений и коммунальные расходы":"operating","expense|ремонт основных средств":"operating","expense|содержание офиса и склада":"operating","expense|услуги связи":"operating","expense|химчистка товаров после возвратов":"operating","expense|юридические услуги":"operating","expense|америка":"operating","expense|закупка товаров сша":"operating","expense|расходы usa":"operating","expense|налоги":"operating","expense|налоги на доходы (прибыль)":"operating","expense|штрафы и пени по налогам":"operating","expense|ндс":"operating","expense|сервисы, по и разработка":"operating","expense|расходы на программное обеспечение":"operating","expense|расходы на it и техподдержку":"operating","expense|сервисы аналитики мп":"operating","expense|сервисы аналитики не относящиеся к мп":"operating","expense|сервисы аналитики рекламы":"operating","expense|создание нма":"operating","expense|расходы на it, по, сервисы (бренд)":"operating","expense|seo (поисковая оптимизация)":"operating","expense|комиссии и проценты":"operating","expense|банковские комиссии":"operating","expense|комиссия брокеров":"operating","expense|проценты по кредиту":"operating","expense|банковские услуги":"operating","expense|прочие":"operating","expense|благотворительность":"operating","expense|виртуальные списания":"operating","expense|возвраты товара":"operating","expense|выплата роялти":"operating","expense|прочие расходы":"operating","expense|расходы не связанный с деятельностью на мп":"operating","expense|тест кассы":"operating","expense|курсовая разница":"operating","income|возврат выданного займа":"investment","income|возврат депозита":"investment","income|продажа основных средств":"investment","income|прочие поступления по инвестиционной деятельности":"investment","expense|выдача займа другой компании":"investment","expense|депозит":"investment","expense|покупка основных средств":"investment","income|получение кредита":"financing","income|проценты по депозиту":"financing","expense|выплата тела кредита":"financing","expense|дивиденды":"financing","expense|вывод прибыли":"financing"};

const WL_BDDS_GROUP = {"income|продажа товара lamoda":"Поступления от МП","income|продажа товара ozon":"Поступления от МП","income|продажа товара wb":"Поступления от МП","income|виртуальные поступления":"Прочие поступления","income|возвраты от поставщиков":"Прочие поступления","income|корректировка":"Прочие поступления","income|курсовая разница":"Прочие поступления","income|оптовые продажи":"Прочие поступления","income|получение роялти":"Прочие поступления","income|поступления за it-услуги и сервисы":"Прочие поступления","income|приход не связанный с деятельностью мп":"Прочие поступления","income|продажа товаров интернет-магазин (в т.ч. инстаграм)":"Прочие поступления","income|съемка_возврат реквизита":"Прочие поступления","income|торговля на мероприятиях":"Прочие поступления","expense|расходы на персонал":"Расходы на персонал","expense|взносы в фонды":"Расходы на персонал","expense|заработная плата":"Расходы на персонал","expense|командировочные и представительские расходы":"Расходы на персонал","expense|корпоративные мероприятия и подарки":"Расходы на персонал","expense|ндфл":"Расходы на персонал","expense|обучение и развитие персонала":"Расходы на персонал","expense|подбор персонала":"Расходы на персонал","expense|корпоративный социальный пакет":"Расходы на персонал","expense|сервис для оплаты самозанятых":"Расходы на персонал","expense|реклама и маркетинг":"Реклама и маркетинг","expense|выкупы":"Реклама и маркетинг","expense|выплата зп контент мейкеры":"Реклама и маркетинг","expense|выплата зп креаторам":"Реклама и маркетинг","expense|контекстная реклама":"Реклама и маркетинг","expense|маркетинг прочее":"Реклама и маркетинг","expense|маркетинговые и рекламные расходы (если не относятся к коммерческим расходам)":"Реклама и маркетинг","expense|маркировка рекламы":"Реклама и маркетинг","expense|реклама на платформах":"Реклама и маркетинг","expense|реклама на мп":"Реклама и маркетинг","expense|реклама у блогеров (инфлюенс)":"Реклама и маркетинг","expense|реклама у блогеров (имиджевая)":"Реклама и маркетинг","expense|фотосессии и видеосъемки":"Реклама и маркетинг","expense|съемка_аренда помещения/студии":"Реклама и маркетинг","expense|таргетированная реклама":"Реклама и маркетинг","expense|фото/видеосъемка+монтаж":"Реклама и маркетинг","expense|комиссия сервиса оплат":"Реклама и маркетинг","expense|комиссия сервиса оплат ecom":"Реклама и маркетинг","expense|выкупы продукции конкурентов":"Реклама и маркетинг","expense|маркетинговые товары":"Реклама и маркетинг","expense|реклама бартеры (fast-sales)":"Реклама и маркетинг","expense|аффилиаты":"Реклама и маркетинг","expense|самовыкупы через подрядчика (товары)":"Реклама и маркетинг","expense|самовыкупы через подрядчика (услуги)":"Реклама и маркетинг","expense|маркетинговые мероприятия":"Реклама и маркетинг","expense|съемки для каталога":"Реклама и маркетинг","expense|сервис для оплаты самозанятых (фбк)":"Реклама и маркетинг","expense|отзывы покупателей":"Реклама и маркетинг","expense|реклама у блогеров":"Реклама и маркетинг","expense|коллаборации":"Реклама и маркетинг","expense|кураторство креаторов":"Реклама и маркетинг","expense|логистические и складские расходы":"Логистика и склад","expense|дополнительная логистика фулфилмента":"Логистика и склад","expense|доставка по россии":"Логистика и склад","expense|международные доставки":"Логистика и склад","expense|доставка товара со склада на склада мп":"Логистика и склад","expense|прочая логистика":"Логистика и склад","expense|прочие услуги фулфилмента (переклейка, маркировка)":"Логистика и склад","expense|фулфиллмент":"Логистика и склад","expense|доставка образцов по рф":"Логистика и склад","expense|логистика (бренд)":"Логистика и склад","expense|закупка товар":"Закупка товара","expense|закупка товаров":"Закупка товара","expense|новинка":"Закупка товара","expense|закупки план":"Закупка товара","expense|закупка упаковки":"Закупка товара","expense|комиссия по переводам на закупки":"Закупка товара","expense|запуски":"Запуски","expense|закупка образцов невозвратная":"Запуски","expense|прототипирование и тестирование":"Запуски","expense|разработка дизайна и эскизов":"Запуски","expense|разработка новых продуктов":"Запуски","expense|закупка и тестирование новых тканей и материалов":"Запуски","expense|инспекция товара":"Запуски","expense|маркировка товара":"Запуски","expense|сертификация товаров":"Запуски","expense|административно-хозяйственные расходы":"Административные расходы","expense|авансы по услугам":"Административные расходы","expense|аренда офиса и складских помещений":"Административные расходы","expense|ведение б/у и н/у":"Административные расходы","expense|консультация":"Административные расходы","expense|мебель, офисное оборудование, хозтовары, канцелярия":"Административные расходы","expense|обслуживание помещений и коммунальные расходы":"Административные расходы","expense|ремонт основных средств":"Административные расходы","expense|содержание офиса и склада":"Административные расходы","expense|услуги связи":"Административные расходы","expense|химчистка товаров после возвратов":"Административные расходы","expense|юридические услуги":"Административные расходы","expense|америка":"Америка","expense|закупка товаров сша":"Америка","expense|расходы usa":"Америка","expense|налоги":"Налоги","expense|налоги на доходы (прибыль)":"Налоги","expense|штрафы и пени по налогам":"Налоги","expense|ндс":"Налоги","expense|сервисы, по и разработка":"Сервисы, ПО и разработка","expense|расходы на программное обеспечение":"Сервисы, ПО и разработка","expense|расходы на it и техподдержку":"Сервисы, ПО и разработка","expense|сервисы аналитики мп":"Сервисы, ПО и разработка","expense|сервисы аналитики не относящиеся к мп":"Сервисы, ПО и разработка","expense|сервисы аналитики рекламы":"Сервисы, ПО и разработка","expense|создание нма":"Сервисы, ПО и разработка","expense|расходы на it, по, сервисы (бренд)":"Сервисы, ПО и разработка","expense|seo (поисковая оптимизация)":"Сервисы, ПО и разработка","expense|комиссии и проценты":"Комиссии и проценты","expense|банковские комиссии":"Комиссии и проценты","expense|комиссия брокеров":"Комиссии и проценты","expense|проценты по кредиту":"Комиссии и проценты","expense|банковские услуги":"Комиссии и проценты","expense|прочие":"Прочие","expense|благотворительность":"Прочие","expense|виртуальные списания":"Прочие","expense|возвраты товара":"Прочие","expense|выплата роялти":"Прочие","expense|прочие расходы":"Прочие","expense|расходы не связанный с деятельностью на мп":"Прочие","expense|тест кассы":"Прочие","expense|курсовая разница":"Прочие"};
const WL_BDDS_ORDER={operating:{income:['Поступления от МП','Прочие поступления'],expense:['Расходы на персонал','Реклама и маркетинг','Логистика и склад','Закупка товара','Запуски','Административные расходы','Америка','Налоги','Сервисы, ПО и разработка','Комиссии и проценты','Прочие']}};
function wlBddsNormalize_(value){return String(value||'').toLocaleLowerCase('ru').replace(/ё/g,'е').replace(/[\u00a0\u202f]/g,' ').replace(/\s+/g,' ').trim();}
const WL_BDDS_NORM_CLASS=Object.keys(WL_BDDS_CLASSIFICATION).reduce((o,k)=>{o[wlBddsNormalize_(k)]=WL_BDDS_CLASSIFICATION[k];return o;},{});
const WL_BDDS_NORM_GROUP=Object.keys(WL_BDDS_GROUP).reduce((o,k)=>{o[wlBddsNormalize_(k)]=WL_BDDS_GROUP[k];return o;},{});
function wlBddsGroupOptions_(type){
 const options=(WL_BDDS_ORDER.operating[type]||[]).map(group=>({activity:'operating',group}));
 for(const activity of ['investment','financing'])options.push({activity,group:type==='income'?'Поступления':'Списания'});
 return options;
}
function wlBddsGrouping_(ss){
 const defaults={};Object.keys(WL_BDDS_NORM_CLASS).forEach(k=>{const type=k.split('|')[0],activity=WL_BDDS_NORM_CLASS[k];defaults[k]={activity,group:WL_BDDS_NORM_GROUP[k]||(activity==='operating'?(type==='income'?'Прочие поступления':'Прочие'):(type==='income'?'Поступления':'Списания'))};});
 const sh=ss.getSheetByName('WL_Группы_статей'),overrides={};
 if(sh&&sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,4).getValues().forEach(r=>{const type=String(r[0]),key=type+'|'+wlBddsNormalize_(r[1]),activity=String(r[2]),group=String(r[3]);if(wlBddsGroupOptions_(type).some(x=>x.activity===activity&&x.group===group))overrides[key]={activity,group};});
 return {defaults,overrides,options:{income:wlBddsGroupOptions_('income'),expense:wlBddsGroupOptions_('expense')}};
}
function wlBddsMeta_(type,category,grouping){
 const key=type+'|'+wlBddsNormalize_(category);return (grouping&&grouping.overrides[key])||(grouping&&grouping.defaults[key])||{activity:'unclassified',group:'Требуют распределения'};
}
function wlSetArticleGroup(type,category,activity,group){
 type=String(type||'');category=String(category||'').trim();activity=String(activity||'');group=String(group||'');
 if(!['income','expense'].includes(type)||!category||category.length>250)throw Error('Некорректная статья');
 if((activity||group)&&!wlBddsGroupOptions_(type).some(x=>x.activity===activity&&x.group===group))throw Error('Выберите группу из списка');
 return wlLock_(()=>{const ss=wlStore_(),user=wlAdmin_(ss);wlRequireTab_(user,'ЦФО',false);const sh=wlNewSheet_(ss,'WL_Группы_статей',['Тип','Статья','Деятельность','Группа','Изменено','Кто изменил']);const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,4).getValues():[];const key=wlBddsNormalize_(category);const i=rows.findIndex(r=>String(r[0])===type&&wlBddsNormalize_(r[1])===key);if(!activity&&!group){if(i>=0)sh.deleteRow(i+2);}else sh.getRange(i<0?sh.getLastRow()+1:i+2,1,1,6).setValues([[type,category,activity,group,new Date(),user.email]]);return wlBddsGrouping_(ss);});
}
function wlBddsComments_(ss,year){
 const sh=ss.getSheetByName('WL_БДДС_Комментарии'),all={};if(!sh||sh.getLastRow()<2)return all;
 const zone=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone();
 sh.getRange(2,1,sh.getLastRow()-1,4).getValues().forEach(row=>{
  if(String(row[0])!==year)return;
  const month=wlMonthCellKey_(row[1],zone);if(!month)return;
  try{let messages=JSON.parse(String(row[3])||'[]');if(Array.isArray(messages)){const key=month+'|'+row[2];all[key]=(all[key]||[]).concat(messages);}}catch(_){ }
 });Object.keys(all).forEach(key=>{all[key].sort((a,b)=>String(a.at||'').localeCompare(String(b.at||'')));all[key]=all[key].slice(-30);});return all;
}
function wlBddsFunds_(ss,year){
 const result={},sh=ss.getSheetByName('WL_БДДС_Фонды');if(!sh||sh.getLastRow()<2)return result;
 const zone=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone();
 sh.getRange(2,1,sh.getLastRow()-1,3).getValues().forEach(r=>{const m=wlMonthCellKey_(r[0],zone);if(!m||!m.startsWith(year+'-'))return;result[m]={begin:r[1]===''?null:Number(r[1]),end:r[2]===''?null:Number(r[2])};});return result;
}
function wlBddsSetFund(month,field,value){
 month=wlKey_(month);if(!['begin','end'].includes(field))throw Error('Неизвестный показатель');
 if(value!==null&&(typeof value!=='number'||!Number.isFinite(value)))throw Error('Введите корректную сумму');
 return wlLock_(()=>{const ss=wlStore_(),u=wlAdmin_(ss);wlRequireTab_(u,'БДДС',false);wlFindMonth_(ss,month,false);
  const sh=wlNewSheet_(ss,'WL_БДДС_Фонды',['Месяц','На начало периода','На конец периода','Изменено','Кто изменил']);
  const zone=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone(),old=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,1).getValues():[],idx=old.findIndex(r=>wlMonthCellKey_(r[0],zone)===month),line=idx<0?sh.getLastRow()+1:idx+2;
  if(idx<0){if(value===null)return wlBddsFunds_(ss,month.slice(0,4));sh.getRange(line,1).setValue(month);}
  const target=sh.getRange(line,field==='begin'?2:3);if(value===null)target.clearContent();else target.setValue(value);
  sh.getRange(line,4,1,2).setValues([[new Date(),u.email]]);return wlBddsFunds_(ss,month.slice(0,4));});
}
/** Read the separately maintained 2026 report without changing either spreadsheet. */
function wlBddsHistorical2026(){
 const ss=wlStore_(),user=wlAdmin_(ss);wlRequireTab_(user,'БДДС',false);
 const source=SpreadsheetApp.openById('1aVg5VqQkp1WSIeHYuK-SVrR5DRw22JqDy0qGhbN9Xkg'),sh=source.getSheetByName('БДДС');if(!sh)throw Error('В исходном файле не найден лист БДДС');
 const nr=Math.min(sh.getLastRow(),1000),nc=Math.min(sh.getLastColumn(),300);if(nr<2||nc<6)throw Error('В исходном БДДС не найдены данные');
 const range=sh.getRange(1,1,nr,nc),values=range.getValues(),shown=range.getDisplayValues(),zone=source.getSpreadsheetTimeZone()||Session.getScriptTimeZone(),months=[];
 values[0].forEach((v,i)=>{const key=wlMonthCellKey_(v,zone);if(key&&key.startsWith('2026-')&&i+4<nc)months.push({key,index:i});});
 if(!months.length)throw Error('Не найдены месяцы 2026 года в первой строке исходного БДДС');
 const rows=shown.slice(1).filter(row=>String(row[0]||'').trim()).map(row=>({label:String(row[0]),cells:months.map(m=>row.slice(m.index,m.index+5).map(x=>String(x||'')))}));
 return {months:months.map(m=>m.key),rows};
}
function wlGetBddsPlan(year){
 year=String(year||'');if(!/^\d{4}$/.test(year))throw Error('Неверный год');
 const ss=wlStore_(),u=wlUser_(ss);wlRequireTab_(u,'БДДС',false);
 if(!wlCanSeeBudgetYear_(u,year+'-01'))throw Error('Архив доступен только администратору');
 const active=wlMonthList_(ss).filter(m=>m.active&&m.key.startsWith(year));
 if(!active.length)return {year,months:[],articles:[],unmapped:[],comments:{},sectionOrder:WL_BDDS_ORDER};
 const allowed=new Set(active.map(m=>m.key)),departments=new Set(u.view),dict=wlDictionaries_(ss);
 const incomes=new Set(dict.incomeCategories.map(wlBddsNormalize_)),expenses=new Set(dict.categories.map(wlBddsNormalize_));
 const grouping=wlBddsGrouping_(ss),items=new Map(),unmapped=new Set(), populatedMonths=new Set();
 for(const entry of wlOpsRead_(wlOps_(ss))){
  if(!allowed.has(entry.key)||!departments.has(entry.row.department))continue;
  const row=entry.row,category=String(row.category||'').trim();
  if((!row.templateRoot && !String(row.id||'').startsWith('copy-')) || Object.keys(row.amounts||{}).length || Object.keys(row.notes||{}).length || Object.keys(row.formulas||{}).length) populatedMonths.add(entry.key);
  if(!category)continue;
  const normalized=wlBddsNormalize_(category),type=row.cashFlowType==='income'?'income':row.cashFlowType==='expense'?'expense':incomes.has(normalized)&&!expenses.has(normalized)?'income':'expense';
  const id=type+'|'+normalized,meta=wlBddsMeta_(type,category,grouping);
  const item=items.get(id)||{id,category,type,activity:meta.activity,group:meta.group,byMonth:{}};
  for(const value of Object.values(row.amounts||{})){
   const amount=Number(value);if(Number.isFinite(amount))item.byMonth[entry.key]=(item.byMonth[entry.key]||0)+amount;
  }
  if(meta.activity==='unclassified'&&Object.values(row.amounts||{}).some(v=>Number(v)!==0))unmapped.add((type==='income'?'Поступление: ':'Списание: ')+category);
  items.set(id,item);
 }
 const comments=u.admin?wlBddsComments_(ss,year):{};
 Object.keys(comments).forEach(k=>{const m=k.slice(0,7);if(allowed.has(m))populatedMonths.add(m);});
 return {year,months:[...populatedMonths].sort(),articles:[...items.values()].sort((a,b)=>a.category.localeCompare(b.category,'ru')),unmapped:[...unmapped],comments,funds:u.admin?wlBddsFunds_(ss,year):{},sectionOrder:WL_BDDS_ORDER};
}
function wlBddsPostComment(year,month,key,text){
 year=String(year||'');month=wlKey_(month);key=String(key||'').trim();text=String(text||'').trim();
 if(month.slice(0,4)!==year||!key||key.length>250||!text||text.length>1600)throw Error('Некорректный комментарий');
 return wlLock_(()=>{
  const ss=wlStore_(),user=wlAdmin_(ss);
  const sh=wlNewSheet_(ss,'WL_БДДС_Комментарии',['Год','Месяц','Ключ показателя','Сообщения JSON','Изменено']);
  const old=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,3).getValues():[];
  const zone=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone();
  const idx=old.findIndex(r=>String(r[0])===year&&wlMonthCellKey_(r[1],zone)===month&&String(r[2])===key);
  const line=idx<0?sh.getLastRow()+1:idx+2;
  let messages=[];if(idx>=0){try{messages=JSON.parse(String(sh.getRange(line,4).getValue())||'[]');}catch(_){}}
  if(!Array.isArray(messages))messages=[];
  messages.push({text,by:user.name,email:user.email,at:new Date().toISOString()});
  if(idx<0)sh.getRange(line,1,1,3).setValues([[year,month,key]]);
  sh.getRange(line,4,1,2).setValues([[JSON.stringify(messages.slice(-30)),new Date()]]);
  return {ok:true,messages:wlBddsComments_(ss,year)[month+'|'+key]||[]};
 });
}
/** Compare comment fields independently of the property order used by google.script.run. */
function wlBddsCommentEqual_(left,right){
 if(left===right)return true;
 if(!left||!right||typeof left!=='object'||typeof right!=='object'||Array.isArray(left)!==Array.isArray(right))return false;
 const a=Object.keys(left).sort(),b=Object.keys(right).sort();
 return a.length===b.length&&a.every((key,i)=>key===b[i]&&wlBddsCommentEqual_(left[key],right[key]));
}
/** Delete a specific comment, preserving a recoverable snapshot. */
function wlBddsDeleteComment(year,month,key,expected){
 year=String(year||'');month=wlKey_(month);key=String(key||'').trim();
 if(month.slice(0,4)!==year||!key||!expected||typeof expected!=='object')throw Error('Некорректный комментарий');
 return wlLock_(()=>{
  const ss=wlStore_(),user=wlAdmin_(ss);wlRequireTab_(user,'БДДС',false);
  const sh=ss.getSheetByName('WL_БДДС_Комментарии');if(!sh||sh.getLastRow()<2)throw Error('Комментарий не найден');
  const values=sh.getRange(2,1,sh.getLastRow()-1,4).getValues();
  const zone=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone(),matches=[];
  values.forEach((r,idx)=>{if(String(r[0])!==year||wlMonthCellKey_(r[1],zone)!==month||String(r[2])!==key)return;let messages;try{messages=JSON.parse(String(r[3])||'[]');}catch(_){throw Error('История комментариев повреждена');}if(!Array.isArray(messages))throw Error('История комментариев повреждена');messages.forEach((x,i)=>{if(wlBddsCommentEqual_(x,expected))matches.push({idx,i,messages});});});
  if(matches.length!==1)throw Error('Комментарий изменился или неоднозначен. Обновите отчёт');
  const {idx,i,messages}=matches[0],removed=messages[i],now=new Date(),archive=wlNewSheet_(ss,'WL_БДДС_Удалённые комментарии',['Год','Месяц','Ключ','Комментарий JSON','Удалено','Кто удалил']);
  archive.appendRow([year,month,key,JSON.stringify(removed),now,user.email]);
  messages.splice(i,1);sh.getRange(idx+2,4,1,2).setValues([[JSON.stringify(messages),now]]);
  return {ok:true,messages:wlBddsComments_(ss,year)[month+'|'+key]||[]};
 });
}
/** Fact is explicitly requested and read-only. API schema is validated before displaying. */
function wlBddsFetchFact(year,requestedMonth){
 year=String(year||'');if(!/^\d{4}$/.test(year))throw Error('Неверный год');
 requestedMonth=wlKey_(String(requestedMonth||''));if(!requestedMonth.startsWith(year+'-'))throw Error('Выберите месяц выбранного года');
 const currentMonth=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM');if(requestedMonth>=currentMonth)throw Error('Месяц ещё не завершился. План-факт доступен после окончания месяца.');
 const ss=wlStore_(),user=wlAdmin_(ss);wlRequireTab_(user,'БДДС',false);
 const token=PropertiesService.getScriptProperties().getProperty('WL_DDS_FINTABLO_API_TOKEN');
 if(!token)throw Error('Сначала сохраните ключ ФинТабло');
 const categories=wlFinTabloFetch_(),byId={};
 categories.forEach(c=>{if(c&&c.id)byId[String(c.id)]=String(c.name||c.title||'');});
 const moneybagsResponse=UrlFetchApp.fetch('https://api.fintablo.ru/v1/moneybag',{method:'get',muteHttpExceptions:true,followRedirects:false,headers:{Authorization:'Bearer '+token,Accept:'application/json'}});if(moneybagsResponse.getResponseCode()!==200)throw Error('Не удалось проверить валюты счетов ФинТабло');let bags;try{bags=JSON.parse(moneybagsResponse.getContentText()).items;}catch(_){throw Error('Некорректный ответ по счетам ФинТабло');}if(!Array.isArray(bags))throw Error('Неизвестный формат счетов ФинТабло');const currencies={};bags.forEach(b=>{if(b&&b.id)currencies[String(b.id)]=String(b.currency||'').toUpperCase();});
 const rows=[],seen=new Set(),pageSize=500,mm=requestedMonth.slice(5),lastDay=new Date(Date.UTC(Number(year),Number(mm),0)).getUTCDate();let complete=false;
 for(let page=1;page<=40;page++){
  const resp=UrlFetchApp.fetch('https://api.fintablo.ru/v1/transaction?isPlan=0&dateFrom=01.'+mm+'.'+year+'&dateTo='+lastDay+'.'+mm+'.'+year+'&page='+page+'&pageSize='+pageSize,{method:'get',muteHttpExceptions:true,followRedirects:false,headers:{Authorization:'Bearer '+token,Accept:'application/json'}});
  if(resp.getResponseCode()!==200)throw Error('Не удалось прочитать операции ФинТабло (HTTP '+resp.getResponseCode()+'). Факт не изменён');
  let obj;try{obj=JSON.parse(resp.getContentText());}catch(_){throw Error('Некорректный JSON операций ФинТабло');}
  const list=Array.isArray(obj)?obj:Array.isArray(obj.data)?obj.data:Array.isArray(obj.items)?obj.items:Array.isArray(obj.transactions)?obj.transactions:null;
  if(!list)throw Error('Формат ответа API ФинТабло неизвестен, факт не отображён');
  for(const r of list){if(!r||typeof r!=='object')continue;const id=String(r.id||r.transactionId||'');if(id&&seen.has(id))continue;if(id)seen.add(id);rows.push(r);}
  if(list.length<pageSize||obj.hasNext===false||obj.hasMore===false||(obj.pagination&&obj.pagination.hasNext===false)){complete=true;break;}
 }
 if(!complete)throw Error('Достигнут лимит страниц ФинТабло. Частичный факт не отображается');
 const parentIds=new Set(rows.map(r=>String(r.parentId||'')).filter(Boolean));
 const grouping=wlBddsGrouping_(ss),articles=new Map(),unclassified=[],warnings={ignoredPlan:0,ignoredUnknown:0,ignoredTransfer:0,uncategorized:0,ignoredCurrency:0};
 for(const r of rows){
  if(parentIds.has(String(r.id)))continue;
  const planFlag=r.isPlan!==undefined?r.isPlan:r.isPlanned;
  if(planFlag===true||planFlag===1||planFlag==='1'){warnings.ignoredPlan++;continue;}
  if(planFlag!==undefined&&planFlag!==null&&planFlag!==false&&planFlag!==0&&planFlag!=='0'){warnings.ignoredUnknown++;continue;}
  const group=String(r.group||r.type||'').toLowerCase();if(group!=='income'&&group!=='outcome'){warnings.ignoredTransfer++;continue;}
  const currency=String(r.currency||r.currencyCode||currencies[String(r.moneybagId||'')]||'').toUpperCase();if(currency!=='RUB'&&currency!=='RUR'){warnings.ignoredCurrency++;continue;}
  const rawDate=String(r.date||r.operationDate||r.dateFact||''),match=rawDate.match(/^(\d{2})\.(\d{2})\.(\d{4})$/),month=match?match[3]+'-'+match[2]:rawDate.slice(0,7);if(month!==requestedMonth)continue;
  const amount=Number(r.value!==undefined?r.value:r.amount);if(!Number.isFinite(amount))continue;
  let category=String((r.category&&typeof r.category==='object'?(r.category.name||r.category.title):'')||'').trim();
  const catId=String(r.categoryId||r.category_id||(typeof r.category!=='object'?r.category:'')||'');if(!category)category=byId[catId]||'';
  if(!category){warnings.uncategorized++;unclassified.push({id:String(r.id||r.transactionId||''),date:rawDate,type:group==='income'?'income':'expense',amount:Math.abs(amount),description:String(r.comment||r.description||'').slice(0,500),categoryId:catId});continue;}
  const type=group==='income'?'income':'expense',id=type+'|'+wlBddsNormalize_(category),meta=wlBddsMeta_(type,category,grouping);
  const found=articles.get(id)||{id,category,type,activity:meta.activity,group:meta.group,byMonth:{}};
  found.byMonth[month]=(found.byMonth[month]||0)+Math.abs(amount);articles.set(id,found);
 }
 if(rows.length&&warnings.ignoredUnknown===rows.length)throw Error('В API отсутствует надёжный признак план/факт. Факт не показан');
 return {year,month:requestedMonth,articles:[...articles.values()],unclassified,totalTransactions:rows.length,warnings,readOnly:true};
}
