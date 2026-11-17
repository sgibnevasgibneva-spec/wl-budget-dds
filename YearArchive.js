/** Annual archive (Google Sheets import-compatible snapshot). No source operations are deleted. */
const WL_YEAR_ARCHIVE_LINK_PREFIX='WL_DDS_YEAR_ARCHIVE_FILE_';

function wlYearArchiveSnapshot(year,replaceExisting){
  return wlLock_(()=>{const ss=wlStore_(),u=wlAdmin_(ss),value=String(year||''),result=wlCreateYearArchive_(value,!!replaceExisting,ss);
  if(value==='2026')wlArchiveHistoricalBdds2026_(result);
  PropertiesService.getScriptProperties().setProperty('WL_DDS_BDDS_ARCHIVED_'+value,'true');
  return {ok:true,url:result.url,rows:result.rows,year:result.year,alreadyExisted:result.alreadyExisted};});
}
function wlArchiveHistoricalBdds2026_(result){
 const destination=SpreadsheetApp.openByUrl(result.url);if(destination.getSheetByName('БДДС 2026'))return;
 const original=SpreadsheetApp.openById('1aVg5VqQkp1WSIeHYuK-SVrR5DRw22JqDy0qGhbN9Xkg').getSheetByName('БДДС');if(!original)throw Error('Не найден исходный БДДС 2026. Год не перенесён в архив.');
 const values=original.getDataRange().getDisplayValues();if(values.length<2||values[0].length<6)throw Error('Не удалось прочитать исторический БДДС 2026. Год не перенесён в архив.');
 const target=destination.insertSheet('БДДС 2026');try{if(target.getMaxRows()<values.length)target.insertRowsAfter(target.getMaxRows(),values.length-target.getMaxRows());if(target.getMaxColumns()<values[0].length)target.insertColumnsAfter(target.getMaxColumns(),values[0].length-target.getMaxColumns());const safe=values.map(row=>row.map(value=>/^[=+@-]/.test(value)?"'"+value:value));target.getRange(1,1,safe.length,safe[0].length).setNumberFormat('@').setValues(safe);target.setFrozenRows(1);target.setFrozenColumns(1);}catch(e){destination.deleteSheet(target);throw e;}
}
function wlCreateYearArchive_(year,overwrite,source){
  if(!/^(20|21)\d{2}$/.test(year))throw Error('Неверный год');
  const todayYear=Number(Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy'));
  if(Number(year)>=todayYear)throw Error('Архивировать можно только завершённый календарный год');
  const props=PropertiesService.getScriptProperties();
  const prop=WL_YEAR_ARCHIVE_LINK_PREFIX+year;
  const existing=props.getProperty(prop);
  if(existing&&!overwrite){
    const parsed=JSON.parse(existing);
    return {year,url:parsed.url,rows:parsed.rows,alreadyExisted:true};
  }
  // Short lock for idempotency. Data is read in one bulk call; no business rows are changed.
  const ss=source||wlStore_(),all=wlOpsRead_(wlOps_(ss)).filter(e=>String(e.key||'').startsWith(year+'-'));
  const rows=all.map(e=>e.row||{}),headers=['Месяц','Отдел','Как платим','Откуда платим','Статья ДДС','Описание/название затрат',...Array.from({length:31},(_,i)=>String(i+1).padStart(2,'0'))];
  const data=all.map(e=>{
    const r=e.row||{},amounts=r.amounts||{},values=Array.from({length:31},(_,i)=>{
      const raw=amounts[i+1],n=Number(raw);
      return raw===undefined||raw===null||raw===''?'':Number.isFinite(n)?n:'';
    });
    const safe=s=>{let str=String(s||'');return /^[=+@-]/.test(str)?"'"+str:str;};
    return [String(e.key),safe(r.department),safe(r.payment),safe(r.source),safe(r.category),safe(r.description),...values];
  });
  const document=SpreadsheetApp.create('WL DDS — Архив '+year+' — '+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH.mm'));
  const sheet=document.getSheets()[0];sheet.setName('Импорт');
  if(sheet.getMaxColumns()<headers.length)sheet.insertColumnsAfter(sheet.getMaxColumns(),headers.length-sheet.getMaxColumns());
  if(sheet.getMaxRows()<data.length+1)sheet.insertRowsAfter(sheet.getMaxRows(),data.length+1-sheet.getMaxRows());
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  if(data.length){
    sheet.getRange(2,1,data.length,1).setNumberFormat('@');
    sheet.getRange(2,1,data.length,headers.length).setValues(data);
    sheet.getRange(2,7,data.length,31).setNumberFormat('#,##0.00');
  }
  sheet.getRange(1,1,1,headers.length).setBackground('#dfeaf6').setFontWeight('bold');
  sheet.setFrozenRows(1);sheet.setFrozenColumns(2);
  [145,170,140,160,230,350].forEach((width,i)=>sheet.setColumnWidth(i+1,width));
  for(let i=7;i<=37;i++)sheet.setColumnWidth(i,105);
  const meta=document.insertSheet('Метаданные и комментарии');
  if(meta.getMaxRows()<all.length+1)meta.insertRowsAfter(meta.getMaxRows(),all.length+1-meta.getMaxRows());
  const metaHeader=['Месяц','ID строки','Отдел','Тип операции','Статья ДДС','Описание','Комментарии JSON','Исходная строка JSON'];
  meta.getRange(1,1,1,metaHeader.length).setValues([metaHeader]).setFontWeight('bold').setBackground('#dfeaf6');
  if(all.length)meta.getRange(2,1,all.length,metaHeader.length).setValues(all.map(e=>{
    const r=e.row||{},safe=s=>{const t=String(s||'');return /^[=+@-]/.test(t)?"'"+t:t;};
    return [String(e.key),safe(r.id),safe(r.department),safe(r.cashFlowType),safe(r.category),safe(r.description),JSON.stringify(r.notes||{}),JSON.stringify(r)];
  }));
  meta.setFrozenRows(1);meta.setColumnWidth(7,370);meta.setColumnWidth(8,370);
  const about=document.insertSheet('О файле');
  about.getRange(1,1,5,2).setValues([
    ['Назначение','Независимый архив годового бюджета ДДС, не изменяет рабочие данные'],
    ['Год',year],['Создан',new Date()],['Количество строк',all.length],['Источник',ss.getUrl()]
  ]);
  about.setColumnWidth(1,220);about.setColumnWidth(2,550);
  props.setProperty(prop,JSON.stringify({url:document.getUrl(),id:document.getId(),rows:all.length,created:new Date().toISOString()}));
  return {year,url:document.getUrl(),rows:all.length,alreadyExisted:false};
}
/**
 * Один раз запустить вручную из редактора Apps Script под аккаунтом администратора.
 * Google показывает отдельный экран разрешения на управление триггерами.
 * Эта функция ничего не создаёт и не меняет в бюджетных данных.
 */
function wlAuthorizeYearArchive(){
  const ss=wlStore_();wlAdmin_(ss);
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL,[
    'https://www.googleapis.com/auth/script.scriptapp'
  ]);
  const count=ScriptApp.getProjectTriggers().length;
  Logger.log('Разрешение управления триггерами подтверждено. Триггеров у пользователя: '+count);
  return {ok:true,triggerCount:count};
}
/** Only admin may explicitly install the annual trigger; never run implicitly. */
function wlEnableAnnualArchive(){
  const ss=wlStore_();wlAdmin_(ss);
  const handler='wlRunAnnualArchive';
  try{
    if(!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()===handler))
      ScriptApp.newTrigger(handler).timeBased().everyDays(1).atHour(3).create();
  }catch(e){
    if(wlYearArchiveMissingTriggerScope_(e))throw Error('Для автоархива требуется разрешение Google: в редакторе Apps Script один раз выполните wlAuthorizeYearArchive, подтвердите доступ, затем вернитесь и включите автоархивирование.');
    throw e;
  }
  return {enabled:true};
}
/** Google can withhold script.scriptapp even when the web app reads spreadsheets. */
function wlYearArchiveMissingTriggerScope_(e){
  return /script\.scriptapp|permission|authorization|scope|разрешени|авторизац/i.test(String(e&&e.message||e));
}
function wlRunAnnualArchive(){
  // Trigger runs as its creator. Verify identity against the initial project owner.
  const runAs=String(Session.getEffectiveUser().getEmail()||'').toLowerCase();
  if(!runAs)throw Error('Google не передал email исполнителя автоархива');
  const ss=wlStore_();
  if(!wlRules_(ss).some(r=>r.email===runAs&&r.enabled&&r.role==='admin'))throw Error('Автоархив может выполняться только от имени администратора');
  const prev=Number(Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy'))-1;
  if(!wlMonthList_(ss).some(m=>m.key.startsWith(String(prev)+'-')))return;
  wlCreateYearArchive_(String(prev),false,ss);
}
function wlYearArchiveStatus(){
  const ss=wlStore_();wlAdmin_(ss);
  const props=PropertiesService.getScriptProperties(),years=[...new Set(wlMonthList_(ss).map(m=>m.key.slice(0,4)))];
  const archives=years.map(y=>({year:y,detail:props.getProperty(WL_YEAR_ARCHIVE_LINK_PREFIX+y)||''}));
  // Список готовых архивов работает даже если ещё не выдано разрешение на триггеры.
  // Не подменяем неизвестное состояние словом «выключено».
  try{
    return {archives,automationEnabled:ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='wlRunAnnualArchive'),authorizationRequired:false};
  }catch(e){
    if(!wlYearArchiveMissingTriggerScope_(e))throw e;
    return {archives,automationEnabled:null,authorizationRequired:true};
  }
}
