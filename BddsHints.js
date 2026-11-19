/** An explicit empty hint is a saved dismissal. */
function wlBddsSetHint_(year,month,key,text){
 year=String(year||'');month=wlKey_(month);key=String(key||'').trim();text=String(text||'').trim();
 if(month.slice(0,4)!==year||!key||key.length>250||text.length>1600)throw Error('Некорректная подсказка');
 return wlLock_(()=>{const ss=wlStore_(),user=wlAdmin_(ss);wlRequireTab_(user,'БДДС',false);const sh=wlNewSheet_(ss,'WL_БДДС_Комментарии',['Год','Месяц','Ключ показателя','Сообщения JSON','Изменено']),zone=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone(),rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,4).getValues():[],matching=[];
 rows.forEach((r,i)=>{if(String(r[0])===year&&wlMonthCellKey_(r[1],zone)===month&&String(r[2])===key)matching.push(i);});if(matching.length>1)throw Error('Найдено несколько записей комментариев. Обновите отчёт');const idx=matching.length?matching[0]:-1,line=idx<0?sh.getLastRow()+1:idx+2;let notes=[];if(idx>=0){try{notes=JSON.parse(String(rows[idx][3]||'[]'));}catch(_){throw Error('История комментариев повреждена');}if(!Array.isArray(notes))throw Error('История комментариев повреждена');}
 if(idx>=0)notes=wlBddsKeepSourceDismissal_(key,notes);notes=notes.filter(n=>n.kind!=='hint');notes.push({kind:'hint',text,by:user.name,email:user.email,at:new Date().toISOString()});if(idx<0)sh.getRange(line,1,1,3).setValues([[year,month,key]]);sh.getRange(line,4,1,2).setValues([[JSON.stringify(notes),new Date()]]);return {messages:wlBddsComments_(ss,year)[month+'|'+key]||[]};});
}
function wlBddsEditComment_(year,month,key,expected,text){
 year=String(year||'');month=wlKey_(month);key=String(key||'').trim();text=String(text||'').trim();if(month.slice(0,4)!==year||!key||!expected||!text||text.length>1600)throw Error('Некорректный комментарий');
 return wlLock_(()=>{const ss=wlStore_(),user=wlAdmin_(ss);wlRequireTab_(user,'БДДС',false);const sh=ss.getSheetByName('WL_БДДС_Комментарии');if(!sh||sh.getLastRow()<2)throw Error('Комментарий не найден');const rows=sh.getRange(2,1,sh.getLastRow()-1,4).getValues(),zone=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone(),matches=[];
 rows.forEach((r,idx)=>{if(String(r[0])!==year||wlMonthCellKey_(r[1],zone)!==month||String(r[2])!==key)return;let items;try{items=JSON.parse(String(r[3])||'[]');}catch(_){throw Error('История комментариев повреждена');}items.forEach((x,i)=>{if(wlBddsCommentEqual_(x,expected))matches.push({idx,i,items});});});if(matches.length!==1)throw Error('Комментарий изменился или неоднозначен. Обновите отчёт');const {idx,i,items}=matches[0],now=new Date(),backup=wlNewSheet_(ss,'WL_БДДС_Удалённые комментарии',['Год','Месяц','Ключ','Комментарий JSON','Удалено','Кто удалил']);backup.appendRow([year,month,key,JSON.stringify(items[i]),now,user.email]);items[i]={...items[i],text,editedAt:now.toISOString(),editedBy:user.email};sh.getRange(idx+2,4,1,2).setValues([[JSON.stringify(items),now]]);return {messages:wlBddsComments_(ss,year)[month+'|'+key]||[]};});
}

/** Empty legacy history threads mean that the user already cleared the comment. */
function wlBddsKeepSourceDismissal_(key,messages){
 return key.startsWith('history|')&&!messages.length?[{kind:'history-source',text:''}]:messages;
}
/** Source overrides and dismissed hints must survive the ordinary comment limit. */
function wlBddsLimitComments_(messages){
 const controls=new Map(),ordinary=[];
 messages.forEach(n=>{if(n.kind==='history-source'||n.kind==='hint')controls.set(n.kind,n);else ordinary.push(n);});
 return [...controls.values(),...ordinary.slice(-30)];
}
/** Change the app's copy of a source comment; never write to the historical workbook. */
function wlBddsSetSourceComment_(year,month,key,expected,text){
 year=String(year||'');month=wlKey_(month);key=String(key||'').trim();
 if(month.slice(0,4)!==year||!/^history\|.+\|[1-9]\d*$/.test(key)||key.length>250||typeof text!=='string'||text.length>1600)throw Error('Некорректный комментарий');
 text=text.trim();
 return wlLock_(()=>{
  const ss=wlStore_(),user=wlAdmin_(ss);wlRequireTab_(user,'БДДС',false);
  const sh=wlNewSheet_(ss,'WL_БДДС_Комментарии',['Год','Месяц','Ключ показателя','Сообщения JSON','Изменено']),zone=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone(),rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,4).getValues():[],matches=[];
  rows.forEach((r,i)=>{if(String(r[0])===year&&wlMonthCellKey_(r[1],zone)===month&&String(r[2])===key)matches.push(i);});
  if(matches.length>1)throw Error('Найдено несколько записей комментариев. Обновите отчёт');
  const idx=matches.length?matches[0]:-1;let messages=[];
  if(idx>=0){try{messages=JSON.parse(String(rows[idx][3]||'[]'));}catch(_){throw Error('История комментариев повреждена');}if(!Array.isArray(messages))throw Error('История комментариев повреждена');messages=wlBddsKeepSourceDismissal_(key,messages);}
  const current=messages.find(n=>n.kind==='history-source')||null;
  if(!wlBddsCommentEqual_(current,expected||null))throw Error('Комментарий изменился. Обновите отчёт');
  const now=new Date();
  if(current){const backup=wlNewSheet_(ss,'WL_БДДС_Удалённые комментарии',['Год','Месяц','Ключ','Комментарий JSON','Удалено','Кто удалил']);backup.appendRow([year,month,key,JSON.stringify(current),now,user.email]);}
  messages=messages.filter(n=>n.kind!=='history-source');messages.push({kind:'history-source',text,by:user.name,email:user.email,at:now.toISOString()});
  const line=idx<0?sh.getLastRow()+1:idx+2;
  if(idx<0)sh.getRange(line,1,1,3).setValues([[year,month,key]]);
  sh.getRange(line,4,1,2).setValues([[JSON.stringify(messages),now]]);
  return {messages:wlBddsComments_(ss,year)[month+'|'+key]||[]};
 });
}
