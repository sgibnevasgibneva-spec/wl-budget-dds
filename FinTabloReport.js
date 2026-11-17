/** FinTablo's official MCP is read-only; request only the finished cash-flow report. */
function wlFtReportClient_(token){
 const endpoint='https://my.fintablo.ru/mcp',base={Authorization:'Bearer '+token,Accept:'application/json, text/event-stream'};
 let session='';
 function send(body){
  const headers=Object.assign({},base);if(session)headers['Mcp-Session-Id']=session;
  return UrlFetchApp.fetch(endpoint,{method:'post',contentType:'application/json',headers,payload:JSON.stringify(body),muteHttpExceptions:true,followRedirects:false});
 }
 const init=send({jsonrpc:'2.0',id:1,method:'initialize',params:{protocolVersion:'2025-03-26',capabilities:{},clientInfo:{name:'wl-budget-dds',version:'1.0'}}});
 if(init.getResponseCode()!==200)throw Error('Не удалось подключить отчёт ФинТабло (HTTP '+init.getResponseCode()+')');
 const headers=init.getAllHeaders();session=headers['Mcp-Session-Id']||headers['mcp-session-id']||'';
 const ready=send({jsonrpc:'2.0',method:'notifications/initialized'});if(![200,202,204].includes(ready.getResponseCode()))throw Error('Не удалось подготовить чтение отчёта ФинТабло');
 return function(month,group,direction){
  const y=Number(month.slice(0,4)),m=Number(month.slice(5)),last=new Date(Date.UTC(y,m,0)).getUTCDate();
  const args={fromDate:month+'-01',toDate:month+'-'+String(last).padStart(2,'0'),group,dateGroup:'month'};
  if(direction!==undefined)args.direction=[direction];
  const response=send({jsonrpc:'2.0',id:2,method:'tools/call',params:{name:'dds_report',arguments:args}});
  if(response.getResponseCode()!==200)throw Error('ФинТабло не вернул отчёт (HTTP '+response.getResponseCode()+')');
  let envelope,data;try{envelope=JSON.parse(response.getContentText());if(envelope.error||!envelope.result||envelope.result.isError)throw Error();data=JSON.parse(envelope.result.content[0].text);}catch(_){throw Error('Не удалось прочитать готовый отчёт ФинТабло');}
  if(!data||data.fromDate!==args.fromDate||data.toDate!==args.toDate||!Array.isArray(data.periods)||data.periods.length!==1||!Array.isArray(data.periods[0].groups))throw Error('Отчёт ФинТабло вернул другой период или неизвестный формат');
  return data.periods[0].groups;
 };
}
function wlFtReportLeaves_(rows){
 const out=[];function visit(items,depth){if(depth>8)throw Error('Неизвестная структура отчёта ФинТабло');for(const row of items){if(!row||typeof row!=='object')throw Error('Некорректная строка отчёта');if(Array.isArray(row.items)&&row.items.length)visit(row.items,depth+1);else{if(!row.amount||typeof row.amount!=='object'||Object.keys(row.amount).some(c=>c!=='RUB'&&Number(row.amount[c])!==0)||!Number.isFinite(Number(row.amount.RUB)))throw Error('В отчёте есть суммы в другой или неизвестной валюте');out.push(row);}}}visit(rows,0);return out;
}
function wlFtReportArticles_(groups,candidates,month,grouping){
 const byId=new Map(candidates.map(c=>[String(c.id),c])),out=new Map(),ids=new Map();
 for(const row of wlFtReportLeaves_(groups)){
  const c=byId.get(String(row.id));if(!c||!['income','outcome'].includes(c.group))throw Error('Статья отчёта не найдена в справочнике ФинТабло: '+String(row.name||row.id));
  const type=c.group==='income'?'income':'expense',category=String(c.name||row.name||'').trim(),id=type+'|'+wlBddsNormalize_(category);
  if(ids.has(id)&&ids.get(id)!==String(row.id))throw Error('Одинаковые названия статей ФинТабло требуют сопоставления: '+category);
  ids.set(id,String(row.id));const meta=wlBddsMeta_(type,category,grouping),value=Number(row.amount.RUB)*(type==='expense'?-1:1);
  const item=out.get(id)||{id,category,type,activity:meta.activity,group:meta.group,byMonth:{}};item.byMonth[month]=(item.byMonth[month]||0)+value;out.set(id,item);
 }
 return [...out.values()];
}
