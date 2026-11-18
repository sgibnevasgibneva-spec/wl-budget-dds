/** Persist finished FinTablo reports in our own spreadsheet; each refresh is a separate snapshot. */
function wlBddsStoreFact_(ss,result){
 const text=JSON.stringify(result),chunks=[];for(let i=0;i<text.length;){let end=Math.min(i+20000,text.length);if(end<text.length&&/[\uD800-\uDBFF]/.test(text.charAt(end-1)))end--;chunks.push(text.slice(i,end));i=end;}
 if(!chunks.length||chunks.length>100)throw Error('Отчёт слишком большой для сохранения');
 return wlLock_(()=>{const sh=wlNewSheet_(ss,'WL_ПланФакт_Снимки',['Месяц','Снимок','Часть','Всего частей','Данные JSON','Получен']);const id=Utilities.getUuid(),now=new Date(),rows=chunks.map((part,i)=>[result.month,id,i+1,chunks.length,part,now]);sh.getRange(sh.getLastRow()+1,1,rows.length,6).setValues(rows);return true;});
}
function wlBddsReadFactSnapshots_(rows,year,zone){
 const versions=new Map();for(const row of rows){const month=wlMonthCellKey_(row[0],zone),id=String(row[1]||''),part=Number(row[2]),count=Number(row[3]);if(!month||!month.startsWith(year+'-')||!id||!Number.isInteger(part)||!Number.isInteger(count)||part<1||part>count||count>100)continue;const key=month+'|'+id,v=versions.get(key)||{month,count,parts:[],at:row[5]};if(v.count!==count)continue;v.parts[part-1]=String(row[4]||'');versions.set(key,v);}
 const latest=new Map();for(const v of versions.values()){if(v.parts.length!==v.count||Array.from({length:v.count},(_,i)=>v.parts[i]).some(x=>!x))continue;let report;try{report=JSON.parse(v.parts.join(''));}catch(_){continue;}if(report.month!==v.month||report.year!==year||!Array.isArray(report.articles))continue;latest.set(v.month,report);}
 return [...latest.values()].sort((a,b)=>a.month.localeCompare(b.month));
}
function wlBddsSavedFacts_(year){
 year=String(year||'');if(!/^\d{4}$/.test(year))throw Error('Неверный год');const ss=wlStore_(),user=wlAdmin_(ss);wlRequireTab_(user,'БДДС',false);const sh=ss.getSheetByName('WL_ПланФакт_Снимки');if(!sh||sh.getLastRow()<2)return [];return wlBddsReadFactSnapshots_(sh.getRange(2,1,sh.getLastRow()-1,6).getValues(),year,ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone());
}
