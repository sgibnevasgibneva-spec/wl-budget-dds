const fs=require('fs'),vm=require('vm'),assert=require('assert'),server=fs.readFileSync('Code.js','utf8'),client=fs.readFileSync('Index.html','utf8');
const sc={Date};vm.createContext(sc);vm.runInContext(server.slice(server.indexOf('const WL_WORK_CALENDAR='),server.indexOf('function wlBoot_('))+'\nthis.rules=WL_WORK_CALENDAR;',sc);
const cc={Date,boot:{workCalendar:sc.rules}};vm.createContext(cc);vm.runInContext(client.slice(client.indexOf('function calendarPaymentDay('),client.indexOf('function canEdit(')),cc);
for(const year of [2025,2026,2027]){let working=0;for(let month=1;month<=12;month++){const key=year+'-'+String(month).padStart(2,'0');for(let day=1;day<=new Date(Date.UTC(year,month,0)).getUTCDate();day++){const s=sc.wlCalendarDay_(sc.rules,key,day);assert.equal(s,cc.calendarPaymentDay(key,day));if(s==='work')working++;}}assert.equal(working,247,String(year));}
for(const [key,day,status] of [['2026-03',9,'off'],['2026-12',31,'off'],['2027-02',20,'work'],['2027-02',22,'off'],['2027-11',5,'off'],['2025-11',1,'work'],['2025-05',8,'off'],['2028-01',3,'unknown']])assert.equal(sc.wlCalendarDay_(sc.rules,key,day),status,key+'/'+day);
assert.throws(()=>sc.wlValidatePaymentDay_('2026-11',4,0,500),/нерабочий день/);
assert.doesNotThrow(()=>sc.wlValidatePaymentDay_('2026-11',4,500,500));assert.doesNotThrow(()=>sc.wlValidatePaymentDay_('2026-11',4,500,''));assert.throws(()=>sc.wlValidatePaymentDay_('2026-11',4,500,501),/нерабочий день/);
assert.doesNotThrow(()=>sc.wlValidatePaymentDay_('2027-02',20,0,500));assert.throws(()=>sc.wlValidatePaymentDay_('2028-01',3,0,500),/не подключён/);
assert.throws(()=>sc.wlValidatePaymentAmounts_('2026-11',{}, {4:500}),/нерабочий день/);assert.doesNotThrow(()=>sc.wlValidatePaymentAmounts_('2026-11',{4:500},{4:500,3:200}));
const old={id:'existing1',department:'HR',payment:'',source:'',category:'Зарплата',description:'До изменений',amounts:{4:777},cashFlowType:'expense'};
const ctx={Date,WL_WORK_CALENDAR:sc.rules,wlCalendarDay_:sc.wlCalendarDay_,wlValidatePaymentDay_:sc.wlValidatePaymentDay_,wlValidatePaymentAmounts_:sc.wlValidatePaymentAmounts_,wlKey_:x=>x,wlStore_:()=>({}),wlAdmin_:()=>{},wlFindMonth_:()=>({status:'Открыт'}),wlDictionaries_:()=>({departments:[{name:'HR'}],payments:[],sources:[],categories:['Зарплата'],incomeCategories:[]}),wlOps_:()=>({}),wlOpsRead_:()=>[{line:2,row:old}],wlDeletedIdsForMonth_:()=>new Set(),wlNormalizeDept_:x=>x,wlExcelFingerprint_:(_key,row)=>row.description==='До изменений'?'oldvalue':'newvalue'};
vm.createContext(ctx);vm.runInContext(server.slice(server.indexOf('function wlExcelRoundTripCore_('),server.indexOf('/** ЦФО/БДДС:',server.indexOf('function wlExcelRoundTripCore_('))),ctx);
const entry={...old,baseSig:'oldvalue',description:'После изменений'};
assert.equal(ctx.wlPreviewExcelRoundTrip_('2026-11',[entry]).updated,1);
assert.throws(()=>ctx.wlPreviewExcelRoundTrip_('2026-11',[{...entry,amounts:{4:778}}]),/нерабочий день/);
assert.throws(()=>ctx.wlPreviewExcelRoundTrip_('2026-11',[{...entry,id:'imp-newrow',baseSig:'',amounts:{4:778}}]),/нерабочий день/);
console.log('Official calendars match 247 work days in 2025–2027; working Saturdays, transfers, existing payments, imports and unknown years verified');
