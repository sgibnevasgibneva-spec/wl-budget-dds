/** Called only by an authenticated administrator after authorizing trigger management. */
function wlMigrateLoginTriggers_(){
 const ss=wlStore_();wlAdmin_(ss);const triggers=ScriptApp.getProjectTriggers(),specs=[{old:'wlMonthlyAuditMaintenance',current:'wlMonthlyAuditMaintenance_',build:b=>b.timeBased().onMonthDay(1).atHour(3)},{old:'wlRunAnnualArchive',current:'wlRunAnnualArchive_',build:b=>b.timeBased().everyDays(1).atHour(3)}],done=[];
 for(const spec of specs){const old=triggers.filter(t=>t.getHandlerFunction()===spec.old),current=triggers.filter(t=>t.getHandlerFunction()===spec.current);if(old.length>1||current.length>1||old.length&&current.length)throw Error('Обнаружены неоднозначные расписания. Перенос автоматических задач остановлен.');if(!old.length)continue;const created=spec.build(ScriptApp.newTrigger(spec.current)).create();try{ScriptApp.deleteTrigger(old[0]);}catch(e){ScriptApp.deleteTrigger(created);throw e;}done.push(spec.current);}
 return {migrated:done};
}
