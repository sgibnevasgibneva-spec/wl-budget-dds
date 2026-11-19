function wlRpc(token,name,args){
 const dispatch=()=>{
 const requestStarted=Date.now(),email=wlAuthSession_(token),authenticated=Date.now();if(!Array.isArray(args)||args.length>10)throw Error('Некорректный запрос');
 const methods={
 "wlStartup":wlStartup_,
 "wlMigrateLoginTriggers":wlMigrateLoginTriggers_,
"wlGenerateLoginCode":wlGenerateLoginCode_,
"wlApplyExcelRoundTrip":wlApplyExcelRoundTrip_,
"wlBddsDeleteComment":wlBddsDeleteComment_,
"wlBddsEditComment":wlBddsEditComment_,
"wlBddsFetchFact":wlBddsFetchFact_,
"wlBddsHistorical2026":wlBddsHistorical2026_,
"wlBddsPostComment":wlBddsPostComment_,
"wlBddsSavedFacts":wlBddsSavedFacts_,
"wlBddsSetFund":wlBddsSetFund_,
"wlBddsSetHint":wlBddsSetHint_,
"wlBoot":wlBoot_,
"wlCfoGetFunds":wlCfoGetFunds_,
"wlDeleteRow":wlDeleteRow_,
"wlDeleteRows":wlDeleteRows_,
"wlDictionaryAdd":wlDictionaryAdd_,
"wlDictionaryRemove":wlDictionaryRemove_,
"wlDictionaryRename":wlDictionaryRename_,
"wlEnableAnnualArchive":wlEnableAnnualArchive_,
"wlEnableMonthlyAuditArchive":wlEnableMonthlyAuditArchive_,
"wlFinTabloApplySelected":wlFinTabloApplySelected_,
"wlFinTabloBudgetPreview":wlFinTabloBudgetPreview_,
"wlFinTabloBudgetSend":wlFinTabloBudgetSend_,
"wlFinTabloPreview":wlFinTabloPreview_,
"wlFinTabloSetToken":wlFinTabloSetToken_,
"wlFinTabloKeyInfo":wlFinTabloKeyInfo_,
"wlGetBddsPlan":wlGetBddsPlan_,
"wlGetMonth":wlGetMonth_,
"wlGetSettings":wlGetSettings_,
"wlMonthManage":wlMonthManage_,
"wlMoveBudgetRow":wlMoveBudgetRow_,
"wlPatchChanges":wlPatchChanges_,
"wlPreviewExcelRoundTrip":wlPreviewExcelRoundTrip_,
"wlRemoveAccess":wlRemoveAccess_,
"wlReviewList":wlReviewList_,
"wlReviewResolve":wlReviewResolve_,
"wlSaveAccess":wlSaveAccess_,
"wlSetArticleGroup":wlSetArticleGroup_,
"wlTrashList":wlTrashList_,
"wlTrashPurge":wlTrashPurge_,
"wlTrashRestore":wlTrashRestore_,
"wlTrashRestoreMany":wlTrashRestoreMany_,
"wlYearArchiveSnapshot":wlYearArchiveSnapshot_,
"wlYearArchiveStatus":wlYearArchiveStatus_,
"wlYearManage":wlYearManage_
 };if(!Object.prototype.hasOwnProperty.call(methods,name))throw Error('Неизвестная команда');
 WL_AUTH_EMAIL=email;try{const result=methods[name](...args);if(name==='wlStartup')result.timing={...result.timing,authMs:authenticated-requestStarted,requestServerMs:Date.now()-requestStarted};return result;}finally{WL_AUTH_EMAIL='';}
 };
 if(name==='wlStartup'){
  const started=Date.now(),fast=wlTryStartupBatch_(token,dispatch);if(fast)return fast;
  const result=wlReadSnapshot_(dispatch);result.timing={...result.timing,readMode:'regular',fallbackReason:WL_STARTUP_BATCH_FALLBACK,requestServerMs:Date.now()-started};return result;
 }
 return dispatch();
}
