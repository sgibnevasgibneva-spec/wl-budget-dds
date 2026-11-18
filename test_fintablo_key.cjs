const fs=require('fs'),vm=require('vm'),assert=require('assert');let stored='old-working-key',mode='invalid',writes=0;
const props={getProperty:()=>stored,setProperty:(_k,v)=>{writes++;stored=v}},ctx={PropertiesService:{getScriptProperties:()=>props},wlStore_:()=>({}),wlAdmin_:()=>{},UrlFetchApp:{fetch:()=>({getResponseCode:()=>mode==='valid'?200:401,getContentText:()=>JSON.stringify(mode==='valid'?{items:[{id:1,name:'test'}]}:{reason:mode})})}};
vm.createContext(ctx);vm.runInContext(fs.readFileSync('FinTabloSync.js','utf8'),ctx);
assert.throws(()=>ctx.wlFinTabloSetToken_('new-invalid-key'),/отклонил/);assert.equal(stored,'old-working-key');assert.equal(writes,0);
mode='legacy_key_expired';assert.throws(()=>ctx.wlFinTabloFetch_(),/устарел/);mode='key_reissued';assert.throws(()=>ctx.wlFinTabloFetch_(),/перевыпущен/);
mode='valid';ctx.wlFinTabloSetToken_('new-valid-key');assert.equal(stored,'new-valid-key');assert.equal(writes,1);const info=ctx.wlFinTabloKeyInfo_();assert.equal(info.lastFour,'-key');assert(!JSON.stringify(info).includes(stored));
console.log('Invalid keys are rejected without replacing the existing key; specific 401 reasons are readable');
