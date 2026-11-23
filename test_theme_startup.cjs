const fs=require('fs'),vm=require('vm'),assert=require('assert'),html=fs.readFileSync('Index.html','utf8');
const source=html.slice(html.indexOf('function initTheme(){'),html.indexOf('function start(){'));
const el=()=>({setAttribute(k,v){this[k]=v;}}),buttons={themeToggle:el(),loginThemeToggle:el()},store={};
const context={document:{documentElement:{dataset:{}}},localStorage:{getItem:k=>store[k],setItem:(k,v)=>store[k]=v},$:id=>buttons[id]};vm.createContext(context);vm.runInContext(source,context);context.initTheme();
assert.equal(context.document.documentElement.dataset.theme,'light');buttons.themeToggle.onclick();assert.equal(store.wl_theme,'dark');assert.equal(buttons.loginThemeToggle['aria-pressed'],'true');context.initTheme();assert.equal(context.document.documentElement.dataset.theme,'dark');buttons.loginThemeToggle.onclick();assert.equal(store.wl_theme,'light');
const summary=html.slice(html.indexOf('function renderDeptSummary(){'),html.indexOf('function setDeptMode('));let calls=0;
const c={deptMode:'expenses',grouped(){calls++},groupedSections(){calls++;return[]},rows:[],dept:'HR',$:()=>({})};vm.createContext(c);vm.runInContext(summary,c);c.renderDeptSummary();assert.equal(calls,0);c.deptMode='summary';c.renderDeptSummary();assert.equal(calls,2);
console.log('Theme persists between sessions; hidden department summary skips work');
