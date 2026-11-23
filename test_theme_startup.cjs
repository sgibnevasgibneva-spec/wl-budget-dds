const fs=require('fs'),vm=require('vm'),assert=require('assert'),html=fs.readFileSync('Index.html','utf8');
assert(!html.includes('themeToggle'));assert(!html.includes('data-theme="dark"'));assert(!html.includes('initTheme'));
const summary=html.slice(html.indexOf('function renderDeptSummary(){'),html.indexOf('function setDeptMode('));let calls=0;
const c={deptMode:'expenses',grouped(){calls++},groupedSections(){calls++;return[]},rows:[],dept:'HR',$:()=>({})};vm.createContext(c);vm.runInContext(summary,c);c.renderDeptSummary();assert.equal(calls,0);c.deptMode='summary';c.renderDeptSummary();assert.equal(calls,2);
console.log('Dark theme removed; hidden department summary skips work');
