const fs=require('fs'),vm=require('vm'),assert=require('assert');const html=fs.readFileSync('Index.html','utf8'),start=html.indexOf('function bddsDepartmentExplanation('),end=html.indexOf('function bddsOpenComment(',start),ctx={money:x=>String(x)};vm.createContext(ctx);vm.runInContext(html.slice(start,end),ctx);
let plan={byDepartment:{'Отдел продаж':{'2026-11':100}}},fact={byMonth:{'2026-11':130}};
let out=ctx.bddsDepartmentExplanation(plan,fact,'2026-11',false,true).join(' ');assert(out.includes('Отдел продаж'));assert(out.includes('больше на 30'));assert(!out.includes('Возможные причины'));
fact.byMonth['2026-11']=70;out=ctx.bddsDepartmentExplanation(plan,fact,'2026-11',false,true).join(' ');assert(out.includes('меньше на 30'));
plan.byDepartment.IT={'2026-11':20};out=ctx.bddsDepartmentExplanation(plan,fact,'2026-11',false,true).join(' ');assert(out.includes('нескольких отделах'));assert(!out.includes('Отдел продаж: план'));
plan={byDepartment:{Финансы:{'2026-11':100}}};out=ctx.bddsDepartmentExplanation(plan,fact,'2026-10',true,true).join(' ');assert(out.includes('План отдела за выбранный месяц отсутствует'));assert(!out.includes('больше на'));
console.log('Department explanations use only supported plan and fact and do not invent causes');
