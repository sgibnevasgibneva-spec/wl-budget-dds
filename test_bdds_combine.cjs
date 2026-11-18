const fs=require('fs'),vm=require('vm'),assert=require('assert');
class Node {constructor(tag='td',text=''){this.tag=tag;this.textContent=text;this.children=[];this.dataset={};this.className='';this._html='';}appendChild(n){this.children.push(n);return n}get cells(){return this.children}get rows(){return this.children}set innerHTML(v){this._html=v;this.children=[]}get innerHTML(){return this._html}cloneNode(deep){const n=new Node(this.tag,this.textContent);n.dataset={...this.dataset};n.className=this.className;if(deep)this.children.forEach(c=>n.appendChild(c.cloneNode(true)));return n}querySelectorAll(){const walk=n=>[...(n.dataset.expandMonth?[n]:[]),...n.children.flatMap(walk)];return walk(this)}}
const el={}, E=(id)=>el[id]=new Node(id), row=(...cells)=>{const n=new Node('tr');cells.forEach(c=>n.appendChild(typeof c==='string'?new Node('td',c):c));return n},month=(name)=>{const n=new Node('th',name);n.dataset.expandMonth=name;return n};
for(const id of ['bddsCombinedHead','bddsCombinedBody','bddsHead','bddsBody','bddsHistoryHead','bddsHistoryBody','bddsFoldMonths'])E(id);
el.bddsHead.appendChild(row('Показатель',month('2026-11')));el.bddsHead.appendChild(row('План'));
el.bddsHistoryHead.appendChild(row('Показатель',month('2026-10')));el.bddsHistoryHead.appendChild(row('План'));
el.bddsBody.appendChild(row('Продажи','200'));el.bddsBody.appendChild(row('Поступления','250'));
el.bddsHistoryBody.appendChild(row('Продажи','100'));el.bddsHistoryBody.appendChild(row('Поступления','80'));el.bddsHistoryBody.appendChild(row('Поступления','20'));el.bddsHistoryBody.appendChild(row('Старая статья','30'));
const s=fs.readFileSync('Index.html','utf8'),src=s.slice(s.indexOf('function bddsCombineTables(){'),s.indexOf('function bddsDscr('));const ctx={document:{createElement:t=>new Node(t)},$:id=>el[id],bddsHistoryMode:true,bddsHistoryData:{},bddsExpandedMonths:new Set(),bddsFocusTarget:''};vm.createContext(ctx);vm.runInContext(src,ctx);ctx.bddsCombineTables();
assert.deepEqual(el.bddsCombinedHead.rows[0].cells.map(c=>c.textContent),['Показатель','2026-10','2026-11']);
assert.deepEqual(el.bddsCombinedBody.rows[0].cells.map(c=>c.textContent),['Продажи','100','200']);
assert.deepEqual(el.bddsCombinedBody.rows[1].cells.map(c=>c.textContent),['Поступления','—','250']);
assert.equal(el.bddsCombinedBody.rows.length,6);assert.equal(el.bddsCombinedBody.rows[3].cells[1].textContent,'80');assert.equal(el.bddsCombinedBody.rows[4].cells[1].textContent,'20');assert.equal(el.bddsCombinedBody.rows[5].cells[1].textContent,'30');
console.log('History precedes current month; uniquely matched rows combine; ambiguous and history-only rows remain separate');
el.bddsBody.children=[];el.bddsHistoryBody.children=[];
el.bddsBody.appendChild(row('ОПЕРАЦИОННАЯ ДЕЯТЕЛЬНОСТЬ','—'));el.bddsBody.appendChild(row('Поступления','200'));el.bddsBody.appendChild(row('ФИНАНСОВАЯ ДЕЯТЕЛЬНОСТЬ','—'));el.bddsBody.appendChild(row('Поступления','400'));el.bddsBody.appendChild(row('Списания','500'));
el.bddsHistoryBody.appendChild(row('Сальдо по операционной деятельности (ОД)','-10'));el.bddsHistoryBody.appendChild(row('Поступления','100'));el.bddsHistoryBody.appendChild(row('Сальдо по финансовой деятельности (ФД)','-20'));el.bddsHistoryBody.appendChild(row('Поступления','300'));el.bddsHistoryBody.appendChild(row('Выплаты','350'));
ctx.bddsCombineTables();assert.deepEqual(el.bddsCombinedBody.rows[1].cells.map(c=>c.textContent),['Поступления','100','200']);assert.deepEqual(el.bddsCombinedBody.rows[3].cells.map(c=>c.textContent),['Поступления','300','400']);assert.deepEqual(el.bddsCombinedBody.rows[4].cells.map(c=>c.textContent),['Списания','350','500']);
console.log('Activity distinguishes identical subtotal labels');
