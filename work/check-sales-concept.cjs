const fs=require('fs'),vm=require('vm'),assert=require('assert');
const html=fs.readFileSync('output/kairos-sales-concept.html','utf8');
const nodes=new Map();
for(const m of html.matchAll(/id="([^"]+)"/g)){assert(!nodes.has(m[1]),'duplicate id');nodes.set(m[1],{value:'',innerHTML:'',textContent:'',hidden:false,dataset:{},add(){},focus(){},setAttribute(){},querySelectorAll(){return []}})}
const document={getElementById:id=>nodes.get(id)||null,addEventListener(){}};
const code=html.match(/<script>([\s\S]*?)<\/script>/)[1];
new vm.Script(code).runInNewContext({document,Option:function(){},Intl,console});
const n=id=>nodes.get(id),change=(id,value)=>{n(id).value=value;n(id).onchange({target:{value}})};
const initial=n('metrics').innerHTML;assert(initial.includes('млн ₽'));assert.equal((n('months').innerHTML.match(/data-month=/g)||[]).length,12);
change('period','7');assert.notEqual(n('metrics').innerHTML,initial);assert(n('metrics').innerHTML.includes('Август'));assert.equal((n('months').innerHTML.match(/data-month=/g)||[]).length,12);
change('branch','Москва');const moscow=n('metrics').innerHTML;change('brand','HeadRock');assert.notEqual(n('metrics').innerHTML,moscow);
change('group','brand');assert.equal(n('dimension').textContent,'Бренд');assert(n('rows').innerHTML.includes('HeadRock'));
change('manager','Иван Петров');assert(n('rows').innerHTML.includes('Нет данных'));assert(n('metrics').innerHTML.includes('из 0'));
change('manager','all');change('brand','all');change('branch','all');change('group','client');assert(n('rowCount').textContent.includes('Показано 8'));n('more').onclick();assert(n('rowCount').textContent.includes('Показано 16'));
n('search').oninput({target:{value:'несуществующее название'}});assert(n('rows').innerHTML.includes('Нет данных'));
n('extra').hidden=true;n('filterButton').onclick();assert.equal(n('extra').hidden,false);n('closeFilters').onclick();assert.equal(n('extra').hidden,true);
console.log('PASS: JavaScript syntax and execution, month selection, full-year chart, combined filters, grouping, empty state, pagination, search and filter disclosure.');
