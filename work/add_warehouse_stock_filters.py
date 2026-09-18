import re
from pathlib import Path
p=Path('work/turnover-monitor-preview.html'); s=p.read_text(encoding='utf-8')
pat=r'<select id="kp-category".*?</select><select id="kp-sort".*?</select>'
new='<select id="kp-category" aria-label="Фильтр по категории"><option value="">Все категории</option></select><select id="kp-warehouse" aria-label="Фильтр по складу"><option value="">Все склады</option></select><select id="kp-sort" aria-label="Сортировка результатов"><option value="risk">По общему риску</option><option value="expiry">Срок: сначала ближайший</option><option value="expiryDesc">Срок: сначала дальний</option><option value="stockDesc">Остаток: сначала больше</option><option value="stockAsc">Остаток: сначала меньше</option><option value="turnover">Оборачиваемость: сначала больше</option><option value="name">По названию</option></select>'
s,n=re.subn(pat,new,s,count=1)
if n!=1: raise SystemExit('markup regex not found')
s=s.replace("const search=$('#kp-search'),type=$('#kp-type'),category=$('#kp-category'),sort=$('#kp-sort'),grid", "const search=$('#kp-search'),type=$('#kp-type'),category=$('#kp-category'),warehouse=$('#kp-warehouse'),sort=$('#kp-sort'),grid")
s=s.replace("categories.forEach(x=>{const o=document.createElement('option');o.value=x;o.textContent=x;category.appendChild(o)});", "categories.forEach(x=>{const o=document.createElement('option');o.value=x;o.textContent=x;category.appendChild(o)});const warehouses=[...new Set(data.map(x=>x.warehouse).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'ru'));warehouses.forEach(x=>{const o=document.createElement('option');o.value=x;o.textContent=x;warehouse.appendChild(o)});")
s=s.replace("(!q||text.includes(q))&&(!category.value||x.category===category.value)", "(!q||text.includes(q))&&(!category.value||x.category===category.value)&&(!warehouse.value||x.warehouse===warehouse.value)")
s=s.replace("else if(sort.value==='stock')a.sort((x,y)=>(y.quantity??y.excessQty??0)-(x.quantity??x.excessQty??0));", "else if(sort.value==='stockDesc')a.sort((x,y)=>(y.quantity??y.excessQty??y.endStock??0)-(x.quantity??x.excessQty??x.endStock??0));else if(sort.value==='stockAsc')a.sort((x,y)=>(x.quantity??x.excessQty??x.endStock??0)-(y.quantity??y.excessQty??y.endStock??0));")
s=s.replace('[search,type,category,sort].forEach', '[search,type,category,warehouse,sort].forEach')
p.write_text(s,encoding='utf-8')
