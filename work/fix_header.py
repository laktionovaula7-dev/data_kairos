from pathlib import Path
p=Path('work/turnover-monitor-preview.html'); s=p.read_text(encoding='utf-8')
s=s.replace('Срок годности уже истёк','Срок истекает в ближайшие 90 дней',1)
s=s.replace('Срок истекает за 90 дней','Срок истекает в ближайшие 30 дней',1)
# replace the two summary values in order
old='349 партий · 51 831 шт.'
pos=s.find(old)
if pos<0: raise SystemExit('summary value missing')
pos2=s.find(old,pos+len(old))
s=s[:pos2]+'129 партий · 13 690 шт.'+s[pos2+len(old):]
s=s.replace('<option value="expired">Срок уже истёк · 582 партии</option>','<option value="expired" disabled>Срок уже истёк · 0 партий</option>')
s=s.replace('Все SKU из отчёта · 148','Все SKU из отчёта оборачиваемости · 148')
s=s.replace('<option value="risk">Сначала самые срочные</option><option value="stock">Сначала большой остаток</option><option value="name">По названию</option>','<option value="risk">По общему риску</option><option value="expiry">Сначала ближайший срок</option><option value="turnover">Сначала большая оборачиваемость</option><option value="stock">Сначала большой остаток</option><option value="name">По названию</option>')
# ensure runtime sort block supports new choices
oldsort="if(sort.value==='name')a.sort((x,y)=>x.name.localeCompare(y.name,'ru'));else if(sort.value==='stock')a.sort((x,y)=>(y.quantity??y.excessQty??0)-(x.quantity??x.excessQty??0));else a.sort((x,y)=>{const rank=z=>z.kind==='expired'?0:z.kind==='expiring'?1:z.kind==='turnover'&&z.days>180?2:3;return rank(x)-rank(y)||(x.daysLeft??99999)-(y.daysLeft??99999)||(y.excessCost??0)-(x.excessCost??0)});"
newsort="if(sort.value==='name')a.sort((x,y)=>x.name.localeCompare(y.name,'ru'));else if(sort.value==='stock')a.sort((x,y)=>(y.quantity??y.excessQty??0)-(x.quantity??x.excessQty??0));else if(sort.value==='expiry')a.sort((x,y)=>(x.daysLeft??99999)-(y.daysLeft??99999));else if(sort.value==='turnover')a.sort((x,y)=>(y.days??-1)-(x.days??-1));else a.sort((x,y)=>{const rank=z=>z.kind==='expiring'?0:z.kind==='turnover'&&z.days>180?1:2;return rank(x)-rank(y)||(x.daysLeft??99999)-(y.daysLeft??99999)||(y.excessCost??0)-(x.excessCost??0)});"
if oldsort in s: s=s.replace(oldsort,newsort)
p.write_text(s,encoding='utf-8')
print('updated header and signal labels')
