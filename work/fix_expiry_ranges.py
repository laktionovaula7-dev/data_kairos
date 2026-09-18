from pathlib import Path
p=Path('work/turnover-monitor-preview.html'); s=p.read_text(encoding='utf-8')
s=s.replace('<option value="expiring30">Срок до 30 дней · 129 партий</option><option value="expiring">Срок до 90 дней · 349 партий</option>', '<option value="expiring30">Срок 0–30 дней · 129 партий</option><option value="expiring31_90">Срок 31–90 дней · 220 партий</option><option value="expiring">Срок 0–90 дней · 349 партий</option>')
s=s.replace("t==='expiring30'?x.kind==='expiring'&&x.daysLeft>=0&&x.daysLeft<=30:t==='expiring'?x.kind==='expiring':", "t==='expiring30'?x.kind==='expiring'&&x.daysLeft>=0&&x.daysLeft<=30:t==='expiring31_90'?x.kind==='expiring'&&x.daysLeft>=31&&x.daysLeft<=90:t==='expiring'?x.kind==='expiring':")
s=s.replace('<option value="expiry">Сначала ближайший срок</option><option value="turnover">', '<option value="expiry">Срок: сначала ближайший</option><option value="expiryDesc">Срок: сначала дальний</option><option value="turnover">')
s=s.replace("else if(sort.value==='turnover')a.sort((x,y)=>(y.days??-1)-(x.days??-1));", "else if(sort.value==='expiryDesc')a.sort((x,y)=>(y.daysLeft??-1)-(x.daysLeft??-1));else if(sort.value==='turnover')a.sort((x,y)=>(y.days??-1)-(x.days??-1));")
p.write_text(s,encoding='utf-8')
