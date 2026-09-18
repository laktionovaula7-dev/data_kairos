from pathlib import Path
p=Path('work/turnover-monitor-preview.html'); s=p.read_text(encoding='utf-8')
s=s.replace('<option value="expiring">Срок до 90 дней · 349 партий</option><option value="expired" disabled>Срок уже истёк · 0 партий</option>', '<option value="expiring30">Срок до 30 дней · 129 партий</option><option value="expiring">Срок до 90 дней · 349 партий</option>')
s=s.replace("const typ=t==='alerts'?x.kind!=='regular':t==='slow'?x.kind==='turnover'&&x.days>180:t==='regular'?x.kind==='turnover':x.kind===t;", "const typ=t==='alerts'?x.kind!=='regular':t==='slow'?x.kind==='turnover'&&x.days>180:t==='expiring30'?x.kind==='expiring'&&x.daysLeft>=0&&x.daysLeft<=30:t==='expiring'?x.kind==='expiring':t==='regular'?x.kind==='turnover':x.kind===t;")
s=s.replace('<option value="risk">Сначала самые срочные</option><option value="stock">Сначала большой остаток</option><option value="name">По названию</option>', '<option value="risk">По общему риску</option><option value="expiry">Сначала ближайший срок</option><option value="turnover">Сначала большая оборачиваемость</option><option value="stock">Сначала большой остаток</option><option value="name">По названию</option>')
# make category options and filter controls unambiguously active
s=s.replace('aria-label="Тип сигнала"', 'aria-label="Фильтр по типу сигнала"')
s=s.replace('aria-label="Категория товара"', 'aria-label="Фильтр по категории"')
s=s.replace('aria-label="Сортировка"', 'aria-label="Сортировка результатов"')
p.write_text(s,encoding='utf-8')
