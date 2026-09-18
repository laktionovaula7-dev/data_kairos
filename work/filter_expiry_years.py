import json,re
from pathlib import Path
p=Path('work/turnover-monitor-preview.html'); s=p.read_text(encoding='utf-8')
m=re.search(r'const data=(\[.*?\]);',s); data=json.loads(m.group(1))
old=len(data)
data=[x for x in data if x.get('kind') not in ('expired','expiring') or str(x.get('expires',''))[-4:] in ('2026','2027')]
s=s[:m.start(1)]+json.dumps(data,ensure_ascii=False,separators=(',',':'))+s[m.end(1):]
s=s.replace('582 партии · 41 713 шт.','297 партий · 24 210 шт.')
s=s.replace('В 722 партиях нет даты, а в 347 партиях с читаемой датой срок годности уже истёк; они не включены в карточки срока годности.','В карточки срока годности включены только партии с датой окончания в 2026 или 2027 году; более ранние даты исключены.')
p.write_text(s,encoding='utf-8')
print('filtered',old-len(data),'old expiry records; left',len(data))
