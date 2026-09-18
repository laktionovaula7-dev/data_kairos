import json,re
from pathlib import Path
p=Path('work/turnover-monitor-preview.html'); s=p.read_text(encoding='utf-8')
m=re.search(r'const data=(\[.*?\]);',s); data=json.loads(m.group(1)); data=[x for x in data if x.get('kind')!='expired']
s=s[:m.start(1)]+json.dumps(data,ensure_ascii=False,separators=(',',':'))+s[m.end(1):]
s=re.sub(r'\d+ партий · [\d ]+ шт\.', '349 партий · 51 831 шт.', s, count=1)
s=s.replace("if((x.kind==='expired'||x.kind==='expiring')&&!['2026','2027'].includes(expiryYear))return false;", "if(x.kind==='expired')return false;if(x.kind==='expiring'&&!['2026','2027'].includes(expiryYear))return false;")
s=re.sub(r'<div class="kp-note">.*?</div>', '<div class="kp-note"><strong>Фокус по срокам:</strong> показываются только партии, срок годности которых ещё не истёк и закончится в ближайшие 90 дней от 17.09.2026. Просроченные партии вынесены из этого списка.</div>', s, count=1)
p.write_text(s,encoding='utf-8')
print('upcoming only',sum(1 for x in data if x.get('kind')=='expiring'))
