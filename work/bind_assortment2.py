from pathlib import Path
p=Path('Kairos_dashboard_final.html'); s=p.read_text(encoding='utf-8')
old='function clientAssortment2(c){return `<div class="embedded-preview"><iframe title="Ассортимент 2" src="work/assortment-review.html" style="display:block;width:100%;height:1450px;border:0;border-radius:12px;background:#fff"></iframe></div>`;}'
if old not in s: raise SystemExit('assortment2 iframe function not found')
s=s.replace(old, 'function clientAssortment2(c){return clientAssortment(c);}', 1)
p.write_text(s,encoding='utf-8')
print('assortment_2 now uses selected client data')
