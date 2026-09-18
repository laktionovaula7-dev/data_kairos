from pathlib import Path
p=Path('Kairos_dashboard_final.html'); s=p.read_text(encoding='utf-8')
old="['logistics','Логистика и остатки']"
if old not in s: raise SystemExit('top nav marker missing')
s=s.replace(old, "['logistics','Логистика и остатки'],['signals','Товарные сигналы']", 1)
old="else if(state.page==='logistics')page=logisticsPage();"
if old not in s: raise SystemExit('render marker missing')
s=s.replace(old, old+"else if(state.page==='signals')page=signalsPage();", 1)
marker='  window.render=function render(){'
if marker not in s: raise SystemExit('render function marker missing')
fn="""  function signalsPage(){return `<div class=\"embedded-preview\" style=\"margin:-6px -2px 0\"><iframe title=\"Товарные сигналы\" src=\"work/turnover-monitor-preview-browser-v7.html\" style=\"display:block;width:100%;height:1550px;border:0;border-radius:12px;background:#fff\"></iframe></div>`;}\n  function clientAssortment2(c){return `<div class=\"embedded-preview\"><iframe title=\"Ассортимент 2\" src=\"work/assortment-review.html\" style=\"display:block;width:100%;height:1450px;border:0;border-radius:12px;background:#fff\"></iframe></div>`;}\n\n"""
s=s.replace(marker,fn+marker,1)
old="['assortment','Ассортимент'],['sales','Продажи']"
if old not in s: raise SystemExit('client tabs marker missing')
s=s.replace(old,"['assortment','Ассортимент'],['assortment2','Ассортимент_2'],['sales','Продажи']",1)
old="if(state.clientTab==='assortment')body=clientAssortment(c);"
if old not in s: raise SystemExit('client body marker missing')
s=s.replace(old,old+"if(state.clientTab==='assortment2')body=clientAssortment2(c);",1)
p.write_text(s,encoding='utf-8')
print('integrated signals and assortment_2 tabs')
