from pathlib import Path
import shutil

root=Path(__file__).resolve().parents[1]
paths=[root/'Kairos_proto_template.html',root/'Kairos_dashboard_proto.html']
def replace_once(s,a,b):
    if s.count(a)!=1: raise ValueError(f'Expected one match, got {s.count(a)}: {a[:90]}')
    return s.replace(a,b,1)

css='''<style>
.sales-workspace .sf-filter-panel{background:#fff;border:1px solid var(--line);border-radius:14px;padding:18px 20px;margin-bottom:12px}.sales-workspace .sf-bar{border:0;padding:0;margin:0;display:grid;grid-template-columns:minmax(0,2fr) minmax(0,1fr);gap:24px}.sf-filter-group{min-width:0}.sf-group-label{font-size:12px;font-weight:700;color:#555;margin-bottom:10px}.sf-fields{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}.sf-filter-group:last-child .sf-fields{grid-template-columns:repeat(2,minmax(0,1fr))}.sales-workspace .sf-f{min-width:0;color:#666;text-transform:none;letter-spacing:0;font-size:12px}.sales-workspace .msel-sum{min-width:0;width:100%;font-family:inherit;font-size:12px;min-height:42px;overflow-wrap:anywhere}.sales-workspace .msel-pop{min-width:240px;max-width:min(380px,80vw)}.sales-workspace .sf-filter-group:last-child .sf-f:last-child .msel-pop{left:auto;right:0}.sales-workspace .sf-chips{gap:6px}.sales-workspace .sf-chip{background:#eeece7;color:#333;font-weight:500;font-size:11px;padding:5px 10px}.sales-workspace .sf-chip b{color:#555}.sf-metrics{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));padding:18px 0 24px;margin-bottom:4px}.sf-stat{padding:0 16px;border-right:1px solid var(--line)}.sf-stat:first-child{padding-left:0}.sf-stat:last-child{border:0}.sf-stat span{display:block;font-size:11px;color:#666;margin-bottom:7px}.sf-stat strong{font-size:19px;font-weight:700;font-variant-numeric:tabular-nums}.sf-stat small{display:block;color:#777;font-size:10px;margin-top:4px}.sales-workspace .rz-note{color:#686868}.sf-breakdown-head{display:flex;justify-content:space-between;gap:15px;align-items:center;flex-wrap:wrap;margin-bottom:16px}.sf-breakdown-head .card-title{margin:0}.sf-table-actions{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-top:14px;font-size:12px;color:#666}.sf-row-button{background:none;border:0;color:inherit;text-align:left;font:inherit;padding:5px 0;cursor:pointer;text-decoration:underline;text-decoration-color:#ccc;text-underline-offset:4px}.sf-row-selected{background:#fff7d6}.sf-row-hint{display:block;color:#777;font-size:10px;margin-top:3px}.sf-context-heading{font-size:12px;color:#666;margin:8px 0 14px}.sf-focus{border:1px solid var(--line);border-radius:14px;padding:22px;background:#fff;margin:0 0 16px;scroll-margin-top:80px}.sf-focus-head{display:flex;align-items:start;justify-content:space-between;gap:14px;margin-bottom:16px}.sf-focus-head h2{font-size:18px;line-height:1.5;margin:0;overflow-wrap:anywhere}.sf-focus-actions{display:flex;gap:8px;flex-wrap:wrap}.sf-focus-metrics{display:flex;gap:30px;flex-wrap:wrap;margin-bottom:18px}.sf-focus-metrics span{font-size:12px;color:#666}.sf-focus-metrics b{color:#2a2a2a;margin-left:8px}.sf-growth{margin:26px 0 12px;font-size:18px;font-weight:700}.sf-disclosure{background:white;border:1px solid var(--line);border-radius:14px;margin-bottom:14px}.sf-disclosure>summary{padding:20px;cursor:pointer;font-size:15px;font-weight:700}.sf-disclosure>summary span{font-weight:400;color:#666;font-size:12px;margin-left:14px}.sf-disclosure-body{padding:0 20px 20px}.sf-disclosure-body>section.card{border:0;box-shadow:none;padding:0;margin:0}.sales-workspace .scope-lbl{display:none}.sales-workspace .sf-summary-note{font-size:12px;color:#666;margin:0 0 16px}.sales-workspace button:focus-visible,.sales-workspace summary:focus-visible{outline:2px solid #7a6415;outline-offset:3px}
@media(max-width:1100px){.sales-workspace .sf-bar{grid-template-columns:1fr}.sf-fields{grid-template-columns:repeat(4,minmax(0,1fr))}.sf-metrics{grid-template-columns:repeat(3,minmax(0,1fr));gap:20px}.sf-stat:nth-child(3){border:0}.sf-stat:nth-child(4){padding-left:0}}@media(max-width:650px){.sf-fields{grid-template-columns:repeat(2,minmax(0,1fr))}.sf-metrics{grid-template-columns:repeat(2,minmax(0,1fr))}.sf-stat{padding:0 10px!important;border:0}.sf-stat strong{font-size:17px}.sf-focus-head{flex-direction:column}.sf-disclosure>summary span{display:block;margin:6px 0 0}.sf-disclosure-body{padding:0 12px 16px}.sales-workspace .msel-pop{min-width:190px}.sales-workspace .sf-f:nth-child(even) .msel-pop{left:auto;right:0}}
</style>'''

for path in paths:
    s=path.read_text(encoding='utf-8')
    backup=root/'output'/(path.stem+'-before-sales-layout.html')
    if not backup.exists(): shutil.copy2(path,backup)
    s=replace_once(s,"var dims=DIMS.filter(function(d){return !_hide[d[0]];});","var dims=DIMS; // Stable grouping controls, independent of filters.")
    old="var bar='<div class=\"sf-bar\">'+msel('Филиал','filial',FILIALS,false)+msel('Регион','okrug',okrugs,regionDisabled)+msel('Город','city',cities,false)+msel('Менеджер','manager',mgrs,false)+msel('Бренд','brand',['HeadRock','KRONbuild','ENKI'],false)+cliField()+'</div>';"
    new="""var bar='<div class="sf-filter-panel"><div class="sf-bar"><div class="sf-filter-group"><div class="sf-group-label">Территория и ответственность</div><div class="sf-fields">'+msel('Филиал','filial',FILIALS,false)+msel('Регион','okrug',okrugs,regionDisabled)+msel('Город','city',cities,false)+msel('Менеджер','manager',mgrs,false)+'</div></div><div class="sf-filter-group"><div class="sf-group-label">Объект анализа</div><div class="sf-fields">'+msel('Бренд','brand',['HeadRock','KRONbuild','ENKI'],false)+cliField()+'</div></div></div></div>';"""
    s=replace_once(s,old,new)
    start=s.index("    var kpis='<div class=\"kpi-grid\">'",s.index('function salesFilterPage'))
    end=s.index("    var MN=",start)
    s=s[:start]+'''    function compactStat(label,value,note){return '<div class="sf-stat"><span>'+label+'</span><strong>'+value+'</strong>'+(note?'<small>'+note+'</small>':'')+'</div>';}
    var kpis='<div class="sf-metrics">'+compactStat('Продажи среза',fmtMoney(sales),'за выбранный период')+compactStat('Δ к прошлому году',growth!=null?delta(+growth.toFixed(1)):'—','по текущей методике АППГ')+compactStat('Клиентов',String(nCl),'с операциями за период')+compactStat('Заказы¹',String(nOrd),'оценка: клиент × день')+compactStat('Средний заказ¹',fmtMoney(nOrd?Math.round(sales/nOrd):0),'по оценке числа заказов')+compactStat('Количество, шт.',qty.toLocaleString('ru-RU'),'в выбранном срезе')+'</div>';
''' +s[end:]
    start=s.index('    var arrR=full.slice(0,30)',s.index('function salesFilterPage'))
    end=s.index('    // ---------- 2)',start)
    s=s[:start]+'''    var limit=state.sfLimit||30,arrR=full.slice(0,limit);
    var byLabel={clients:'Клиент',cities:'Город',managers:'Менеджер',okrugs:'Регион',brands:'Бренд',products:'Товар'};
    var focus=state.sfFocus; if(focus&&(focus.by!==by||!full.some(function(x){return x.name===focus.name;})))focus=null;
    window.__sfRowNames=arrR.map(function(x){return x.name;});
    var rowsR=arrR.map(function(x,idx){var ac=_cls[x.name]||'C';return '<tr'+(focus&&focus.name===x.name?' class="sf-row-selected"':'')+'><td><button class="sf-row-button" onclick="selectSFRow('+idx+')" aria-expanded="'+!!(focus&&focus.name===x.name)+'">'+esc(x.name)+'</button></td><td>'+abcBadge(ac)+'</td><td class="num">'+fmtMoney(x.v)+'</td><td class="num">'+pct(+(x.v/(sales||1)*100).toFixed(1))+'</td></tr>';}).join('');
    var razbivka='<section class="card pad razrez"><div class="sf-breakdown-head"><div class="card-title">Разбивка продаж</div><span class="rz-note">Нажмите название строки — детали откроются ниже</span></div>'
      +'<div class="rz-seg-l">Группировать по</div>'+segR
      +'<div class="table-wrap" style="max-height:520px;margin-top:12px"><table class="data-table"><thead><tr><th>'+(byLabel[by]||'—')+'</th><th>ABC</th><th class="num">Продажи</th><th class="num">Доля</th></tr></thead><tbody>'+(rowsR||'<tr><td colspan=4 class="empty">Нет данных. Проверьте сочетание фильтров.</td></tr>')+'</tbody></table></div>'
      +'<div class="sf-table-actions"><span>Показано '+arrR.length+' из '+full.length+' · позиции с положительными продажами</span>'+(full.length>limit?'<button class="ghost" onclick="showMoreSF()">Показать ещё 30</button>':'')+'</div>'
      +'<div class="rz-note">ABC рассчитан внутри выбранного среза. Доля — от его продаж нетто; возвраты могут влиять на сумму долей.</div></section>';
    var focusHtml='';
    if(focus){
      var fm={},fp={},ft=0,fc={},fb={};
      for(var fi=0;fi<D.tx.length;fi++){var tx=D.tx[fi],cn=D.txClients[tx[1]],sk=tx[2];if(!match(cn,sk))continue;
        var dk=by==='clients'?cn:by==='cities'?(CT[cn]||'— город не указан'):by==='managers'?(CM[cn]||'— не назначен'):by==='okrugs'?(CR[cn]||'—'):by==='brands'?brOf(sk):(D.txSkus[sk]||['','',0])[1];if(dk!==focus.name)continue;
        var yy=Math.floor(tx[0]/10000),mn=Math.floor(tx[0]/100)%100;
        if(yy===yr)fm[mn]=(fm[mn]||0)+tx[4];if(yy===yr-1)fp[mn]=(fp[mn]||0)+tx[4];
        if(tx[0]>=lo&&tx[0]<=hi){ft+=tx[4];fc[cn]=1;var bk=by==='clients'?(D.txSkus[sk]||['',''])[1]:cn;fb[bk]=(fb[bk]||0)+tx[4];}}
      var fd=MN.map(function(m,i){return {m:m,i:i+1,fact:fm[i+1]||0,prev:fp[i+1]||0};});
      window.__sfFocusClient=by==='clients'?focus.name:null;
      var frows=Object.keys(fb).sort(function(a,b){return fb[b]-fb[a];}).slice(0,8).map(function(n){return '<tr><td>'+esc(n)+'</td><td class="num">'+fmtMoney(fb[n])+'</td></tr>';}).join('');
      focusHtml='<section class="sf-focus" id="sfFocus"><div class="sf-focus-head"><div><div class="sf-context-heading">Детализация строки · общий отбор сохранён</div><h2>'+esc(focus.name)+'</h2></div><div class="sf-focus-actions">'+(by==='clients'?'<button class="ghost" onclick="openSFClient()">Открыть карточку клиента</button>':'')+'<button class="ghost" onclick="closeSFFocus()">Закрыть детали</button></div></div><div class="sf-focus-metrics"><span>Продажи за период <b>'+fmtMoney(ft)+'</b></span><span>Клиентов <b>'+Object.keys(fc).length+'</b></span></div><div class="card-title">Динамика выбранной строки · '+yr+'</div>'+sfChart(fd,selM)+'<div class="card-title" style="margin-top:22px">'+(by==='clients'?'Товары клиента':'Клиенты выбранной строки')+' · топ-8</div><div class="table-wrap"><table class="data-table"><thead><tr><th>'+(by==='clients'?'Товар':'Клиент')+'</th><th class="num">Продажи за период</th></tr></thead><tbody>'+frows+'</tbody></table></div></section>';
    }

''' +s[end:]
    old="    return st+pageHead('Продажи','Фильтры сверху задают срез. Все блоки ниже — строго в этом срезе.')+bar+chips+kpis+chart+razbivka+assortCard+coverageCard;"
    new="""    var opportunity='<div class="sf-growth">Возможности роста</div><p class="sf-summary-note">Блоки ниже относятся ко всему выбранному срезу, а не к открытой строке.</p>'
      +'<details class="sf-disclosure" '+(state.sfAssortOpen?'open':'')+' ontoggle="state.sfAssortOpen=this.open"><summary>Ассортимент<span>'+tot.b+' из '+tot.t+' позиций покупают в срезе · '+tot.na+' непокупаемых позиций в наличии</span></summary><div class="sf-disclosure-body"><p class="sf-summary-note">Для группы клиентов «берёт» означает покупку хотя бы одним клиентом группы.</p>'+assortCard+'</div></details>'
      +'<details class="sf-disclosure" '+(state.sfCoverageOpen?'open':'')+' ontoggle="state.sfCoverageOpen=this.open"><summary>Покрытие территории<span>'+covC+' из '+totC+' крупных городов с покупками</span></summary><div class="sf-disclosure-body">'+coverageCard+'</div></details>';
    return st+SF_LAYOUT_STYLE+'<div class="sales-workspace">'+pageHead('Продажи','Детализация по территории, ответственным, брендам и клиентам.')+bar+chips+'<div class="sf-context-heading">Выбранный срез: '+esc(scopeText())+'</div>'+kpis+razbivka+focusHtml+(focus?'':chart)+opportunity+'<p class="sf-summary-note">¹ Текущая оценка заказов объединяет покупки одного клиента за день. Методика расчётов сохранена.</p></div>';"""
    # Inline handlers cannot access the closure's state; use exported functions.
    new=new.replace('ontoggle="state.sfAssortOpen=this.open"','ontoggle="setSFDisclosure(\'assort\',this.open)"').replace('ontoggle="state.sfCoverageOpen=this.open"','ontoggle="setSFDisclosure(\'coverage\',this.open)"')
    # Escape quotes for surrounding JS string literals.
    new=new.replace("setSFDisclosure('assort',this.open)","setSFDisclosure(\\'assort\\',this.open)").replace("setSFDisclosure('coverage',this.open)","setSFDisclosure(\\'coverage\\',this.open)")
    s=replace_once(s,old,new)
    import json
    helpers='''  window.selectSFRow=function(index){var name=(window.__sfRowNames||[])[index];if(name==null)return;var by=state.sfBy||'clients';state.sfFocus=state.sfFocus&&state.sfFocus.by===by&&state.sfFocus.name===name?null:{by:by,name:name};render();var el=document.getElementById('sfFocus');if(el)el.scrollIntoView({block:'start',behavior:'auto'});};
  window.closeSFFocus=function(){state.sfFocus=null;render();};
  window.openSFClient=function(){if(window.__sfFocusClient)openClientByName(window.__sfFocusClient);};
  window.showMoreSF=function(){state.sfLimit=(state.sfLimit||30)+30;render();};
  window.setSFDisclosure=function(which,open){state[which==='assort'?'sfAssortOpen':'sfCoverageOpen']=open;};
'''
    s=replace_once(s,'  function salesFilterPage(){','  var SF_LAYOUT_STYLE='+json.dumps(css,ensure_ascii=False)+';\n'+helpers+'  function salesFilterPage(){')
    s=replace_once(s,"window.setSFBy=function(v){state.sfBy=v;if(window.render)render();};","window.setSFBy=function(v){state.sfBy=v;state.sfFocus=null;state.sfLimit=30;if(window.render)render();};")
    path.write_text(s,encoding='utf-8')
    print('Updated',path.name)
