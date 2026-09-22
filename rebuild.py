# -*- coding: utf-8 -*-
"""
Регенератор дашборда Кайрос.
Читает актуальные Excel из папки, считает всё сам (ABC, оборачиваемость, год-к-году),
вставляет данные в шаблон (Kairos_dashboard_final.html) и пересобирает дашборд.

Запуск:  python rebuild.py
Обновление: положить свежие выгрузки в папку с теми же именами -> запустить скрипт.

Охват: ВСЯ компания, все бренды (HeadRock/KRONbuild/ENKI), все регионы/страны, 2024-2026.
"""
import sys, re, json, zipfile, pickle, time, os
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path
from datetime import date, datetime

try: sys.stdout.reconfigure(encoding="utf-8")
except Exception: pass

ROOT   = Path(__file__).resolve().parent
F_SALES= ROOT/"Продажи по бизнес регионам 24г-26г.xlsx"
F_STOCK= ROOT/"Остатки и доступность по сериям.xlsx"
F_PRICE= ROOT/"Прайс NEW HR+KRON+ENKI МСК-ВЛ (03 Сентября 2026г) (1).xlsx"
CRM_DIR= ROOT/"Регионы"                                # выгрузки CRM по округам (ЮФО.xls, ПФО.xls, …, ЦФО.files)
F_PLAN = ROOT/"План_25_26.xlsx"

def _find_client_inn():
    if CRM_DIR.exists():
        for p in CRM_DIR.glob("*лиент*ИНН*.xlsx"): return p
    return None
TEMPLATE_SRC = ROOT/"Kairos_dashboard_final.html"      # источник вёрстки (из него берём слот данных)
OUT    = ROOT/"Kairos_dashboard_rebuilt.html"          # результат (отдельный файл до проверки)
CACHE  = ROOT/"work"/"_sales_rows.pkl"                 # кэш сырых строк продаж

CUR_YEAR, PREV_YEAR = 2026, 2025
TURN_WINDOW_DAYS = 90          # окно для оборачиваемости (DOS)

# Правило: исключаем ВСЕ компании с нулевыми (и отрицательными) продажами.
# Явный список имён не ведём — служебные/тестовые записи и так без продаж и отсекаются сами.
EXCLUDE_CLIENTS_RAW = []
MN=['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек']

def nrm(s):
    s=str(s or '').lower().strip().replace('ё','е')
    s=re.sub(r'[«»"\'`(),]',' ',s); s=s.replace('х','x'); s=re.sub(r'\s+',' ',s)
    return s.strip()

# ---------------- ПРАЙС: имя/артикул -> бренд, категория ----------------
def load_price():
    import openpyxl
    if not F_PRICE.exists(): print("  [!] нет прайса"); return {},{}
    wb=openpyxl.load_workbook(F_PRICE,data_only=True,read_only=True); rows=[]
    if 'HeadRock' in wb.sheetnames:
        ws=wb['HeadRock']; cat=None
        for r in ws.iter_rows(min_row=9,values_only=True):
            art=r[1] if len(r)>1 else None; name=r[2] if len(r)>2 else None
            if art and (name is None or str(name).strip()==''): cat=str(art).strip(); continue
            if art and name: rows.append((str(art).strip(),str(name).strip(),'HeadRock',cat))
    if 'KRONbuild' in wb.sheetnames:
        ws=wb['KRONbuild']; cat=None
        for r in ws.iter_rows(min_row=10,values_only=True):
            art=r[0] if len(r)>0 else None; name=r[2] if len(r)>2 else None
            if (art is None or str(art).strip()=='') and name and str(name).strip(): cat=str(name).strip(); continue
            if art and name:
                brand='ENKI' if (cat and 'ENKI' in cat.upper()) else 'KRONbuild'
                rows.append((str(art).strip(),str(name).strip(),brand,cat))
    wb.close()
    price={nrm(n):(a,n,b,c) for a,n,b,c in rows}
    print(f"  прайс: {len(rows)} SKU")
    return price

def brand_of(name, price):
    m=price.get(nrm(name))
    if m: return m[2]
    u=name.lower()
    if 'kronbuild' in u or 'кронбилд' in u: return 'KRONbuild'
    if 'enki' in u or 'энки' in u: return 'ENKI'
    if 'headrock' in u or 'хедрок' in u: return 'HeadRock'
    return 'HeadRock'   # инструмент по умолчанию — HeadRock

def cat_art(name, price):
    m=price.get(nrm(name)); return (m[3],m[0]) if m else (None,None)

# ---------------- ПРОДАЖИ: быстрый загрузчик (значения read_only + уровни из XML) --------
def load_sales_rows():
    if CACHE.exists() and CACHE.stat().st_mtime>=F_SALES.stat().st_mtime:
        print("  продажи: из кэша"); return pickle.loads(CACHE.read_bytes())
    import openpyxl
    t0=time.time()
    ns='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    levels={}
    with zipfile.ZipFile(F_SALES).open('xl/worksheets/sheet1.xml') as fp:
        for ev,el in ET.iterparse(fp,events=('end',)):
            if el.tag==ns+'row':
                r=int(el.get('r')); ol=el.get('outlineLevel'); levels[r]=int(ol) if ol else 0
                el.clear()
    wb=openpyxl.load_workbook(F_SALES,read_only=True,data_only=True); ws=wb['Лист_1']
    rows=[]; r=0
    for v in ws.iter_rows(min_row=1,values_only=True):
        r+=1
        if r<8: continue
        c1=v[0] if len(v)>0 else None
        q =v[5] if len(v)>5 else None
        gr=v[6] if len(v)>6 else None
        rt=v[7] if len(v)>7 else None
        nt=v[8] if len(v)>8 else None
        if c1 is None and nt is None and q is None: continue
        rows.append([levels.get(r,0),(str(c1).strip() if c1 is not None else None),q,gr,rt,nt])
    wb.close()
    CACHE.parent.mkdir(exist_ok=True)
    CACHE.write_bytes(pickle.dumps(rows))
    print(f"  продажи: {len(rows)} строк за {round(time.time()-t0,1)}с")
    return rows

DOC=('заказ клиента','реализация','корректировка','возврат','поступление','перемещение','списание','оприходование','инвентаризация')
_isdoc=lambda s:any(s.lower().startswith(k) for k in DOC)
_rx=re.compile(r'от (\d{2})\.(\d{2})\.(\d{4})')

def parse_sales(rows, price):
    n=len(rows); stack={}
    company={'gross':0.0,'ret':0.0,'net':0.0,'qty':0}
    compYear=defaultdict(lambda:{'gross':0.0,'ret':0.0,'net':0.0,'qty':0})
    brandYear=defaultdict(lambda:defaultdict(float))     # year -> brand -> net
    byBrand=defaultdict(float)
    regSales=defaultdict(float); regClients=defaultdict(lambda:defaultdict(float))
    clientRegion={}
    client_net=defaultdict(float); client_qty=defaultdict(float)
    client_orders=defaultdict(int); client_last=defaultdict(str); client_orderhist=defaultdict(list)
    client_prod=defaultdict(lambda:defaultdict(lambda:[0.0,0.0]))   # client-> name -> [qty,net]
    comp=defaultdict(lambda:[0.0,0.0])                              # name -> [qty,net] по компании
    yearMonth=defaultdict(float)                                    # 'YYYYMM' -> net (компания)
    clientYear=defaultdict(lambda:defaultdict(float))               # client -> year -> net
    tx=[]; tx_clients=[]; tx_cats=[]; tx_skus=[]; ci_m={}; ct_m={}; sk_m={}
    inwork=[]; maxdate=0; orderStack={}   # level -> (dateInt, iso, num, osum) заказа на текущем пути
    def ci(x):
        if x not in ci_m: ci_m[x]=len(tx_clients); tx_clients.append(x)
        return ci_m[x]
    def kt(c):
        c=c or 'Без категории'
        if c not in ct_m: ct_m[c]=len(tx_cats); tx_cats.append(c)
        return ct_m[c]
    def si(name):
        cat,art=cat_art(name,price); key=art or ('n:'+nrm(name))
        if key not in sk_m: sk_m[key]=len(tx_skus); tx_skus.append([art or '—',name,kt(cat),brand_of(name,price)])
        return sk_m[key]
    curorder=None  # (dateInt, isoDate, num, sumnet)
    for i in range(n):
        lvl,name,q,gr,rt,nt=rows[i]
        if name is None:
            stack[lvl]=None
            for L in [x for x in stack if x>lvl]: del stack[L]
            continue
        stack[lvl]=name
        for L in [x for x in stack if x>lvl]: del stack[L]
        for L in [x for x in orderStack if x>=lvl]: del orderStack[L]   # глубже/на этом уровне — не предки
        if lvl==0: continue
        if lvl==1:
            clientRegion[name]=stack.get(0) or '(без региона)'; continue
        low=name.lower()
        def cur_ord():
            ks=[k for k in orderStack if k<lvl]
            return orderStack[max(ks)] if ks else None
        if low.startswith('заказ клиента'):
            client=stack.get(1); m=_rx.search(name)
            client_orders[client]+=1
            osum=nt or 0
            if m:
                di=int(f"{m.group(3)}{m.group(2)}{m.group(1)}"); iso=f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
                orderStack[lvl]=(di,iso,name.split('от')[0].replace('Заказ клиента','').strip(),osum)
                if di>maxdate: maxdate=di
                if iso>client_last[client]: client_last[client]=iso
                client_orderhist[client].append({'date':f"{m.group(1)}.{m.group(2)}.{m.group(3)}",
                    'num':orderStack[lvl][2],'sum':round(osum),'qty':int(q or 0)})
                # заказы в работе: заказ минус его реализации (только за текущий год)
                if di//10000==CUR_YEAR:
                    shipped=0.0; j=i+1; l=lvl
                    while j<n and rows[j][0]>l and rows[j][1] is not None:
                        if rows[j][1].lower().startswith('реализация'): shipped+=rows[j][5] or 0
                        j+=1
                    ns=osum-shipped
                    if ns>1: inwork.append({'client':client,'order':orderStack[lvl][2],
                        'date':f"{m.group(1)}.{m.group(2)}.{m.group(3)}",'ordered':round(osum),
                        'shipped':round(shipped),'not_shipped':round(ns)})
            continue
        if _isdoc(name): continue
        # товар-лист?
        nxt=rows[i+1][0] if i+1<n else -1
        if nxt>lvl: continue     # есть дети — не лист
        if nt is None: continue
        client=stack.get(1); reg=stack.get(0) or '(без региона)'
        q=q or 0; net=nt
        company['net']+=net; company['gross']+=(gr or 0); company['ret']+=(rt or 0); company['qty']+=q
        b=brand_of(name,price); byBrand[b]+=net
        regSales[reg]+=net
        if client:
            regClients[reg][client]+=net
            client_net[client]+=net; client_qty[client]+=q
            client_prod[client][name][0]+=q; client_prod[client][name][1]+=net
        comp[name][0]+=q; comp[name][1]+=net
        co=cur_ord()
        if co:
            di=co[0]; yr=di//10000; ym=str(di)[:6]
            yearMonth[ym]+=net
            cy=compYear[yr]; cy['net']+=net; cy['gross']+=(gr or 0); cy['ret']+=(rt or 0); cy['qty']+=q
            brandYear[yr][b]+=net
            if client: clientYear[client][yr]+=net
            if net!=0 and client:
                tx.append([di,ci(client),si(name),round(q,1),round(net)])
    inwork.sort(key=lambda x:-x['not_shipped'])
    return dict(company=company,byBrand=dict(byBrand),regSales=dict(regSales),
        regClients={k:dict(v) for k,v in regClients.items()},clientRegion=clientRegion,
        client_net=dict(client_net),client_qty=dict(client_qty),client_orders=dict(client_orders),
        client_last=dict(client_last),client_orderhist=dict(client_orderhist),
        client_prod={k:dict(v) for k,v in client_prod.items()},comp=dict(comp),
        yearMonth=dict(yearMonth),clientYear={k:dict(v) for k,v in clientYear.items()},
        compYear={y:dict(d) for y,d in compYear.items()},brandYear={y:dict(d) for y,d in brandYear.items()},
        tx=tx,tx_clients=tx_clients,tx_cats=tx_cats,tx_skus=tx_skus,inwork=inwork,maxdate=maxdate)

# ---------------- ОСТАТКИ (новый файл по сериям) ----------------
def load_stock():
    import openpyxl
    if not F_STOCK.exists(): print("  [!] нет файла остатков"); return {}
    wb=openpyxl.load_workbook(F_STOCK,read_only=True,data_only=True); ws=wb['Лист_1']
    stock={}; r=0
    for v in ws.iter_rows(min_row=1,values_only=True):
        r+=1
        if r<12: continue
        art=v[0] if len(v)>0 else None; name=v[2] if len(v)>2 else None
        if not art or not name: continue          # только строки уровня «Артикул»
        a=str(art).strip()
        g=lambda idx: (v[idx] if len(v)>idx and isinstance(v[idx],(int,float)) else 0)
        stock[a]={'name':str(name).strip(),'in_stock':g(7),'shipping':g(8),
                  'reserved':g(9),'available':g(10),'incoming':g(11)}
    wb.close()
    print(f"  остатки: {len(stock)} артикулов")
    return stock

# ---------------- CRM (ЦФО) ----------------
def _cnorm(s):
    s=str(s or '').lower().replace('ё','е')
    for ch in '«»"\'`(),.-–—/\\+№': s=s.replace(ch,' ')
    s=s.replace('новая компания',' ')
    LEG={'ооо','оао','зао','пао','ао','ип','тпк','тд','тк','нпо','нпф','гк','ooo'}
    return ' '.join(w for w in s.split() if w and w not in LEG)
def _clean(v):
    v=str(v or '').replace('&quot;','"').replace('&amp;','&').replace('&#160;',' ').replace('&nbsp;',' ')
    v=re.sub(r'\s+',' ',v).strip()
    if re.fullmatch(r'[\d,.]+[eE]\+\d+', v): return ''
    return v
def _crm_files():
    files=[]
    if CRM_DIR.exists():
        files += sorted(CRM_DIR.glob("*.xls"))
        f=CRM_DIR/"ЦФО.files"/"sheet001.htm"
        if f.exists(): files.append(f)
    old=ROOT/"ЦФО.files"/"sheet001.htm"
    if old.exists(): files.append(old)
    return files

def _parse_crm_file(path):
    """Разбор одного CRM-файла ПО НАЗВАНИЯМ колонок (разные округа = разные сдвиги)."""
    t=path.read_bytes().decode('utf-8','replace')
    rows=re.findall(r'<tr[^>]*>(.*?)</tr>',t,re.S|re.I)
    if len(rows)<2: return []
    cell=lambda r:[re.sub(r'<[^>]+>','',c) for c in re.findall(r'<t[hd][^>]*>(.*?)</t[hd]>',r,re.S|re.I)]
    hdr=[_clean(h) for h in cell(rows[0])]
    H=defaultdict(list)
    for i,h in enumerate(hdr): H[h].append(i)
    first=lambda name:(H[name][0] if H.get(name) else None)
    typ=H.get('Тип компании',[]); i_cat=(typ[-1] if typ else None); i_cat0=(typ[0] if typ else None)
    F={'name':first('Название компании'),'id':first('ID'),
       'inn':first('Реквизит (Россия): ИНН'),'kpp':first('Реквизит (Россия): КПП'),
       'ogrn':first('Реквизит (Россия): ОГРН'),
       'nameFull':first('Реквизит (Россия): Полное наименование организации'),
       'nameShort':first('Реквизит (Россия): Сокращенное наименование организации'),
       'director':first('Реквизит (Россия): Ген. директор'),'address':first('Реквизит (Россия): Адрес'),
       'phone1':first('Рабочий телефон'),'phone2':first('Мобильный телефон'),
       'email1':first('Рабочий e-mail'),'email2':first('Реквизит (Россия): E-Mail'),
       'site':first('Корпоративный сайт'),'responsible':first('Ответственный'),
       'employees':first('Кол-во сотрудников'),'contractNo':first('Номер договора'),
       'contractDate':first('Дата договора'),'bank':first('Банковский реквизит (Россия): Наименование банка'),
       'city':first('Город'),'subject':first('Субъект РФ'),
       'fam':first('Реквизит (Россия): Фамилия'),'im':first('Реквизит (Россия): Имя'),
       'ot':first('Реквизит (Россия): Отчество')}
    out=[]
    for r in rows[1:]:
        c=cell(r)
        gv=lambda i:(_clean(c[i]) if (i is not None and i<len(c)) else '')
        name=gv(F['name'])
        if not name: continue
        rec={k:gv(F[k]) for k in F}
        rec['name']=name; rec['category']=gv(i_cat) or gv(i_cat0)
        out.append(rec)
    return out

def load_crm(client_names):
    files=_crm_files()
    if not files: print("  CRM: файлы не найдены (папка Регионы)"); return {}
    recs=[]
    for p in files:
        rr=_parse_crm_file(p); recs+=rr
    # индекс: нормализованное имя -> запись (имя + юр.наименования + ФИО ИП)
    idx={}
    for rec in recs:
        fio=' '.join(x for x in [rec.get('fam',''),rec.get('im',''),rec.get('ot','')] if x).strip()
        for x in [rec.get('name',''),rec.get('nameShort',''),rec.get('nameFull',''),fio]:
            k=_cnorm(x)
            if k and k not in idx: idx[k]=rec
    tokidx=[(set(k.split()),rec) for k,rec in idx.items()]
    cid=lambda n:'c_'+re.sub(r'[^a-zа-я0-9]','',n.lower())[:20]
    out={}; matched=0
    for name in client_names:
        k=_cnorm(name); rec=idx.get(k)
        if not rec:
            ts=set(k.split())
            for kt2,rr in tokidx:
                if ts and (ts<=kt2 or kt2<=ts): rec=rr; break
        if not rec: continue
        matched+=1
        phones=', '.join(x for x in [rec['phone1'],rec['phone2']] if x)
        emails=', '.join(dict.fromkeys(x for x in [rec['email1'],rec['email2']] if x))
        out[cid(name)]={'nameShort':rec['nameShort'] or rec['name'],'nameFull':rec['nameFull'],
            'inn':rec['inn'],'kpp':rec['kpp'],'ogrn':rec['ogrn'],'director':rec['director'],
            'address':rec['address'],'city':rec['city'],'subject':rec['subject'],'phone':phones,'email':emails,
            'site':rec['site'],'responsible':rec['responsible'],'category':rec['category'],
            'employees':rec['employees'],'contractNo':rec['contractNo'],'contractDate':rec['contractDate'],
            'bank':rec['bank'],'crmId':rec['id']}
    print(f"  CRM: файлов {len(files)}, компаний {len(recs)}, сматчено {matched}/{len(client_names)}")
    return out

# ---------------- ПЛАН (25/26, помесячно) ----------------
_PMONTHS=['январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь']
def _plan_block(row, ci):
    """6 значений канала после названия месяца в позиции ci:
       [ВЛ ОПТ, МСК ОПТ, МСК СЕТИ] (KRONbuild+Enki) + [ВЛ, МСК, МСК СЕТИ] (HEADROCK)."""
    c=[x if isinstance(x,(int,float)) else 0 for x in row[ci+1:ci+7]]
    if len(c)<6: return None
    return {'total':round(sum(c)),
            'vl':round(c[0]+c[3]),               # Владивосток = ВЛ ОПТ + ВЛ
            'msk':round(c[1]+c[2]+c[4]+c[5]),     # Москва = МСК ОПТ + МСК СЕТИ + МСК + МСК СЕТИ
            'kronenki':round(c[0]+c[1]+c[2]),     # KRONbuild+ENKI
            'headrock':round(c[3]+c[4]+c[5])}     # HEADROCK
def load_plan():
    import openpyxl
    if not F_PLAN.exists(): print("  [!] нет файла плана"); return {}
    wb=openpyxl.load_workbook(F_PLAN,data_only=True,read_only=True)
    plan={}
    # приоритет — сводный «Лист1» (там оба года и корректная разметка «Филиал»)
    sheets=(['Лист1'] if 'Лист1' in wb.sheetnames else [])+[s for s in wb.sheetnames if s.strip().isdigit()]
    def blank(): return {'months':{},'filials':{'Москва':{},'Владивосток':{}},
                         'brands':{'headrock':{},'kronenki':{}},'annual':0}
    for sn in sheets:
        for v in wb[sn].iter_rows(values_only=True):
            for ci,cell in enumerate(v):
                if isinstance(cell,str) and cell.strip().lower() in _PMONTHS:
                    b=_plan_block(v,ci)
                    if not b: continue
                    # год: из соседнего блока справа тоже месяц? определяем по заголовку года не всегда есть.
                    # На «Лист1» левый блок=2025 (кол.D), правый=2026 (кол.K). На year-листах — имя листа.
                    if sn=='Лист1':
                        yr = 2025 if ci<7 else 2026
                    else:
                        try: yr=int(sn.strip())
                        except ValueError: continue
                    mi=_PMONTHS.index(cell.strip().lower())+1
                    p=plan.setdefault(yr,blank())
                    if mi in p['months']: continue
                    p['months'][mi]=b['total']; p['filials']['Москва'][mi]=b['msk']; p['filials']['Владивосток'][mi]=b['vl']
                    p['brands']['headrock'][mi]=b['headrock']; p['brands']['kronenki'][mi]=b['kronenki']
    for yr,p in plan.items():
        p['annual']=sum(p['months'].values())
        p['filialAnnual']={k:sum(m.values()) for k,m in p['filials'].items()}
        p['brandAnnual']={k:sum(m.values()) for k,m in p['brands'].items()}
    wb.close()
    yrs=", ".join(f"{y}: {plan[y]['annual']:,.0f}" for y in sorted(plan))
    print(f"  план: {yrs}")
    return plan

# ---------------- ABC ----------------
def abc_of(comp):
    items=sorted(((n,d[1]) for n,d in comp.items() if d[1]>0), key=lambda x:-x[1])
    tot=sum(v for _,v in items) or 1; cum=0.0; abc={}
    for n,v in items:
        cum+=v; sh=cum/tot
        abc[n]='A' if sh<=0.8 else('B' if sh<=0.95 else 'C')
    for n in comp:
        if n not in abc: abc[n]='C'
    return abc,tot

# ---------------- оборачиваемость (DOS) ----------------
def turnover(comp_recent_qty, stock, price):
    """DOS = В наличии / (продано за окно / дней окна). comp_recent_qty: name->qty за окно."""
    # свести продажи за окно к артикулу
    qty_by_art=defaultdict(float)
    for name,q in comp_recent_qty.items():
        _,art=cat_art(name,price)
        if art: qty_by_art[art]+=q
    out={}
    for art,s in stock.items():
        inst=s['in_stock'] or 0
        sold=qty_by_art.get(art,0)
        perday=sold/TURN_WINDOW_DAYS
        if inst<=0: continue
        dos=(inst/perday) if perday>0 else None   # None = мёртвый запас
        out[art]={'dos':(round(dos) if dos is not None else None),'in_stock':inst,
                  'sold_window':round(sold),'status':('dead' if dos is None else
                    ('short' if dos<30 else ('slow' if dos>180 else 'ok')))}
    return out

def _cid(n): return 'c_'+re.sub(r'[^a-zа-я0-9]','',str(n).lower())[:20]
def _daysago(iso, ref):
    if not iso: return None
    try: y,m,d=map(int,iso.split('-')); return (ref-date(y,m,d)).days
    except Exception: return None

def build_krs(S, stock, plan, crm, price):
    _exc={_cnorm(x) for x in EXCLUDE_CLIENTS_RAW}
    def excluded(nm): return (_cnorm(nm) in _exc) or (S['client_net'].get(nm,0)<=0)
    n_excluded=sum(1 for nm in S['client_net'] if excluded(nm))
    comp=S['comp']; abc,tot_rev = abc_of(comp)
    cy=S['compYear']; by=S['brandYear']
    CUR=CUR_YEAR if CUR_YEAR in cy else max(cy) if cy else CUR_YEAR
    PREV=CUR-1
    maxd=str(S['maxdate']); refiso=f"{maxd[:4]}-{maxd[4:6]}-{maxd[6:]}" if len(maxd)==8 else f"{CUR}-12-31"
    refdt=date(int(refiso[:4]),int(refiso[5:7]),int(refiso[8:]))
    # ---- каталог + ABC (по продажам) ----
    cat_items=sorted(((n,d) for n,d in comp.items() if d[1]>0), key=lambda x:-x[1][1])
    catalog=[]; name2cat={}
    for n,d in cat_items:
        c,a=cat_art(n,price)
        name2cat[n]=len(catalog)
        catalog.append([a or '—', n, c or 'Без категории', abc.get(n,'C'), round(d[1])])
    A=[n for n in abc if abc[n]=='A']; B=[n for n in abc if abc[n]=='B']; C=[n for n in abc if abc[n]=='C']
    rev_by={'A':sum(comp[n][1] for n in A),'B':sum(comp[n][1] for n in B),'C':sum(comp[n][1] for n in C)}
    # покупатели по SKU (для резерва)
    buyers=defaultdict(int)
    for cl,prods in S['client_prod'].items():
        for n in prods:
            if prods[n][1]!=0: buyers[n]+=1
    nclients=len(S['client_prod'])
    a_sorted=sorted(A,key=lambda n:-comp[n][1])
    a_list=[{'article':(cat_art(n,price)[1] or '—'),'name':n,'rev':round(comp[n][1])} for n in a_sorted[:20]]
    reserve=[{'name':n,'nonbuyers':nclients-buyers[n],'rev':round(comp[n][1])}
             for n in sorted(A,key=lambda n:-((nclients-buyers[n])*comp[n][1]))[:15]]
    cat_total=defaultdict(int)
    for n in comp:
        c,_=cat_art(n,price)
        if c: cat_total[c]+=1
    # ---- оборачиваемость (DOS) из tx за окно ----
    cutoff=int((refdt.toordinal()-TURN_WINDOW_DAYS))
    def d2ord(di):
        s=str(di);
        try: return date(int(s[:4]),int(s[4:6]),int(s[6:])).toordinal()
        except Exception: return 0
    recent_qty_art=defaultdict(float)
    for t in S['tx']:
        if d2ord(t[0])>=cutoff:
            sk=S['tx_skus'][t[2]]; art=sk[0]
            if art and art!='—': recent_qty_art[art]+=t[3]
    turn={}
    for art,s in stock.items():
        inst=s['in_stock'] or 0
        if inst<=0: continue
        perday=recent_qty_art.get(art,0)/TURN_WINDOW_DAYS
        dos=(inst/perday) if perday>0 else None
        turn[art]={'dos':(round(dos) if dos is not None else None),'in_stock':inst,
                   'available':s['available'],'name':s['name'],
                   'status':('dead' if dos is None else('short' if dos<30 else('slow' if dos>180 else 'ok')))}
    # ---- overview (текущий год + план + АППГ) ----
    o_cur=cy.get(CUR,{'net':0,'gross':0,'ret':0,'qty':0}); o_prev=cy.get(PREV,{'net':0})
    plancur=plan.get(CUR,{}); planprev=plan.get(PREV,{})
    ym=S['yearMonth']
    def mfact(y,mi): return round(ym.get(f"{y}{mi:02d}",0)/1e6,3)
    months=[{'m':MN[mi-1],'fact':mfact(CUR,mi),
             'prev':mfact(PREV,mi),
             'plan':round(plancur.get('months',{}).get(mi,0)/1e6,3) if plancur else None} for mi in range(1,13)]
    clients_cur={cl:S['clientYear'].get(cl,{}).get(CUR,0) for cl in S['client_net'] if not excluded(cl)}
    clients_cur={cl:v for cl,v in clients_cur.items() if v}
    orders_cur=0
    for cl,hist in S['client_orderhist'].items():
        orders_cur+=sum(1 for h in hist if h['date'].endswith(str(CUR)))
    grow=lambda a,b: (round((a-b)/b*100,1) if b else None)
    cur_month=int(maxd[4:6]) if len(maxd)==8 else 12
    prev_same=sum(ym.get(f"{PREV}{mi:02d}",0) for mi in range(1,cur_month+1))   # АППГ: те же месяцы прошлого года
    plan_ytd=sum(plancur.get('months',{}).get(mi,0) for mi in range(1,cur_month+1)) if plancur else None
    overview={'sales':round(o_cur['net']),'grossSales':round(o_cur['gross']),'returns':round(o_cur['ret']),
        'units':int(o_cur['qty']),'returnsRate':round(-o_cur['ret']/o_cur['gross']*100,2) if o_cur['gross'] else 0,
        'plan':plancur.get('annual'),'planYTD':round(plan_ytd) if plan_ytd else None,
        'planDone':round(o_cur['net']/plan_ytd*100,1) if plan_ytd else None,
        'growth':grow(o_cur['net'],prev_same),
        'clients':len(clients_cur),'orders':orders_cur,'avgOrder':round(o_cur['net']/orders_cur) if orders_cur else 0,
        'months':months,
        'branches':[{'id':'all','name':'Вся компания','sales':round(o_cur['net']),
            'plan':plancur.get('annual'),'growth':grow(o_cur['net'],o_prev.get('net',0)),
            'headrock':round(by.get(CUR,{}).get('HeadRock',0)),'kron':round(by.get(CUR,{}).get('KRONbuild',0)),
            'enki':round(by.get(CUR,{}).get('ENKI',0))}],
        'topClients':[{'name':n,'sales':round(v),'share':round(v/(o_cur['net'] or 1)*100,1)}
            for n,v in sorted(clients_cur.items(),key=lambda x:-x[1])[:5]],
        'topProducts':[{'sku':(cat_art(n,price)[1] or '—'),'name':n,'abc':abc.get(n,'C'),'sales':round(comp[n][1]),
            'stock':(stock.get(cat_art(n,price)[1] or '',{}) or {}).get('in_stock',0)}
            for n,_ in cat_items[:10]],
        'categories':[{'name':c,'sales':round(sum(comp[n][1] for n in comp if cat_art(n,price)[0]==c)),
            'share':0} for c in list(cat_total)[:8]]}
    # ---- бренды (для salesDetail) ----
    brand_cur=by.get(CUR,{})
    salesDetail={'branches':{'moscow':{'name':'Вся компания','sales':round(o_cur['net']),'clients':len(clients_cur),
            'orders':orders_cur,'avgOrder':overview['avgOrder'],
            'brands':[{'id':'headrock','name':'HeadRock','fact':round(brand_cur.get('HeadRock',0))},
                      {'id':'kron','name':'KRONbuild','fact':round(brand_cur.get('KRONbuild',0))},
                      {'id':'enki','name':'ENKI','fact':round(brand_cur.get('ENKI',0))}]},
            'vlad':{'name':'Владивосток','sales':None,'missing':True}},
        'brands':{b.lower().replace('headrock','headrock').replace('kronbuild','kron'):{'name':b,
            'sales':round(brand_cur.get(b,0))} for b in ['HeadRock','KRONbuild','ENKI']},
        'categories':[{'name':c,'sales':round(sum(comp[n][1] for n in comp if cat_art(n,price)[0]==c)),
            'growth':None,'sku':cat_total[c]} for c in sorted(cat_total,key=lambda c:-cat_total[c])],
        'products':[]}
    salesDetail['brands']={'headrock':{'name':'HeadRock','sales':round(brand_cur.get('HeadRock',0)),'clients':len(clients_cur),'orders':orders_cur,'avgOrder':overview['avgOrder'],'brands':[]},
        'kron':{'name':'KRONbuild','sales':round(brand_cur.get('KRONbuild',0)),'clients':len(clients_cur),'orders':orders_cur,'avgOrder':overview['avgOrder'],'brands':[]},
        'enki':{'name':'ENKI','sales':round(brand_cur.get('ENKI',0)),'clients':len(clients_cur),'orders':orders_cur,'avgOrder':overview['avgOrder'],'brands':[]}}
    # ---- клиенты и карточки ----
    clients_list=[]; clientDetail={}; clientAssort={}; clientRegion=S['clientRegion']
    A_set=set(A)
    for name in sorted(S['client_net'],key=lambda c:-S['clientYear'].get(c,{}).get(CUR,0)):
        if excluded(name): continue
        cid=_cid(name); prods=S['client_prod'].get(name,{})
        bought=set(n for n in prods if prods[n][1]!=0)
        rev_cur=S['clientYear'].get(name,{}).get(CUR,0)
        hist=S['client_orderhist'].get(name,[])
        hist_cur=[h for h in hist if h['date'].endswith(str(CUR))]
        ordn=len(hist_cur); last=S['client_last'].get(name,'')
        qty=round(sum(prods[n][0] for n in prods))
        # категории клиента
        catpen=defaultdict(lambda:[0,0.0])
        for n in bought:
            c,_=cat_art(n,price)
            if c: catpen[c][0]+=1; catpen[c][1]+=prods[n][1]
        cats=[]
        for c,(got,r) in sorted(catpen.items(),key=lambda x:-x[1][1]):
            totc=cat_total.get(c,got) or got; pen=round(got/totc*100) if totc else 0
            st='Хорошо' if pen>=50 else('Среднее' if pen>=25 else('Зона роста' if pen>0 else 'Не представлена'))
            cats.append({'name':c,'total':totc,'buy':got,'share':round(r/(sum(prods[x][1] for x in bought) or 1)*100,1),'status':st})
        cov={k:{'total':len(L),'buy':sum(1 for n in L if n in bought)} for k,L in (('A',A),('B',B),('C',C))}
        gaps=sorted([n for n in A if n not in bought],key=lambda n:-comp[n][1])[:12]
        rec=[{'sku':(cat_art(n,price)[1] or '—'),'name':n,'category':(cat_art(n,price)[0] or '—'),
              'brand':brand_of(n,price),'abc':'A','status':'Не покупает','action':'Добавить','effect':None,
              'orgSales':round(comp[n][1])} for n in gaps]
        days=_daysago(last,refdt); st='Активный' if (days is not None and days<=45) else('Риск' if (days or 999)>90 else 'Снижение')
        if rev_cur<=0: st='Риск'
        reg=clientRegion.get(name,'—')
        clients_list.append({'id':cid,'name':name,'manager':reg,'region':reg,'sales':round(rev_cur),
            'plan':None,'growth':grow(rev_cur,S['clientYear'].get(name,{}).get(PREV,0)),
            'brands':1,'categories':len(cats),'sku':len(bought),'last':last.replace('-','.') if last else '—',
            'potential':'Высокий' if len(gaps)>=100 else('Средний' if len(gaps)>=40 else 'Низкий'),'status':st,
            'avgCheck':round(rev_cur/ordn) if ordn else 0,'freq':ordn,'aGap':cov['A']['total']-cov['A']['buy'],
            'days':days})
        # помесячно клиента (текущий год)
        cmonth=defaultdict(float)
        for t in S['tx']:
            if t[1]<len(S['tx_clients']) and S['tx_clients'][t[1]]==name and str(t[0]).startswith(str(CUR)):
                cmonth[int(str(t[0])[4:6])]+=t[4]
        cmonths=[{'m':MN[mi-1],'cur':round(cmonth.get(mi,0)/1e6,3),'prev':None} for mi in range(1,13)]
        clientDetail[cid]={'name':name,'manager':reg,'region':reg,'city':'—','segment':'—','status':st,
            'sales':round(rev_cur),'units':qty,'orders':ordn,'avgOrder':round(rev_cur/ordn) if ordn else 0,
            'last':last.replace('-','.') if last else '—','growth':grow(rev_cur,S['clientYear'].get(name,{}).get(PREV,0)),
            'discount':None,'brands':[{'name':'HeadRock','sales':round(rev_cur),'share':100}],'months':cmonths,
            'categories':cats,'abc':cov,'recommended':rec,
            'ordersHistory':[{'date':h['date'],'order':h['num'],'brand':'—','sum':h['sum'],'sku':h['qty'],
                'discount':None,'status':'Из отчёта'} for h in hist_cur[-12:]],'buySku':len(bought)}
        clientAssort[cid]={str(name2cat[n]):round(prods[n][1]) for n in bought if n in name2cat}
    # ---- регионы ----
    regionsData={};
    for reg,cls in S['regClients'].items():
        cls={n:v for n,v in cls.items() if not excluded(n)}
        lst=sorted([{'name':n,'sales':round(v)} for n,v in cls.items()],key=lambda x:-x['sales'])
        regionsData[reg]={'sales':round(sum(cls.values())),'clients':len(cls),'list':lst}
    foreign=[r for r in S['regSales'] if r in ('Беларусь','Казахстан','Кыргызстан')]
    regionsFull=[{'name':r,'sales':round(v)} for r,v in sorted(S['regSales'].items(),key=lambda x:-x[1])]
    # ---- остатки/дефицит ----
    stockOut={a:{'name':s['name'],'in_stock':s['in_stock'],'shipping':s['shipping'],
        'reserved_client':s['reserved'],'reserved_sale':0,'available':s['available'],
        'incoming':s['incoming'],'total_available':s['available']} for a,s in stock.items()}
    low=[{'article':a,'name':s['name'],'available':s['available'],'incoming':s['incoming'],'in_stock':s['in_stock']}
         for a,s in stock.items() if s['available']<=0]
    low.sort(key=lambda x:x['available']); low=low[:40]
    # ---- сборка ----
    KRS={
      'meta':{'prototype':True,'note':f'Пересобрано rebuild.py из актуальных Excel. Вся компания, {CUR} (АППГ {PREV}).',
        'sourceCoverage':{'period':f'{CUR} (АППГ {PREV})','scope':'Вся компания / все бренды',
          'actualNetSales':round(o_cur['net']),'actualUnits':int(o_cur['qty']),'actualReturns':round(o_cur['ret']),
          'missing':['маржинальность','себестоимость','скидки']}},
      'periods':[{'id':'ytd','label':f'{CUR} год','available':True}],
      'overview':overview,
      'abcSummary':{'totalRevenue':round(tot_rev),'counts':{'A':len(A),'B':len(B),'C':len(C)},'totalSku':len(catalog)},
      'salesDetail':salesDetail,'regions':{},'managers':{},
      'clients':clients_list,'clientDetail':clientDetail,'defaultClient':clients_list[0]['id'] if clients_list else '',
      'sourceLists':{'groups':regionsFull},
      'stock':stockOut,'ordersInWork':S['inwork'],'lowStock':low,
      'managersFull':{}, 'regionsFull':regionsFull,
      'abcDetail':{'counts':{'A':len(A),'B':len(B),'C':len(C)},'revenue':{k:round(v) for k,v in rev_by.items()},
        'revenue_share':{k:round(v/(tot_rev or 1)*100,1) for k,v in rev_by.items()},'total_rev':round(tot_rev),
        'total_sku':len(catalog),'nclients':nclients,'a_list':a_list,'reserve':reserve},
      'tx':S['tx'],'txClients':S['tx_clients'],'txCats':S['tx_cats'],'txSkus':S['tx_skus'],'catTotal':dict(cat_total),
      'regionsData':regionsData,'clientRegion':clientRegion,'foreignRegions':foreign or ['Беларусь'],
      'refDate':refiso,'dataAsOf':refiso,'dataAsOfHuman':refiso[8:]+'.'+refiso[5:7]+'.'+refiso[:4],
      'catalog':catalog,'clientAssort':clientAssort,'crm':crm,
      'turnover':turn,'plan':{str(y):plan[y] for y in plan},
    }
    print(f"  исключено служебных/нулевых: {n_excluded}")
    return KRS

def inject(KRS):
    src=TEMPLATE_SRC.read_text(encoding='utf-8')
    key='window.KRS_DATA = '
    i=src.find(key)
    if i<0: raise SystemExit('слот window.KRS_DATA не найден в шаблоне')
    j=i+len(key); _,rel=json.JSONDecoder().raw_decode(src[j:]); end=j+rel
    out=src[:j]+json.dumps(KRS,ensure_ascii=False)+src[end:]
    OUT.write_text(out,encoding='utf-8')
    print(f"  [OK] {OUT.name} пересобран ({len(out):,} байт)")

if __name__=='__main__':
    print("Читаю источники…")
    price=load_price()
    rows=load_sales_rows()
    S=parse_sales(rows,price)
    stock=load_stock()
    plan=load_plan()
    crm=load_crm(list(S['client_net'].keys()))
    print(f"  ИТОГО нетто: {S['company']['net']:,.0f} | клиентов: {len(S['client_net'])} | tx: {len(S['tx'])}")
    print(f"  бренды: "+", ".join(f"{k} {v:,.0f}" for k,v in sorted(S['byBrand'].items(),key=lambda x:-x[1])))
    print(f"  регионы: "+", ".join(f"{k} {v:,.0f}" for k,v in sorted(S['regSales'].items(),key=lambda x:-x[1])))
    pickle.dump({'S':S,'stock':stock,'plan':plan,'crm':crm,'price':price}, open(ROOT/"work"/"_parsed.pkl","wb"))
    print("Собираю модель…")
    KRS=build_krs(S,stock,plan,crm,price)
    ov=KRS['overview']
    print(f"  overview {CUR_YEAR}: продажи {ov['sales']:,} ₽ | план YTD {ov['planYTD']:,} | "
          f"выполн {ov['planDone']}% | АППГ {ov['growth']}% (к тем же мес. {CUR_YEAR-1})")
    print(f"  клиентов {len(KRS['clients'])} | каталог {len(KRS['catalog'])} SKU | оборач. {len(KRS['turnover'])} арт.")
    inject(KRS)
