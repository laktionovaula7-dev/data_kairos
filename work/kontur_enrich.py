# -*- coding: utf-8 -*-
"""Массовое обогащение Контур.Фокус с КЭШЕМ и ЛИМИТОМ.
Линковка строго по ИНН (ИНН берём из _client_inn.xlsx по точному имени из 1С).

Запуск (безопасно, с ограничением числа НОВЫХ запросов):
    python work/kontur_enrich.py --max 40
    python work/kontur_enrich.py --inns 2312245627,7722770040
Уже закэшированные ИНН повторно НЕ запрашиваются (квота не тратится).
"""
import os, sys, json, time, argparse, pickle, urllib.request, urllib.parse, urllib.error
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from kontur_lib import compact, nrm_name, ANALYTICS

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = "https://focus-api.kontur.ru"
CACHE = os.path.join(ROOT, "work", "_kontur_cache.json")

def get_key():
    kf = os.path.join(ROOT, "work", "kontur_key.txt")
    if os.path.exists(kf):
        k = open(kf, encoding="utf-8").read().strip()
        if k: return k
    return (os.environ.get("KONTUR_API_KEY") or "").strip()

def call(method, inn, key):
    qs = urllib.parse.urlencode({"key": key, "inn": inn})
    req = urllib.request.Request(f"{BASE}/api3/{method}?{qs}", headers={"Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.getcode(), json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        return e.code, {"__error__": e.read().decode("utf-8", "replace")[:300]}
    except Exception as e:
        return None, {"__error__": repr(e)}

def sales_inns():
    """ИНН клиентов с продажами -> через _client_inn.xlsx (по точному имени)."""
    import openpyxl
    wb = openpyxl.load_workbook(os.path.join(ROOT, "work", "_client_inn.xlsx"), read_only=True, data_only=True)
    name2inn = {}
    for r in wb.active.iter_rows(min_row=2, values_only=True):
        nm, reg, inn = (tuple(r) + (None, None, None))[:3]
        if nm and inn:
            name2inn[nrm_name(str(nm))] = str(inn).strip()
    d = pickle.loads(open(os.path.join(ROOT, "work", "_parsed.pkl"), "rb").read())
    net = d["S"]["client_net"]
    out = []
    for cn in sorted(net, key=lambda x: -net[x]):
        if net[cn] <= 0: continue
        inn = name2inn.get(nrm_name(cn))
        if inn and inn not in out:
            out.append(inn)
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--max", type=int, default=10, help="макс. число НОВЫХ ИНН за запуск (защита квоты)")
    ap.add_argument("--inns", default="", help="явный список ИНН через запятую")
    ap.add_argument("--no-analytics", action="store_true", help="только req+egrDetails (экономит квоту)")
    ap.add_argument("--sleep", type=float, default=0.3)
    args = ap.parse_args()

    key = get_key()
    if not key:
        print("НЕТ КЛЮЧА (work/kontur_key.txt)"); sys.exit(2)

    cache = json.load(open(CACHE, encoding="utf-8")) if os.path.exists(CACHE) else {}
    targets = [x.strip() for x in args.inns.split(",") if x.strip()] if args.inns else sales_inns()
    todo = [i for i in targets if i not in cache][: args.max]
    print(f"в кэше: {len(cache)} | кандидатов: {len(targets)} | будет запрошено сейчас: {len(todo)} (лимит --max {args.max})")

    done = 0
    for inn in todo:
        code1, d1 = call("req", inn, key)
        code2, d2 = call("egrDetails", inn, key)
        if code1 != 200:
            print(f"  ! {inn}: req HTTP {code1} {str(d1)[:120]}")
            continue
        req0 = (d1 or [None])[0]
        egr0 = (d2 or [None])[0] if code2 == 200 else None
        an = None
        if not args.no_analytics:
            an = {}
            for m in ANALYTICS:
                cc, dd = call(m, inn, key)
                if cc == 200:
                    an[m] = (dd[0] if isinstance(dd, list) and dd else dd)
        cache[inn] = compact(req0, egr0, an)
        done += 1
        print(f"  + {inn}: {cache[inn].get('name')} | {cache[inn].get('status')}")
        json.dump(cache, open(CACHE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)  # пишем сразу
        time.sleep(args.sleep)
    print(f"готово: добавлено {done}, всего в кэше {len(cache)}")

if __name__ == "__main__":
    main()
