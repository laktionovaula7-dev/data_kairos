# -*- coding: utf-8 -*-
# Тест Контур.Фокус на ОДНОМ клиенте: req (базовые реквизиты) + egrDetails (расширенные).
# Ключ берётся из work/kontur_key.txt или env KONTUR_API_KEY. В чат ключ не попадает.
import os, sys, json, pickle, urllib.request, urllib.parse
sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = "https://focus-api.kontur.ru"

def get_key():
    kf = os.path.join(ROOT, "work", "kontur_key.txt")
    if os.path.exists(kf):
        k = open(kf, encoding="utf-8").read().strip()
        if k: return k
    return (os.environ.get("KONTUR_API_KEY") or "").strip()

def pick_client(inn_arg=None):
    """Берём одного клиента с ИНН (по умолчанию — с наибольшими продажами)."""
    d = pickle.loads(open(os.path.join(ROOT, "work", "_parsed.pkl"), "rb").read())
    crm = d.get("crm", {})
    if inn_arg:
        for name, r in crm.items():
            if isinstance(r, dict) and str(r.get("inn", "")).strip() == inn_arg:
                return name, r
        return None, {"inn": inn_arg}
    # иначе — первый осмысленный с ИНН
    for name, r in crm.items():
        if isinstance(r, dict) and r.get("inn"):
            return name, r
    return None, None

def call(method, inn, key):
    qs = urllib.parse.urlencode({"key": key, "inn": inn})
    url = f"{BASE}/api3/{method}?{qs}"
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.getcode(), json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "replace")[:500]
        return e.code, {"__error__": body}
    except Exception as e:
        return None, {"__error__": repr(e)}

if __name__ == "__main__":
    key = get_key()
    if not key:
        print("НЕТ КЛЮЧА. Положи ключ в work/kontur_key.txt (одной строкой) или env KONTUR_API_KEY.")
        sys.exit(2)
    inn_arg = sys.argv[1] if len(sys.argv) > 1 else None
    name, rec = pick_client(inn_arg)
    inn = str((rec or {}).get("inn", "")).strip()
    print(f"Тестовый клиент: {name!r} | ИНН {inn}")
    if not inn:
        print("У клиента нет ИНН."); sys.exit(2)
    out = {}
    for method in ("req", "egrDetails"):
        code, data = call(method, inn, key)
        out[method] = {"http": code, "data": data}
        print(f"/api3/{method}: HTTP {code}", "— ОШИБКА" if (isinstance(data, dict) and data.get("__error__")) else "— ок")
    open(os.path.join(ROOT, "work", "_kontur_test_out.json"), "w", encoding="utf-8").write(
        json.dumps(out, ensure_ascii=False, indent=2))

    def g(d, *path, default=None):
        for p in path:
            if isinstance(d, list): d = d[p] if isinstance(p, int) and len(d) > p else default
            elif isinstance(d, dict): d = d.get(p, default)
            else: return default
            if d is None: return default
        return d

    req = g(out, "req", "data", 0) or {}
    egr = g(out, "egrDetails", "data", 0) or {}
    UL, egUL = req.get("UL", {}), egr.get("UL", {})
    print("\n===== СВОДКА (что доступно для карточки) =====")
    print("Наименование :", g(UL, "legalName", "readable"))
    print("ИНН / КПП    :", req.get("inn"), "/", g(UL, "kpp"))
    print("ОГРН         :", req.get("ogrn"), "| рег.дата:", g(UL, "legalName", "date"))
    print("Статус       :", g(req, "status", "statusString") or g(UL, "status", "statusString") or req.get("status"))
    print("ОПФ          :", g(UL, "opf"))
    head = g(UL, "heads", 0) or g(egUL, "heads", 0) or {}
    print("Руководитель :", (head.get("fio") if isinstance(head, dict) else None), "|", (head.get("position") if isinstance(head, dict) else None))
    print("Капитал      :", g(UL, "authorizedCapital", "sum") or g(egUL, "authorizedCapital", "sum"))
    print("Адрес        :", g(UL, "legalAddress", "parsedAddressRF", "oneLineFormatOfAddress"))
    okveds = g(egUL, "okveds") or g(UL, "okveds") or []
    main_okved = next((o for o in okveds if isinstance(o, dict) and o.get("isMain")), (okveds[0] if okveds else None))
    print("ОКВЭД осн.   :", (str(main_okved.get("code")) + " " + str(main_okved.get("text"))) if isinstance(main_okved, dict) else None)
    print("\nВерхние ключи req   :", list(req.keys()))
    print("Верхние ключи UL    :", list(UL.keys()))
    print("Верхние ключи egrUL :", list(egUL.keys()))
    print("\nСырой ответ сохранён: work/_kontur_test_out.json")
