# -*- coding: utf-8 -*-
"""Контур.Фокус: сжатие ответа req+egrDetails в компактный набор для карточки.
Линковка клиент->ИНН строго по ИНН (имя из 1С только как ключ к справочнику _client_inn.xlsx)."""
import re

def _g(d, *path, default=None):
    for p in path:
        if isinstance(d, list):
            d = d[p] if isinstance(p, int) and len(d) > p else default
        elif isinstance(d, dict):
            d = d.get(p, default)
        else:
            return default
        if d is None:
            return default
    return d

ANALYTICS = ["legalAnalytics", "bankruptcyAnalytics", "courtAnalytics",
             "financeAnalytics", "fsspAnalytics", "linkAnalytics", "purchasesAnalytics"]

def _dv(block, field):
    o = (block or {}).get(field)
    return o.get("data") if isinstance(o, dict) else None

def _analytics_compact(an):
    """an = {method: item}, где item содержит <method>Data. Возвращает finance+risk или None."""
    if not an:
        return None, None
    def blk(m):
        it = an.get(m) or {}
        if isinstance(it, list):
            it = it[0] if it else {}
        return (it.get(m + "Data") or {})
    fin_b, legal, bankr = blk("financeAnalytics"), blk("legalAnalytics"), blk("bankruptcyAnalytics")
    court, fssp, link, purch = blk("courtAnalytics"), blk("fsspAnalytics"), blk("linkAnalytics"), blk("purchasesAnalytics")
    finance = None
    if fin_b:
        finance = {
            "year": _dv(fin_b, "lastBuhYear"),
            "revenue": _dv(fin_b, "incomeEnd") or _dv(fin_b, "incomeSum"),
            "profit": _dv(fin_b, "profitEnd"),
            "tax": _dv(fin_b, "taxSum"),
            "balance": _dv(fin_b, "balanceEnd"),
        }
    risk = None
    if legal or bankr or fssp or court or purch or link:
        risk = {
            "unreliable": any([_dv(legal, "unreliableAddress"), _dv(legal, "unreliableManager"),
                               _dv(legal, "unreliableFounder"), _dv(legal, "unreliableGroup")]),
            "inUnreliableRegister": any([_dv(legal, "register3Month"), _dv(legal, "register6Month"),
                                         _dv(legal, "register12Month")]),
            "fnsAccountBlockDate": _dv(legal, "lastBankAccountFnsBlockDate"),
            "additionalCheck": _dv(legal, "additionalCheckNeeded"),
            "bankruptcyInProgress": bool(_dv(bankr, "defendantInProgressCount")) or bool((bankr.get("currentStage") or {}).get("data")),
            "fsspCount": _dv(fssp, "count"),
            "fsspSum": _dv(fssp, "sum"),
            "courtDefendantSum": _dv(court, "defendantAllSum3Year"),
            "courtPetitionerSum": _dv(court, "petitionerAllSum3Year"),
            "inRnp": _dv(purch, "inRnp"),
            "linkAddressToday": _dv(link, "addressTodayCount"),
            "linkManagerCount": _dv(link, "managerFioCount"),
            "linkFounderCount": _dv(link, "founderFioCount"),
        }
    return finance, risk

def compact(req0, egr0, an=None):
    """req0 = элемент /api3/req; egr0 = элемент /api3/egrDetails; an = {method: item} аналитики."""
    req0 = req0 or {}
    egr0 = egr0 or {}
    UL = req0.get("UL", {}) or {}
    egUL = egr0.get("UL", {}) or {}
    head = _g(UL, "heads", 0) or {}
    act = _g(egUL, "activities", "principalActivity") or {}
    founders = []
    for f in (egUL.get("foundersFL") or []):
        founders.append({"fio": f.get("fio"), "share": _g(f, "share", "percentagePlain")})
    return {
        "inn": req0.get("inn"),
        "ogrn": req0.get("ogrn"),
        "kpp": UL.get("kpp"),
        "name": _g(UL, "legalName", "readable") or _g(UL, "legalName", "short"),
        "fullName": _g(UL, "legalName", "full"),
        "opf": UL.get("opf"),
        "status": _g(UL, "status", "statusString"),
        "statusCode": _g(UL, "status", "code"),
        "regDate": UL.get("registrationDate") or _g(UL, "legalName", "date"),
        "head": {"fio": head.get("fio"), "position": head.get("position")} if head else None,
        "address": _g(UL, "legalAddress", "parsedAddressRF", "oneLineFormatOfAddress"),
        "okvedCode": act.get("code"),
        "okvedText": act.get("text"),
        "capital": _g(egUL, "statedCapital", "sum"),
        "founders": founders,
        "phones": _g(req0, "contactPhones", "count", default=0),
        "emails": _g(req0, "contactEmails", "count", default=0),
        "focusHref": req0.get("focusHref"),
        "finance": (lambda fr: fr[0])(_analytics_compact(an)),
        "risk": (lambda fr: fr[1])(_analytics_compact(an)),
    }

def nrm_name(s):
    return re.sub(r"\s+", " ", (s or "").strip().lower().replace("ё", "е"))
