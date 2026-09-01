#!/usr/bin/env python3
"""
Монітор закупівель ТМО у Prozorro.

Тримає список закупівель установи, щогодини перевіряє зміни статусів,
пише data/procurement.json для додатку і шле сповіщення в Telegram.

Запуск: python3 prozorro_monitor.py
"""

import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

EDRPOU = "44496574"
CPV_PREFIX = ("336", "3314")

API = "https://public-api.prozorro.gov.ua/api/2.5"
PORTAL = "https://prozorro.gov.ua"

STATE_PATH = "prozorro_state.json"
OUT_PATH = "data/procurement.json"

# Скільки тримати закупівлю в data/procurement.json.
# Незавершені лишаються завжди; решта живе рівно стільки, скільки має
# практичний сенс: підписаний договір ще їде на склад, зірваний торг ще
# треба перезапустити. Далі закупівля забувається.
RETENTION_MONTHS = 12          # глибина сканування пошуку
KEEP_ACTIVE_DAYS  = 60         # незавершені: далі це вже покинуті торги
KEEP_SIGNED_DAYS  = 30         # підписані: товар доїжджає на склад і тема закрита
KEEP_PROBLEM_DAYS = 30         # зірвані й скасовані: місяць, далі забуваємо
# Версія логіки розбору. Зміна цього числа знецінює збережені знімки:
# статуси перечитуються заново, а масові "зміни" не йдуть у Telegram.
PARSER_VERSION = 20
UA = "likion-procurement-monitor/1.0 (+https://github.com/oskhmil/medsklad)"

TG_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TG_CHAT = os.environ.get("TELEGRAM_CHAT_ID", "")

ST_SIGNED = "Підписано"
ST_WINNER = "Переможець"
ST_PROGRESS = "В процесі"
ST_FAILED = "Не відбулось"
ST_CANCELLED = "Скасовано"

PROBLEM = (ST_FAILED, ST_CANCELLED)


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def get_json(url, tries=3):
    for attempt in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
            with urllib.request.urlopen(req, timeout=45) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 429:
                time.sleep(5 * (attempt + 1))
                continue
            if attempt == tries - 1:
                raise
            time.sleep(2)
        except Exception:
            if attempt == tries - 1:
                raise
            time.sleep(2)
    return None


# ── пошук закупівель установи ───────────────────────────────────────────────
# ЦБД Prozorro не має пошуку, тому первинний перелік беремо з пошукового
# сервісу порталу. Він не задокументований і приймає POST (GET дає 405),
# тому пробуємо кілька форматів тіла і друкуємо відповідь сервера.

SEARCH_URL = PORTAL + "/api/search/tenders"

# Знайдено дослідним шляхом: поле замовника — buyer (масив),
# код ДК — у повному форматі 33600000-6. Стеля видачі пошуку — 10000.
BUYER_FIELD = "buyer"
CPV_ROOT = "33600000-6"
# Розділ 336 (фармацевтична продукція) + група 3314 (медичні матеріали).
# Пошук не розкриває групу за кореневим кодом, тому перелічуємо підкоди явно.
# Неіснуючі коди відсіються самі — скрипт опитує кожен і лишає результативні.
CPV_CHILDREN = [
    # 336 — фармацевтична продукція
    "33600000-6", "33610000-9", "33620000-2", "33630000-5", "33640000-8",
    "33650000-1", "33660000-4", "33670000-7", "33680000-0", "33690000-3",
    # 3314 — медичні матеріали
    "33140000-3", "33141000-0", "33141100-1", "33141110-4", "33141120-7",
    "33141200-2", "33141300-3", "33141310-6", "33141320-9", "33141400-4",
    "33141500-5", "33141600-6", "33141610-9", "33141620-2", "33141640-8",
    "33141700-7", "33141800-8", "33141900-9", "33142000-7", "33143000-4",
    "33144000-1", "33145000-8", "33146000-5", "33147000-2", "33148000-9",
    "33149000-6",
]

# як може називатись поле замовника — перевіряємо перебором
ENTITY_FIELDS = [
    "edrpou", "procuringEntity", "procuringEntityEdrpou", "procuringEntityId",
    "customer", "customers", "buyer", "entity", "identifier", "organization",
]

IGNORED_TOTAL = 9000  # якщо результатів стільки — фільтр не спрацював


def post_json(url, body, timeout=45):
    """Повертає (статус, розібраний_json_або_None, сирий_текст)."""
    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST", headers={
        "User-Agent": UA,
        "Accept": "application/json",
        "Content-Type": "application/json",
    })
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", "replace")
            try:
                return resp.status, json.loads(raw), raw
            except json.JSONDecodeError:
                return resp.status, None, raw
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            return e.code, json.loads(raw), raw
        except json.JSONDecodeError:
            return e.code, None, raw


def extract_rows(payload):
    if not isinstance(payload, dict):
        return None
    for key in ("data", "items", "results", "tenders", "hits"):
        v = payload.get(key)
        if isinstance(v, list):
            return v
        if isinstance(v, dict):
            for k2 in ("items", "data", "hits"):
                if isinstance(v.get(k2), list):
                    return v[k2]
    return None


def pick(rec, *keys):
    for k in keys:
        v = rec.get(k)
        if v:
            return v
    return None


def total_count(payload):
    if not isinstance(payload, dict):
        return None
    for k in ("total", "count", "totalCount", "found"):
        v = payload.get(k)
        if isinstance(v, int):
            return v
    return None


def record_date(rec):
    d = pick(rec, "dateModified", "date", "datePublished", "dateCreated")
    if d:
        return d
    tp = rec.get("tenderPeriod")
    if isinstance(tp, dict):
        return tp.get("startDate") or tp.get("endDate") or ""
    return ""


def msg_of(payload, raw):
    """Читабельне пояснення помилки замість \\u-екранування."""
    if isinstance(payload, dict):
        m = payload.get("message")
        if m:
            errs = payload.get("errors")
            if isinstance(errs, dict):
                return f"{m} | {json.dumps(errs, ensure_ascii=False)[:160]}"
            return str(m)[:200]
    return (raw or "")[:160]


def probe(body):
    """Повертає (total, перший_запис, пояснення)."""
    try:
        status, payload, raw = post_json(SEARCH_URL, body)
    except Exception as ex:
        return None, None, f"мережа: {ex}"
    rows = extract_rows(payload)
    if status != 200 or rows is None:
        return None, None, f"HTTP {status}: {msg_of(payload, raw)}"
    return total_count(payload), (rows[0] if rows else None), ""


_cpv_used = None
_search_meta = {}
_stats = {"short_codes": [], "expected": 0, "seen": 0}


def build_body(page, cpv=None):
    return {BUYER_FIELD: [EDRPOU], "cpv": cpv or _cpv_used or [CPV_ROOT], "page": page}


def discover_tender_ids():
    global _cpv_used
    found = {}
    cutoff = (datetime.now(timezone.utc) - timedelta(days=30 * RETENTION_MONTHS)).isoformat()
    log(f"  межа за датою: {cutoff[:10]}")

    base_total, _, err = probe({BUYER_FIELD: [EDRPOU], "page": 1})
    log(f"  тільки buyer: {base_total if not err else err}")

    # Опитуємо кожен код окремо: неіснуючі відсіються, а заразом видно,
    # яка категорія скільки дає. Далі шукаємо одним запитом по робочих кодах.
    good = []
    for code in CPV_CHILDREN:
        total, _, err = probe({BUYER_FIELD: [EDRPOU], "cpv": [code], "page": 1})
        if err:
            log(f"  ДК {code}: {err[:70]}")
            continue
        if total:
            good.append((code, total))
        time.sleep(0.15)

    if good:
        log("  робочі коди ДК:")
        for code, n in sorted(good, key=lambda x: -x[1]):
            log(f"    {code}: {n}")
        log(f"  разом кодів: {len(good)} з {len(CPV_CHILDREN)}")

    if not good:
        log("  ЖОДЕН код ДК не дав результату — далі не йдемо")
        return found

    # Кожен код опитуємо окремо і зливаємо результати. Складений запит
    # з десятка кодів портал обрізає: 18 кодів дали менше записів, ніж один.
    _cpv_used = [c for c, _ in good]
    cutoff_hit = 0

    for code, expected in sorted(good, key=lambda x: -x[1]):
        group = "Матеріали" if code.startswith("3314") else "Ліки"
        page = 1
        got = 0
        retries = 0
        while page <= 60:
            try:
                status, payload, raw = post_json(
                    SEARCH_URL, {BUYER_FIELD: [EDRPOU], "cpv": [code], "page": page})
            except Exception as ex:
                log(f"    {code} стор.{page}: {ex}")
                time.sleep(5)
                retries += 1
                if retries > 3:
                    break
                continue
            rows = extract_rows(payload) or []
            if not rows:
                # Порожня сторінка при недобраних записах — це майже завжди
                # придушення запитів порталом, а не кінець даних. Чекаємо і пробуємо ще.
                if got < expected and retries < 4:
                    retries += 1
                    time.sleep(4 * retries)
                    continue
                break
            retries = 0
            for r in rows:
                d = record_date(r)
                got += 1
                if d and d < cutoff:
                    cutoff_hit += 1
                    continue
                tid = pick(r, "id", "tender_id", "_id") or pick(r, "tenderID")
                if not tid:
                    continue
                if tid not in found:
                    found[tid] = d
                    _search_meta[tid] = {
                        "title": r.get("title"),
                        "status": r.get("status"),
                        "date": d,
                        "value": r.get("value"),
                        "group": group,
                        "cpv": code,
                    }
                elif group == "Ліки" and _search_meta.get(tid, {}).get("group") == "Матеріали":
                    _search_meta[tid]["group"] = "Ліки"
            page += 1
            time.sleep(0.4)
        _stats["expected"] += expected
        _stats["seen"] += got
        if got < expected:
            _stats["short_codes"].append(f"{code} ({got}/{expected})")
        flag = "" if got >= expected else "  ← НЕДОБІР"
        log(f"    {code} ({group}): переглянуто {got}, очікувалось {expected}{flag}")
        time.sleep(1.0)

    groups = {}
    for m in _search_meta.values():
        g = m.get("group") or "?"
        groups[g] = groups.get(g, 0) + 1
    log(f"  старших за межу відкинуто: {cutoff_hit}")
    log(f"  у межах {RETENTION_MONTHS} міс.: {len(found)} закупівель · " +
        ", ".join(f"{k} {v}" for k, v in sorted(groups.items())))
    return found


def seed_from_file():
    """Запасний варіант: список id вручну в prozorro_seed.txt, по одному в рядок."""
    if not os.path.exists("prozorro_seed.txt"):
        return []
    out = []
    with open("prozorro_seed.txt", encoding="utf-8") as f:
        for line in f:
            s = line.strip()
            if s and not s.startswith("#"):
                out.append(s)
    return out


# ── розбір закупівлі ────────────────────────────────────────────────────────

def lot_status(tender, lot_id):
    """Статус лота (або всієї закупівлі, якщо лотів немає).

    Портал віддає договори без посилання на нагороду (awardID відсутній),
    тому зв'язок відновлюємо за структурою: коли лот один або лотів немає,
    будь-який чинний договір стосується цієї нагороди.
    """
    lots = {l["id"]: l for l in (tender.get("lots") or []) if l.get("id")}
    lot = lots.get(lot_id) if lot_id else None
    # Лот з /lots може не збігтися з переліком у /details — тоді статус
    # рахуємо на рівні закупівлі, а не вигадуємо зрив.
    if lot_id and lot is None and len(lots) == 1:
        lot = next(iter(lots.values()))

    if lot:
        if lot.get("status") == "unsuccessful":
            return ST_FAILED, None, None
        if lot.get("status") == "cancelled":
            return ST_CANCELLED, None, None

    all_awards = tender.get("awards") or []
    if lot_id and lots:
        awards = [a for a in all_awards if a.get("lotID") == lot_id]
        # Позиції відкритих торгів приходять з окремого ендпоінта і несуть
        # номер лота, а нагороди в /details його не мають. Тоді зіставляти
        # нема за чим — беремо всі нагороди закупівлі.
        if not awards and not any(a.get("lotID") for a in all_awards):
            awards = all_awards
    else:
        awards = all_awards

    active = [a for a in awards if a.get("status") == "active"]

    if active:
        a = active[-1]
        sup = (a.get("suppliers") or [{}])[0]
        supplier = sup.get("name")
        amount = (a.get("value") or {}).get("amount")

        contracts = [c for c in (tender.get("contracts") or []) if c.get("status") == "active"]

        linked = [c for c in contracts if c.get("awardID") == a.get("id")]
        if linked:
            return ST_SIGNED, supplier, (linked[0].get("value") or {}).get("amount", amount)

        if lot_id:
            tag = str(lot_id)[:8]
            byid = [c for c in contracts if tag in str(c.get("contractID", ""))]
            if byid:
                return ST_SIGNED, supplier, (byid[0].get("value") or {}).get("amount", amount)

        active_awards = [x for x in all_awards if x.get("status") == "active"]
        if contracts and (not lots or len(contracts) >= len(active_awards)):
            return ST_SIGNED, supplier, (contracts[0].get("value") or {}).get("amount", amount)

        return ST_WINNER, supplier, amount

    if awards and all(a.get("status") == "unsuccessful" for a in awards):
        return ST_FAILED, None, None

    tstatus = tender.get("status", "")
    if tstatus == "unsuccessful":
        return ST_FAILED, None, None
    if tstatus == "cancelled":
        return ST_CANCELLED, None, None
    if tstatus == "complete" and not active:
        return ST_FAILED, None, None
    return ST_PROGRESS, None, None


def matches_cpv(tender):
    codes = [((tender.get("classification") or {}).get("id") or "")]
    for it in tender.get("items", []):
        codes.append(((it.get("classification") or {}).get("id") or ""))
    return any(c.startswith(CPV_PREFIX) for c in codes)


def entity_edrpou(t):
    """ЄДРПОУ замовника: у портальній відповіді він може бути в buyers[]."""
    pe = t.get("procuringEntity") or {}
    ident = pe.get("identifier") or {}
    if ident.get("id"):
        return str(ident["id"])
    for b in (t.get("buyers") or []):
        bi = (b or {}).get("identifier") or {}
        if bi.get("id"):
            return str(bi["id"])
    return None


def tender_items(t):
    """Позиції закупівлі. У частини процедур портал не кладе їх на верхній
    рівень, а ховає всередину лотів — тоді збираємо звідти, підставляючи
    relatedLot, щоб статус рахувався по своєму лоту."""
    items = list(t.get("items") or [])
    if items:
        return items
    for lot in (t.get("lots") or []):
        for it in (lot.get("items") or []):
            it = dict(it)
            it.setdefault("relatedLot", lot.get("id"))
            items.append(it)
    return items


def item_codes(t):
    out = []
    for it in tender_items(t):
        cls = (it.get("classification") or {}).get("id") or ""
        if cls:
            out.append(cls)
    return out


_drop_log = []


def parse_tender(t, meta=None):
    """Плоский опис закупівлі: позиції зі статусами лотів."""
    edr = entity_edrpou(t)
    codes = item_codes(t)
    root = ((t.get("classification") or {}).get("id") or "")

    # пошук уже відфільтрував за buyer, тому розбіжність лише логуємо
    if edr and edr != EDRPOU:
        _drop_log.append((t.get("tenderID"), f"чужий ЄДРПОУ {edr}"))
        return None

    items = []
    for it in tender_items(t):
        cls = (it.get("classification") or {}).get("id") or ""
        if cls and not cls.startswith(CPV_PREFIX):
            continue
        lot_id = it.get("relatedLot")
        st, supplier, amount = lot_status(t, lot_id)
        unit = (it.get("unit") or {}).get("name") or ""
        items.append({
            "name": (it.get("description") or "").strip(),
            "qty": it.get("quantity"),
            "unit": unit,
            "status": st,
            "supplier": supplier,
            "amount": amount,
            "lot": lot_id,
        })

    if not items and t.get("lots"):
        full_lots = fetch_lots(t.get("tenderID"))
        if full_lots:
            items = items_from_lots(t, full_lots)

    if not items:
        # У відкритих торгах портал не віддає номенклатуру — ні на верхньому
        # рівні, ні в лотах. Але лот має назву, суму й статус, а для великих
        # закупівель саме лот і є одиницею, за якою стежать. Тому будуємо
        # позиції з лотів: краще показати лот, ніж загубити закупівлю.
        for lot in (t.get("lots") or []):
            st, supplier, amount = lot_status(t, lot.get("id"))
            items.append({
                "name": (lot.get("title") or t.get("title") or "Лот").strip(),
                "qty": None,
                "unit": "",
                "status": st,
                "supplier": supplier,
                "amount": (lot.get("value") or {}).get("amount"),
                "lot": lot.get("id"),
                "fromLot": True,
            })

    if not items:
        _drop_log.append((t.get("tenderID"),
                          f"немає ні позицій, ні лотів · тип={t.get('procurementMethodType')}"
                          f" · статус={t.get('status')}"))
        return None

    counts = {"signed": 0, "progress": 0, "failed": 0}
    for i in items:
        if i["status"] == ST_SIGNED:
            counts["signed"] += 1
        elif i["status"] in PROBLEM:
            counts["failed"] += 1
        else:
            counts["progress"] += 1

    contract_end = None
    for c in t.get("contracts", []):
        p = (c.get("period") or {}).get("endDate")
        if p and (contract_end is None or p > contract_end):
            contract_end = p

    # Сума належить закупівлі, а не позиції: нагорода видається на лот цілком.
    # Тому в позиції лишаємо суму тільки тоді, коли позиція одна.
    active_contracts = [c for c in (t.get("contracts") or []) if c.get("status") == "active"]
    active_awards = [a for a in (t.get("awards") or []) if a.get("status") == "active"]
    src_vals = active_contracts or active_awards
    total = None
    if src_vals:
        vals = [(v.get("value") or {}).get("amount") for v in src_vals]
        vals = [v for v in vals if isinstance(v, (int, float))]
        if vals:
            total = round(sum(vals), 2)
    if total is None:
        total = (t.get("value") or {}).get("amount")

    if len(items) > 1:
        for i in items:
            i["amount"] = None

    return {
        "id": t.get("tenderID") or t.get("id"),
        "tenderID": t.get("tenderID"),
        "title": ((t.get("title") or (meta.get("title") if meta else "")
                   or (items[0]["name"] if items else "")) or "").strip(),
        "date": t.get("date") or (meta.get("date") if meta else None),
        "dateModified": t.get("dateModified") or (meta.get("date") if meta else None),
        "method": t.get("procurementMethodType"),
        "status": t.get("status"),
        "contractEnd": contract_end,
        "amount": total,
        "group": (meta or {}).get("group") or "Ліки",
        "cpv": (meta or {}).get("cpv"),
        "url": f"{PORTAL}/tender/{t.get('tenderID')}",
        "counts": counts,
        "items": items,
    }


# ── читання деталей ─────────────────────────────────────────────────────────
# Знайдено дослідним шляхом: закупівля читається за публічним номером
# через портальний ендпоінт /api/tenders/{tenderID}/details.
LOTS_URL = PORTAL + "/api/tenders/{id}/lots"

DETAIL_URLS = [
    PORTAL + "/api/tenders/{id}/details",
]
_detail_url = None


def fetch_tender(tid):
    global _detail_url
    urls = [_detail_url] if _detail_url else DETAIL_URLS
    for u in urls:
        try:
            data = get_json(u.format(id=tid), tries=3)
        except Exception:
            continue
        if data:
            if _detail_url is None:
                _detail_url = u
                log(f"  ЕНДПОІНТ ДЕТАЛЕЙ: {u}")
            return data
    return None


def diagnose_detail(tid):
    """Проба на випадок, якщо ендпоінт колись зміниться."""
    log("")
    log("  === ПРОБА ЕНДПОІНТІВ ДЕТАЛЕЙ ===")
    log(f"  номер: {tid}")
    for u in (PORTAL + "/api/tenders/{id}/details",
              PORTAL + "/api/tenders/{id}",
              API + "/tenders/{id}"):
        url = u.format(id=tid)
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
            with urllib.request.urlopen(req, timeout=25) as r:
                raw = r.read().decode("utf-8", "replace")
                d = json.loads(raw)
                keys = sorted(d.keys())[:10] if isinstance(d, dict) else type(d).__name__
                log(f"  GET {u.split('.ua')[-1]}: 200 · ключі {keys}")
        except urllib.error.HTTPError as e:
            log(f"  GET {u.split('.ua')[-1]}: {e.code} · {e.read().decode('utf-8','replace')[:100]}")
        except Exception as ex:
            log(f"  GET {u.split('.ua')[-1]}: {ex}")
    log("  === КІНЕЦЬ ПРОБИ ===")
    log("")


_items_probe_done = False


def probe_items_endpoint(tid):
    """Разова розвідка будови /lots: чи є там справжня номенклатура."""
    global _items_probe_done
    if _items_probe_done:
        return
    _items_probe_done = True
    try:
        data = get_json(LOTS_URL.format(id=tid), tries=1)
    except Exception as ex:
        log(f"  проба /lots: {ex}")
        return
    lots = (data or {}).get("lots") or []
    log(f"  -- будова /lots для {tid}: лотів {len(lots)} --")
    for lot in lots[:2]:
        log(f"    ключі лота: {sorted(lot.keys())[:16]}")
        for key in ("items", "lotItems", "nomenclature", "products"):
            v = lot.get(key)
            if isinstance(v, list) and v:
                log(f"    {key}: {len(v)} шт · ключі {sorted(v[0].keys())[:12]}")
                log(f"      приклад: {json.dumps(v[0], ensure_ascii=False)[:220]}")
    log("  -- кінець проби --")


def fetch_lots(tid):
    """Повні лоти з окремого ендпоінта: там є і номенклатура, і власні
    нагороди та договори лота. У /details цього немає для відкритих торгів."""
    try:
        data = get_json(LOTS_URL.format(id=tid), tries=2)
    except Exception:
        return []
    return (data or {}).get("lots") or []


def items_from_lots(tender, lots):
    """Позиції з лотів зі статусом, порахованим по самому лоту.

    Нагороди й договори лежать усередині лота, тому будуємо для кожного
    окремий контекст — інакше завершена закупівля виглядає як суцільний зрив.
    """
    out = []
    for lot in lots:
        ctx = {
            "status": tender.get("status"),
            "lots": [{"id": lot.get("id"), "status": lot.get("status")}],
            "awards": lot.get("awards") or [],
            "contracts": lot.get("contracts") or [],
        }
        st, supplier, amount = lot_status(ctx, lot.get("id"))
        lot_items = [i for i in (lot.get("items") or []) if isinstance(i, dict)]
        if not lot_items:
            out.append({
                "name": (lot.get("title") or lot.get("description") or "Лот").strip(),
                "qty": None, "unit": "", "status": st, "supplier": supplier,
                "amount": (lot.get("value") or {}).get("amount"),
                "lot": lot.get("id"), "fromLot": True,
            })
            continue
        for it in lot_items:
            cls = (it.get("classification") or {}).get("id") or ""
            if cls and not cls.startswith(CPV_PREFIX):
                continue
            out.append({
                "name": (it.get("description") or it.get("title") or "").strip(),
                "qty": it.get("quantity"),
                "unit": (it.get("unit") or {}).get("name") or "",
                "status": st, "supplier": supplier, "amount": None,
                "lot": lot.get("id"),
            })
    return out


# ── позначка "перезакуплено" ────────────────────────────────────────────────
# Prozorro не звʼязує повторні торги з початковими. Тому зірвану позицію
# зіставляємо за назвою з тими, що згодом успішно закупили: перше значуще
# слово це фактично МНН, і ще одне спільне слово відсікає однофамільців
# ("натрію хлорид" проти "натрію гідрокарбонат"). Закупівлю не ховаємо —
# лише позначаємо, щоб хибне зіставлення було видно, а не мовчки шкодило.

_STOP_TOKENS = {"для", "по", "та", "і", "й", "з", "у", "в", "на", "мл", "мг", "г",
                "шт", "штука", "флакон", "розчин", "таблетки", "порошок",
                "інєкцій", "ін'єкцій", "інфузій", "приготування", "концентрат"}


def item_tokens(name):
    s = str(name or "").lower().replace("'", "").replace("\u2019", "")
    s = "".join(ch if ch.isalnum() else " " for ch in s)
    return [w for w in s.split()
            if len(w) > 2 and w not in _STOP_TOKENS and not w.isdigit()][:8]


def same_item(a, b):
    if not a or not b or a[0] != b[0]:
        return False
    sa, sb = set(a), set(b)
    return len(sa & sb) >= 2 or (len(sa) == 1 and len(sb) == 1)


def announced(t):
    """Дата оголошення з номера: UA-2026-08-11-005208-a → 2026-08-11.

    Поле dateModified для цього не годиться — воно оновлюється щоразу, коли
    монітор перечитує закупівлю, і всі дати збігаються до дня прогону.
    """
    m = re.match(r"^UA-(\d{4}-\d{2}-\d{2})-", str(t.get("tenderID") or ""))
    if m:
        return m.group(1)
    return (t.get("date") or t.get("dateModified") or "")[:10]


def mark_resolved(parsed):
    """Позначає зірвані позиції, які згодом переоголосили.

    Розрізняємо два ступені: торги вже відбулися (є переможець або договір)
    і торги лише тривають. Перше знімає проблему, друге лише пояснює, що
    процес пішов, — тому проблема лишається відкритою.
    """
    done, pending = [], []
    for t in parsed:
        d = announced(t)
        for i in t.get("items", []):
            st = i.get("status")
            tok = item_tokens(i.get("name"))
            if st in (ST_SIGNED, ST_WINNER):
                done.append((tok, d, t.get("tenderID")))
            elif st == ST_PROGRESS:
                pending.append((tok, d, t.get("tenderID")))

    def best(tok, d, pool):
        hit = None
        for st, sd, stid in pool:
            if sd > d and same_item(tok, st):
                if hit is None or sd < hit[0]:
                    hit = (sd, stid)
        return hit

    marked = 0
    for t in parsed:
        d = announced(t)
        resolved_items = 0
        failed_items = 0
        last = None
        for i in t.get("items", []):
            if i.get("status") not in PROBLEM:
                continue
            failed_items += 1
            tok = item_tokens(i.get("name"))
            hit = best(tok, d, done)
            kind = "done"
            if not hit:
                hit = best(tok, d, pending)
                kind = "pending"
            if hit:
                i["resolved"] = {"d": hit[0], "t": hit[1], "k": kind}
                marked += 1
                if kind == "done":
                    resolved_items += 1
                    if last is None or hit[0] > last[0]:
                        last = hit
        if resolved_items and last:
            t["resolved"] = {"count": resolved_items, "of": failed_items,
                             "d": last[0], "t": last[1]}
    return marked


# ── стан і сповіщення ───────────────────────────────────────────────────────

def load_state():
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {"tenders": {}, "seen": {}}


def save_state(state):
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=1, sort_keys=True)


def item_key(t_id, item):
    return f"{t_id}|{item['lot'] or ''}|{item['name'][:80]}"


def diff_statuses(old, parsed, archived=()):
    """Що змінилось порівняно з попереднім запуском."""
    events = []
    for t in parsed:
        if t["id"] in archived:
            continue
        prev = old.get(t["id"])
        if prev is None:
            events.append(("new", t, None, None))
            continue
        prev_items = {k: v for k, v in prev.get("items", {}).items()}
        for it in t["items"]:
            k = item_key(t["id"], it)
            was = prev_items.get(k)
            if was is None:
                continue
            if was != it["status"]:
                events.append(("change", t, it, was))
    return events


def tg_send(text):
    if not TG_TOKEN or not TG_CHAT:
        log("  Telegram не налаштовано — сповіщення пропущено")
        return
    url = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = urllib.parse.urlencode({
        "chat_id": TG_CHAT,
        "text": text,
        "parse_mode": "HTML",
        "disable_web_page_preview": "true",
    }).encode()
    try:
        req = urllib.request.Request(url, data=payload)
        with urllib.request.urlopen(req, timeout=30) as r:
            r.read()
    except Exception as ex:
        log(f"  Telegram помилка: {ex}")


def esc(s):
    return (str(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def notify(events):
    if not events:
        return
    lines = ["🛒 <b>Закупівлі ТМО</b>"]
    for kind, t, item, was in events[:25]:
        if kind == "new":
            lines.append(f"\n🆕 <b>{esc(t['title'][:90])}</b>")
            lines.append(f"{len(t['items'])} позицій · {esc(t['tenderID'])}")
            lines.append(f"{t['url']}")
        else:
            icon = "⛔" if item["status"] in PROBLEM else ("✅" if item["status"] == ST_SIGNED else "🔄")
            lines.append(f"\n{icon} <b>{esc(item['name'][:80])}</b>")
            lines.append(f"{esc(was)} → <b>{esc(item['status'])}</b>")
            if item.get("supplier"):
                lines.append(f"{esc(item['supplier'])}")
            lines.append(f"{t['url']}")
    if len(events) > 25:
        lines.append(f"\n…і ще {len(events) - 25} змін")
    tg_send("\n".join(lines))


# ── зріз для додатку ────────────────────────────────────────────────────────

def tender_state(t):
    """active — ще в торгах; problem — зірвана; done — все підписано."""
    c = t.get("counts") or {}
    if c.get("progress"):
        return "active"
    if c.get("failed"):
        return "problem"
    return "done"


def in_window(t, now):
    """Що лишається у файлі для додатку."""
    state = tender_state(t)
    dm = t.get("dateModified") or t.get("date") or ""
    if not dm:
        return True                      # без дати не викидаємо, щоб не загубити
    days = {"active": KEEP_ACTIVE_DAYS,
            "problem": KEEP_PROBLEM_DAYS}.get(state, KEEP_SIGNED_DAYS)
    edge = (now - timedelta(days=days)).isoformat()
    return dm >= edge



def main():
    log("=" * 60)
    log(f"Монітор закупівель · ЄДРПОУ {EDRPOU} · ДК {', '.join(c + '*' for c in CPV_PREFIX)}")
    log("=" * 60)

    state = load_state()
    stale_parser = state.get("parser") != PARSER_VERSION
    if stale_parser and state.get("tenders"):
        log(f"  логіка розбору змінилась (v{state.get('parser')} → v{PARSER_VERSION}):"
            " знімки перечитуються, сповіщення цього разу не шлемо")
        state["tenders"] = {}
    known = {k: v for k, v in state.get("seen", {}).items() if str(k).startswith("UA-")}
    dropped = len(state.get("seen", {})) - len(known)
    if dropped:
        log(f"  (зі стану відкинуто {dropped} застарілих ключів)")

    log("\n[1] Пошук закупівель установи")
    discovered = discover_tender_ids()
    for tid, dm in discovered.items():
        if tid not in known:
            known[tid] = ""
    log(f"  всього відомих id: {len(known)}")

    if not known:
        log("\nЖОДНОЇ закупівлі не знайдено. Пошуковий ендпоінт порталу")
        log("не відповів, а збереженого стану немає. Дивись розділ [1] вище.")
        return 1

    log("\n[2] Читання деталей з офіційного API")
    time.sleep(2)
    parsed = []
    fetched = 0
    skipped = 0
    errors = 0
    stubs = 0
    archive = state.get("archive") or {}
    archived_skipped = 0

    for tid in sorted(known, key=lambda k: known.get(k) or '', reverse=True):
        # Закупівля, яку ми вже відпрацювали й забули: якщо пошук показує ту саму
        # дату, читати деталі немає сенсу — вона більше не змінюється.
        arch = archive.get(tid)
        if arch and discovered.get(tid, "")[:10] == arch.get("d", ""):
            archived_skipped += 1
            continue
        newer = discovered.get(tid, "")
        cached = state.get("tenders", {}).get(tid, {}).get("dateModified", "")
        if newer and cached and newer <= cached and tid in state.get("tenders", {}):
            snap = state["tenders"][tid].get("snapshot")
            if snap:
                parsed.append(snap)
                skipped += 1
                continue
        resp = fetch_tender(tid)
        if resp is None:
            errors += 1
            if fetched == 0 and errors == 10:
                diagnose_detail(tid)
                log("  жодна закупівля не прочиталась — далі не йдемо")
                return 1
            if errors <= 5:
                log(f"  {tid}: не вдалось прочитати")
            elif errors == 6:
                log("  …подальші помилки не друкую")
            continue
        fetched += 1
        if fetched % 25 == 0:
            log(f"  …{fetched} прочитано")
        t = resp.get("data") or resp or {}
        if len(t) <= 2:
            stubs += 1
            _drop_log.append((tid, f"порожня відповідь ({len(t)} полів)"))
            continue
        meta = _search_meta.get(tid)
        if not meta:
            arch = archive.get(tid)
            if arch and arch.get("t"):
                meta = {"title": arch["t"], "date": arch.get("d")}
        if not (t.get("items") or []) and (t.get("lots") or []):
            probe_items_endpoint(tid)
        p = parse_tender(t, meta)
        if p:
            parsed.append(p)
        time.sleep(0.25)

    log(f"  завантажено: {fetched}, з кешу: {skipped}, з архіву пропущено: {archived_skipped},"
        f" помилок: {errors}, порожніх: {stubs}")
    log(f"  підходять під фільтр: {len(parsed)}")
    if _drop_log:
        log(f"  відсіяно: {len(_drop_log)}")
        for tid, why in _drop_log[:12]:
            log(f"    {tid}: {why}")
        if len(_drop_log) > 12:
            log(f"    …і ще {len(_drop_log) - 12}")

    if parsed:
        total_items = sum(len(p["items"]) for p in parsed)
        sig = sum(p["counts"]["signed"] for p in parsed)
        prg = sum(p["counts"]["progress"] for p in parsed)
        fail = sum(p["counts"]["failed"] for p in parsed)
        log(f"  позицій: {total_items} · підписано {sig} · в процесі {prg} · проблемних {fail}")

    log("\n[3] Порівняння зі станом")
    events = diff_statuses(state.get("tenders", {}), parsed, set((state.get("archive") or {}).keys()))
    log(f"  подій: {len(events)}")
    for kind, t, item, was in events[:10]:
        if kind == "new":
            log(f"  NEW  {t['tenderID']} · {t['title'][:60]}")
        else:
            log(f"  CHG  {item['name'][:50]} · {was} → {item['status']}")

    first_run = not state.get("tenders")
    if first_run or stale_parser:
        log("  перший запуск або зміна логіки — сповіщення не шлемо")
    else:
        notify(events)

    log("\n[4] Запис файлів")
    now = datetime.now(timezone.utc)
    cutoff = (now - timedelta(days=30 * RETENTION_MONTHS)).isoformat()
    now_iso = now.isoformat()

    marked = mark_resolved(parsed)
    if marked:
        log(f"  позначено як перезакуплені: {marked} позицій")
    window = [t for t in parsed if in_window(t, now)]
    window.sort(key=lambda t: t.get("dateModified") or "", reverse=True)
    log(f"  у зрізі за {RETENTION_MONTHS} міс.: {len(window)} з {len(parsed)}")

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    # Самоперевірка: монітор може відпрацювати "успішно", але тихо недобрати
    # даних (портал придушує запити, змінює формат). Тому кладемо в файл
    # звірку "знайдено проти очікуваного" — додаток покаже попередження.
    problems = []
    if _stats["short_codes"]:
        problems.append("недобір по кодах: " + ", ".join(_stats["short_codes"][:5]))
    if errors:
        problems.append(f"не прочитано закупівель: {errors}")
    if _drop_log:
        problems.append(f"відсіяно без позицій: {len(_drop_log)}")

    health = {
        "checked": now_iso,
        "expected": _stats["expected"],
        "seen": _stats["seen"],
        "tenders": len(parsed),
        "errors": errors,
        "dropped": len(_drop_log),
        "status": "ok" if not problems else "warn",
        "problems": problems,
    }
    log("\n[5] Самоперевірка")
    log(f"  очікувалось {_stats['expected']}, переглянуто {_stats['seen']}"
        f", розібрано {len(parsed)}, помилок {errors}")
    log(f"  стан: {health['status']}" + ("" if not problems else " · " + "; ".join(problems)))

    out = {
        "generated": now_iso,
        "edrpou": EDRPOU,
        "source": "prozorro.gov.ua",
        "health": health,
        "tenders": window,
    }
    payload = json.dumps(out, ensure_ascii=False, indent=1)
    old_payload = ""
    if os.path.exists(OUT_PATH):
        with open(OUT_PATH, encoding="utf-8") as f:
            old_payload = f.read()

    # Час перевірки змінюється щогодини, тому порівнюємо лише змістовну
    # частину — інакше репозиторій заростав би порожніми комітами.
    def _meat(text):
        try:
            o = json.loads(text)
        except Exception:
            return text
        o.pop("generated", None)
        h = o.get("health") or {}
        h.pop("checked", None)
        return json.dumps(o, ensure_ascii=False, sort_keys=True)

    data_changed = _meat(payload) != _meat(old_payload)

    # Але раз на 12 годин записуємо все одно, щоб у додатку було видно,
    # що монітор живий, навіть коли в закупівлях тиша.
    stale = True
    try:
        prev = json.loads(old_payload)
        prev_at = (prev.get("health") or {}).get("checked") or prev.get("generated")
        if prev_at:
            age = (now - datetime.fromisoformat(prev_at)).total_seconds()
            stale = age > 12 * 3600
    except Exception:
        pass

    changed = data_changed or stale
    if changed and not data_changed:
        log("  дані не змінились — оновлюю лише час перевірки")
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        f.write(payload)
    log(f"  {OUT_PATH}: {len(payload) // 1024} КБ, змінено: {changed}")

    # Знімки тримаємо лише для того, що лишилось у вікні. Решта йде в архів:
    # id + підсумковий статус + дата, кілька десятків байтів на закупівлю.
    # Архів потрібен, щоб забута закупівля не повернулась як "нова".
    in_window_ids = {t["id"] for t in window}
    new_state = {"parser": PARSER_VERSION, "tenders": {}, "seen": {}, "archive": {}}
    for t in parsed:
        if t["id"] in in_window_ids:
            new_state["tenders"][t["id"]] = {
                "dateModified": t.get("dateModified"),
                "items": {item_key(t["id"], i): i["status"] for i in t["items"]},
                "snapshot": t,
            }
        else:
            new_state["archive"][t["id"]] = {
                "s": tender_state(t),
                "d": (t.get("dateModified") or "")[:10],
                "t": (t.get("title") or "")[:80],
            }
        new_state["seen"][t["id"]] = t.get("dateModified", "")

    # архів попередніх прогонів переносимо далі
    for tid, rec in (state.get("archive") or {}).items():
        new_state["archive"].setdefault(tid, rec)
    for tid, dm in known.items():
        if str(tid).startswith("UA-"):
            new_state["seen"].setdefault(tid, dm)
    save_state(new_state)
    log(f"  {STATE_PATH}: {len(new_state['tenders'])} закупівель у стані")

    marker = "changed" if (changed or events) else "nochange"
    gh_out = os.environ.get("GITHUB_OUTPUT")
    if gh_out:
        with open(gh_out, "a", encoding="utf-8") as f:
            f.write(f"result={marker}\n")
            f.write(f"events={len(events)}\n")
    log(f"\nРезультат: {marker}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
