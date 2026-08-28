#!/usr/bin/env python3
"""
Монітор закупівель ТМО у Prozorro.

Тримає список закупівель установи, щогодини перевіряє зміни статусів,
пише data/procurement.json для додатку і шле сповіщення в Telegram.

Запуск: python3 prozorro_monitor.py
"""

import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

EDRPOU = "44496574"
CPV_PREFIX = ("336",)

API = "https://public-api.prozorro.gov.ua/api/2.5"
PORTAL = "https://prozorro.gov.ua"

STATE_PATH = "prozorro_state.json"
OUT_PATH = "data/procurement.json"

RETENTION_MONTHS = 12
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
# підлеглі коди розділу 336 — на випадок, якщо пошук не розкриває групу сам
CPV_CHILDREN = [
    "33600000-6", "33610000-9", "33620000-2", "33630000-5", "33640000-8",
    "33650000-1", "33660000-4", "33670000-7", "33680000-0", "33690000-3",
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


def build_body(page, cpv=None):
    return {BUYER_FIELD: [EDRPOU], "cpv": cpv or _cpv_used or [CPV_ROOT], "page": page}


def discover_tender_ids():
    global _cpv_used
    found = {}
    cutoff = (datetime.now(timezone.utc) - timedelta(days=30 * RETENTION_MONTHS)).isoformat()
    log(f"  межа за датою: {cutoff[:10]}")

    base_total, _, err = probe({BUYER_FIELD: [EDRPOU], "page": 1})
    log(f"  тільки buyer: {base_total if not err else err}")

    # який варіант ДК ловить більше — корінь групи чи перелік підкодів
    variants = [("корінь", [CPV_ROOT]), ("підкоди", CPV_CHILDREN)]
    best = None
    for label, codes in variants:
        total, sample, err = probe({BUYER_FIELD: [EDRPOU], "cpv": codes, "page": 1})
        if err:
            log(f"  ДК {label}: {err}")
            continue
        title = str((sample or {}).get("title", ""))[:50]
        log(f"  ДК {label}: всього {total} · {title}")
        if total and (best is None or total > best[1]):
            best = (codes, total, label)

    if not best:
        log("  ДК-фільтр не працює — беру всі закупівлі замовника")
        _cpv_used = None
        body_fn = lambda p: {BUYER_FIELD: [EDRPOU], "page": p}
    else:
        _cpv_used = best[0]
        log(f"  ОБРАНО ДК: {best[2]} ({best[1]} закупівель)")
        body_fn = lambda p: build_body(p)

    page = 1
    stale = 0
    seen = 0
    ordered = None
    prev_date = None

    while page <= 40:
        try:
            status, payload, raw = post_json(SEARCH_URL, body_fn(page))
        except Exception as ex:
            log(f"  сторінка {page}: {ex}")
            break
        rows = extract_rows(payload) or []
        if not rows:
            break

        if page == 1:
            log(f"  дати першої сторінки: {[record_date(r)[:10] for r in rows[:5]]}")

        fresh = 0
        for r in rows:
            d = record_date(r)
            seen += 1
            if prev_date is not None and d and prev_date:
                if d > prev_date and ordered is not False:
                    ordered = False
            if d:
                prev_date = d
            if d and d < cutoff:
                continue
            fresh += 1
            tid = pick(r, "id", "tender_id", "_id") or pick(r, "tenderID")
            if tid:
                found[tid] = d
                _search_meta[tid] = {
                    "title": r.get("title"),
                    "status": r.get("status"),
                    "date": d,
                    "value": r.get("value"),
                }

        page += 1
        time.sleep(0.2)

    if ordered is False:
        log("  УВАГА: видача не впорядкована за датою, зупинка за межею ненадійна")
    log(f"  переглянуто {seen} записів на {page - 1} сторінках")
    log(f"  у межах {RETENTION_MONTHS} міс.: {len(found)} закупівель")
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
    """Статус конкретного лота: перемога, договір, зрив."""
    lots = {l["id"]: l for l in tender.get("lots", [])}
    lot = lots.get(lot_id)

    if lot:
        if lot.get("status") == "unsuccessful":
            return ST_FAILED, None, None
        if lot.get("status") == "cancelled":
            return ST_CANCELLED, None, None

    awards = [a for a in tender.get("awards", [])
              if (lot_id is None or a.get("lotID") == lot_id)]
    active = [a for a in awards if a.get("status") == "active"]

    supplier = None
    amount = None
    if active:
        a = active[-1]
        sup = (a.get("suppliers") or [{}])[0]
        supplier = sup.get("name")
        val = a.get("value") or {}
        amount = val.get("amount")

        for c in tender.get("contracts", []):
            if c.get("awardID") == a.get("id") and c.get("status") == "active":
                cv = c.get("value") or {}
                return ST_SIGNED, supplier, cv.get("amount", amount)
        return ST_WINNER, supplier, amount

    if awards and all(a.get("status") == "unsuccessful" for a in awards):
        return ST_FAILED, None, None

    tstatus = tender.get("status", "")
    if tstatus in ("unsuccessful",):
        return ST_FAILED, None, None
    if tstatus in ("cancelled",):
        return ST_CANCELLED, None, None
    return ST_PROGRESS, None, None


def matches_cpv(tender):
    codes = [((tender.get("classification") or {}).get("id") or "")]
    for it in tender.get("items", []):
        codes.append(((it.get("classification") or {}).get("id") or ""))
    return any(c.startswith(CPV_PREFIX) for c in codes)


_diag_left = 3


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


def item_codes(t):
    out = []
    for it in t.get("items", []):
        cls = (it.get("classification") or {}).get("id") or ""
        if cls:
            out.append(cls)
    return out


def parse_tender(t, meta=None):
    """Плоский опис закупівлі: позиції зі статусами лотів."""
    global _diag_left

    edr = entity_edrpou(t)
    codes = item_codes(t)
    root = ((t.get("classification") or {}).get("id") or "")

    if _diag_left > 0:
        _diag_left -= 1
        aw = t.get("awards") or []
        co = t.get("contracts") or []
        lo = t.get("lots") or []
        its = t.get("items") or []
        log(f"  [діаг] {t.get('tenderID')} · статус={t.get('status')} · title={str(t.get('title'))[:30]}")
        log(f"         lots={len(lo)} awards={len(aw)} contracts={len(co)} items={len(its)}")
        if aw:
            log(f"         award[0] ключі: {sorted(aw[0].keys())[:14]}")
            log(f"         award[0] status={aw[0].get('status')} lotID={aw[0].get('lotID')}")
        if co:
            log(f"         contract[0] ключі: {sorted(co[0].keys())[:14]}")
            log(f"         contract[0] status={co[0].get('status')} awardID={co[0].get('awardID')}")
        if lo:
            log(f"         lot[0] ключі: {sorted(lo[0].keys())[:12]} status={lo[0].get('status')}")
        if its:
            log(f"         item[0] ключі: {sorted(its[0].keys())[:14]}")
            log(f"         item[0] relatedLot={its[0].get('relatedLot')} descr={str(its[0].get('description'))[:40]}")
        if not aw and not co:
            log(f"         УСІ ключі: {sorted(t.keys())}")

    # пошук уже відфільтрував за buyer, тому розбіжність лише логуємо
    if edr and edr != EDRPOU:
        return None

    items = []
    for it in t.get("items", []):
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

    if not items:
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

    return {
        "id": t.get("id") or t.get("tenderID"),
        "tenderID": t.get("tenderID"),
        "title": (t.get("title") or (meta.get("title") if meta else "") or "").strip(),
        "date": t.get("date") or (meta.get("date") if meta else None),
        "dateModified": t.get("dateModified") or (meta.get("date") if meta else None),
        "method": t.get("procurementMethodType"),
        "status": t.get("status"),
        "contractEnd": contract_end,
        "url": f"{PORTAL}/tender/{t.get('tenderID')}",
        "counts": counts,
        "items": items,
    }


# ── читання деталей ─────────────────────────────────────────────────────────
# Знайдено дослідним шляхом: закупівля читається за публічним номером
# через портальний ендпоінт /api/tenders/{tenderID}/details.
DETAIL_URLS = [
    PORTAL + "/api/tenders/{id}/details",
]
_detail_url = None


def fetch_tender(tid):
    global _detail_url
    urls = [_detail_url] if _detail_url else DETAIL_URLS
    for u in urls:
        try:
            data = get_json(u.format(id=tid), tries=1)
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


def diff_statuses(old, parsed):
    """Що змінилось порівняно з попереднім запуском."""
    events = []
    for t in parsed:
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

def in_window(t, cutoff_iso, now_iso):
    if (t.get("dateModified") or "") >= cutoff_iso:
        return True
    ce = t.get("contractEnd")
    if ce and ce >= now_iso:
        return True
    return False


def main():
    log("=" * 60)
    log(f"Монітор закупівель · ЄДРПОУ {EDRPOU} · ДК {'/'.join(CPV_PREFIX)}*")
    log("=" * 60)

    state = load_state()
    known = dict(state.get("seen", {}))

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
    parsed = []
    fetched = 0
    skipped = 0
    errors = 0
    for tid in sorted(known, key=lambda k: known.get(k) or '', reverse=True):
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
            if fetched == 0 and errors == 0:
                diagnose_detail(tid)
                log("  деталі не читаються — далі не йдемо, дивись пробу вище")
                return 1
            errors += 1
            if errors <= 5:
                log(f"  {tid}: не вдалось прочитати")
            elif errors == 6:
                log("  …подальші помилки не друкую")
            continue
        fetched += 1
        if fetched % 25 == 0:
            log(f"  …{fetched} прочитано")
        t = resp.get("data") or resp or {}
        p = parse_tender(t, _search_meta.get(tid))
        if p:
            parsed.append(p)
        time.sleep(0.25)

    log(f"  завантажено: {fetched}, з кешу: {skipped}, помилок: {errors}")
    log(f"  підходять під фільтр: {len(parsed)}")

    if parsed:
        total_items = sum(len(p["items"]) for p in parsed)
        sig = sum(p["counts"]["signed"] for p in parsed)
        prg = sum(p["counts"]["progress"] for p in parsed)
        fail = sum(p["counts"]["failed"] for p in parsed)
        log(f"  позицій: {total_items} · підписано {sig} · в процесі {prg} · проблемних {fail}")

    log("\n[3] Порівняння зі станом")
    events = diff_statuses(state.get("tenders", {}), parsed)
    log(f"  подій: {len(events)}")
    for kind, t, item, was in events[:10]:
        if kind == "new":
            log(f"  NEW  {t['tenderID']} · {t['title'][:60]}")
        else:
            log(f"  CHG  {item['name'][:50]} · {was} → {item['status']}")

    first_run = not state.get("tenders")
    if first_run:
        log("  перший запуск — сповіщення не шлемо")
    else:
        notify(events)

    log("\n[4] Запис файлів")
    now = datetime.now(timezone.utc)
    cutoff = (now - timedelta(days=30 * RETENTION_MONTHS)).isoformat()
    now_iso = now.isoformat()

    window = [t for t in parsed if in_window(t, cutoff, now_iso)]
    window.sort(key=lambda t: t.get("dateModified") or "", reverse=True)
    log(f"  у зрізі за {RETENTION_MONTHS} міс.: {len(window)} з {len(parsed)}")

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    out = {
        "generated": now_iso,
        "edrpou": EDRPOU,
        "source": "prozorro.gov.ua",
        "tenders": window,
    }
    payload = json.dumps(out, ensure_ascii=False, indent=1)
    old_payload = ""
    if os.path.exists(OUT_PATH):
        with open(OUT_PATH, encoding="utf-8") as f:
            old_payload = f.read()

    changed = payload.strip() != old_payload.strip()
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        f.write(payload)
    log(f"  {OUT_PATH}: {len(payload) // 1024} КБ, змінено: {changed}")

    new_state = {"tenders": {}, "seen": {}}
    for t in parsed:
        new_state["tenders"][t["id"]] = {
            "dateModified": t.get("dateModified"),
            "items": {item_key(t["id"], i): i["status"] for i in t["items"]},
            "snapshot": t,
        }
        new_state["seen"][t["id"]] = t.get("dateModified", "")
    for tid, dm in known.items():
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
