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
# сервісу порталу. Він не задокументований, тому пробуємо кілька форматів.

SEARCH_CANDIDATES = [
    PORTAL + "/api/search/tenders?edrpou={e}&page={p}",
    PORTAL + "/api/search/tenders?text={e}&page={p}",
    PORTAL + "/api/search/tenders?query={e}&page={p}",
]


def discover_tender_ids():
    found = {}
    working = None
    for tpl in SEARCH_CANDIDATES:
        try:
            probe = tpl.format(e=EDRPOU, p=0)
            data = get_json(probe, tries=1)
        except Exception as ex:
            log(f"  пошук {tpl.split('?')[1].split('=')[0]}: недоступний ({ex})")
            continue
        if not isinstance(data, dict):
            continue
        rows = data.get("data") or data.get("items") or data.get("results") or []
        if rows:
            working = tpl
            log(f"  пошуковий ендпоінт працює: {tpl}")
            break
        log(f"  пошук {tpl.split('?')[1].split('=')[0]}: порожня відповідь")

    if not working:
        log("  ЖОДЕН пошуковий ендпоінт не відповів — працюємо лише з відомими id")
        return found

    page = 0
    while page < 40:
        try:
            data = get_json(working.format(e=EDRPOU, p=page))
        except Exception as ex:
            log(f"  сторінка {page}: помилка {ex}")
            break
        rows = data.get("data") or data.get("items") or data.get("results") or []
        if not rows:
            break
        for r in rows:
            tid = r.get("id") or r.get("tenderID")
            if tid:
                found[tid] = r.get("dateModified", "")
        page += 1
        time.sleep(0.4)

    log(f"  знайдено закупівель у пошуку: {len(found)}")
    return found


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


def parse_tender(t):
    """Плоский опис закупівлі: позиції зі статусами лотів."""
    entity = t.get("procuringEntity") or {}
    ident = entity.get("identifier") or {}
    if ident.get("id") != EDRPOU:
        return None
    if not matches_cpv(t):
        return None

    items = []
    for it in t.get("items", []):
        cls = (it.get("classification") or {}).get("id") or ""
        if not cls.startswith(CPV_PREFIX):
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
        "id": t.get("id"),
        "tenderID": t.get("tenderID"),
        "title": (t.get("title") or "").strip(),
        "date": t.get("date"),
        "dateModified": t.get("dateModified"),
        "method": t.get("procurementMethodType"),
        "status": t.get("status"),
        "contractEnd": contract_end,
        "url": f"{PORTAL}/tender/{t.get('tenderID')}",
        "counts": counts,
        "items": items,
    }


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
    for tid in list(known.keys()):
        newer = discovered.get(tid, "")
        cached = state.get("tenders", {}).get(tid, {}).get("dateModified", "")
        if newer and cached and newer <= cached and tid in state.get("tenders", {}):
            snap = state["tenders"][tid].get("snapshot")
            if snap:
                parsed.append(snap)
                skipped += 1
                continue
        try:
            resp = get_json(f"{API}/tenders/{tid}")
            fetched += 1
        except Exception as ex:
            log(f"  {tid}: помилка {ex}")
            errors += 1
            continue
        t = resp.get("data") or {}
        p = parse_tender(t)
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
