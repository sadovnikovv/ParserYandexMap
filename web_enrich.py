# -*- coding: utf-8 -*-

import json
import re
import time
from typing import Any, Dict, List, Tuple

import requests
from bs4 import BeautifulSoup

from settings import Settings
from utils import safe_str, dedup_keep_order, pick_n, apply_field, log


_PHONE_RE = re.compile(r"\+?\d[\d\s().-]{6,}\d")


def normalize_phone_ru(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    digits = re.sub(r"\D+", "", s)
    if len(digits) == 11 and digits.startswith("8"):
        digits = "7" + digits[1:]
    if len(digits) == 10:
        digits = "7" + digits
    if len(digits) == 11 and digits.startswith("7"):
        return "+" + digits
    return s


def normalize_http_url(href: str) -> str:
    href = (href or "").strip()
    if not href:
        return ""
    if href.startswith("https") and "://" not in href:
        href = href.replace("https", "https://", 1)
    elif href.startswith("http") and "://" not in href:
        href = href.replace("http", "http://", 1)
    return href


def extract_jsonld_blocks(html: str) -> List[Any]:
    blocks = re.findall(
        r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        html or "",
        flags=re.DOTALL | re.IGNORECASE,
    )
    out: List[Any] = []
    for b in blocks:
        b = (b or "").strip()
        if not b:
            continue
        try:
            out.append(json.loads(b))
        except Exception:
            continue
    return out


def walk_find(obj: Any, key: str) -> List[Any]:
    found: List[Any] = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k == key:
                found.append(v)
            found.extend(walk_find(v, key))
    elif isinstance(obj, list):
        for it in obj:
            found.extend(walk_find(it, key))
    return found


def parse_web_jsonld_rating(jsonlds: List[Any]) -> Dict[str, Any]:
    rating_value = ""
    review_count = ""
    for o in jsonlds:
        for agg in walk_find(o, "aggregateRating"):
            if isinstance(agg, dict):
                if not rating_value and agg.get("ratingValue") is not None:
                    rating_value = safe_str(agg.get("ratingValue"))
                if not review_count and agg.get("reviewCount") is not None:
                    review_count = safe_str(agg.get("reviewCount"))
    return {"ratingValue": rating_value, "reviewCount": review_count}


def parse_web_jsonld_profile(jsonlds: List[Any]) -> Dict[str, Any]:
    best = {"name": "", "url": "", "telephone": [], "address": "", "postalCode": ""}

    def consider(obj: dict):
        nonlocal best
        name = safe_str(obj.get("name"))
        url = safe_str(obj.get("url"))
        tel = obj.get("telephone")
        addr = obj.get("address")

        tel_list: List[str] = []
        if isinstance(tel, str):
            tel_list = [normalize_phone_ru(tel)]
        elif isinstance(tel, list):
            tel_list = [normalize_phone_ru(safe_str(x)) for x in tel if safe_str(x)]

        addr_str = ""
        postal = ""
        if isinstance(addr, str):
            addr_str = safe_str(addr)
        elif isinstance(addr, dict):
            postal = safe_str(addr.get("postalCode"))
            parts = [safe_str(addr.get("addressLocality")), safe_str(addr.get("streetAddress"))]
            addr_str = ", ".join([p for p in parts if p])

        score = int(bool(name)) + int(bool(url)) + int(bool(addr_str)) + len(tel_list)
        cur = int(bool(best["name"])) + int(bool(best["url"])) + int(bool(best["address"])) + len(best["telephone"])
        if score > cur:
            best = {
                "name": name,
                "url": url,
                "telephone": dedup_keep_order(tel_list),
                "address": addr_str,
                "postalCode": postal,
            }

    def walk(x: Any):
        if isinstance(x, dict):
            t = x.get("@type")
            if isinstance(t, str) and t.lower() in ("localbusiness", "organization", "place"):
                consider(x)
            if isinstance(t, list) and any(isinstance(z, str) and z.lower() in ("localbusiness", "organization", "place") for z in t):
                consider(x)
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for it in x:
                walk(it)

    for o in jsonlds:
        walk(o)
    return best


def extract_contacts_fast_bs4(html: str) -> Dict[str, Any]:
    soup = BeautifulSoup(html or "", "html.parser")

    # PHONES
    tels = []

    for a in soup.select('a[href^="tel:"]'):
        href = (a.get("href") or "").strip()
        tel = href.replace("tel:", "").strip()
        if tel:
            tels.append(normalize_phone_ru(tel))

    for node in soup.select('[itemprop="telephone"]'):
        t = node.get_text(" ", strip=True)
        if t:
            tels.append(normalize_phone_ru(t))

    if not tels:
        meta_desc = soup.select_one('meta[name="description"][content]')
        if meta_desc:
            content = (meta_desc.get("content") or "").strip()
            for m in _PHONE_RE.findall(content):
                tels.append(normalize_phone_ru(m))

    # EMAILS
    mails = []
    for a in soup.select('a[href^="mailto:"]'):
        href = (a.get("href") or "").strip()
        em = href.replace("mailto:", "").strip()
        if em:
            mails.append(em)

    # SITE
    site = ""
    a_url = soup.select_one('a[itemprop="url"][href]')
    if a_url:
        site = normalize_http_url(a_url.get("href"))

    if not site:
        for a in soup.select('a[href^="http"]'):
            txt = a.get_text(" ", strip=True).lower()
            if "сайт" in txt:
                site = normalize_http_url(a.get("href"))
                break

    return {
        "telephone_list": dedup_keep_order([x for x in tels if x]),
        "email_list": dedup_keep_order([x for x in mails if x]),
        "url": site,
    }


def requests_is_blocked(final_url: str, html: str) -> bool:
    u = (final_url or "").lower()
    if "showcaptcha" in u:
        return True
    soup = BeautifulSoup(html or "", "html.parser")
    if soup.select_one("form[action*='showcaptcha']") is not None:
        return True
    if soup.select_one("iframe[src*='captcha'], iframe[src*='showcaptcha']") is not None:
        return True
    text = soup.get_text(" ", strip=True).lower()
    if "подтвердите, что запросы отправляли вы" in text:
        return True
    return False


def http_get_org_page_requests(session: requests.Session, oid: str) -> Tuple[str, str]:
    url = f"https://yandex.ru/maps/org/{oid}/"
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
        ),
        "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
    }

    last_exc = None
    for attempt in range(1, 4):
        try:
            r = session.get(url, headers=headers, timeout=25, allow_redirects=True)
            html = r.text or ""
            final_url = r.url or url

            if r.status_code in (429, 500, 502, 503, 504):
                log(f"[WEB][REQ] retryable status={r.status_code} attempt={attempt}/3 oid={oid}")
                time.sleep(1.3 * attempt)
                continue

            if r.status_code >= 400:
                raise RuntimeError(f"WEB HTTP {r.status_code} final_url={final_url}")

            return final_url, html

        except (requests.Timeout, requests.ConnectionError) as e:
            last_exc = e
            time.sleep(1.3 * attempt)

    raise RuntimeError(f"requests failed oid={oid} last_exc={last_exc}")


def enrich_from_web_html(html: str) -> Dict[str, Any]:
    jsonlds = extract_jsonld_blocks(html)
    rating_obj = parse_web_jsonld_rating(jsonlds)
    profile_obj = parse_web_jsonld_profile(jsonlds)
    fast = extract_contacts_fast_bs4(html)

    tels = profile_obj.get("telephone") or []
    if not tels:
        tels = fast.get("telephone_list") or []

    emails = fast.get("email_list") or []
    url = safe_str(profile_obj.get("url")) or safe_str(fast.get("url"))

    return {
        "rating": safe_str(rating_obj.get("ratingValue")),
        "review_count": safe_str(rating_obj.get("reviewCount")),
        "name": safe_str(profile_obj.get("name")),
        "url": url,
        "telephone_list": dedup_keep_order([safe_str(x) for x in tels if safe_str(x)]),
        "email_list": dedup_keep_order([safe_str(x) for x in emails if safe_str(x)]),
        "address": safe_str(profile_obj.get("address")),
        "postal": safe_str(profile_obj.get("postalCode")),
        "jsonld_blocks": len(jsonlds),
    }


def enrich_rows_web(st: Settings, rows: List[dict], selenium_pool) -> Tuple[List[dict], Dict[str, Any]]:
    stats = {"attempted": 0, "success": 0, "failed": 0, "skipped": 0, "errors": []}
    done = 0

    with requests.Session() as session:
        for i, row in enumerate(rows, start=1):
            oid = safe_str(row.get("ID"))
            if not oid:
                stats["skipped"] += 1
                continue
            if done >= st.WEB_MAX_ITEMS:
                stats["skipped"] += 1
                continue

            log(f"[ENRICH] row={i}/{len(rows)} oid={oid} mode=WEB")
            stats["attempted"] += 1
            done += 1

            try:
                # 1) быстрый requests
                final_url, html = http_get_org_page_requests(session, oid)

                # 2) если блокировка — только тогда Selenium
                if requests_is_blocked(final_url, html):
                    log(f"[ENRICH][WEB][FAST] blocked (final_url={final_url}), switching to selenium oid={oid}")
                    html = selenium_pool.get_page_html(f"https://yandex.ru/maps/org/{oid}/")

                # 3) парсим
                w = enrich_from_web_html(html)

                # 4) если совсем пусто по нужным полям — fallback на Selenium (дождаться React-частей)
                has_any = bool(w.get("telephone_list") or w.get("url") or w.get("rating") or w.get("review_count"))
                if not has_any:
                    log(f"[ENRICH][WEB][FAST] extracted empty, fallback to selenium oid={oid}")
                    html = selenium_pool.get_page_html(f"https://yandex.ru/maps/org/{oid}/")
                    w = enrich_from_web_html(html)

                changed = 0
                changed += int(apply_field(row, "Название", w.get("name", ""), overwrite=st.WEB_FORCE_OVERWRITE))
                changed += int(apply_field(row, "Адрес", w.get("address", ""), overwrite=st.WEB_FORCE_OVERWRITE))
                changed += int(apply_field(row, "Индекс", w.get("postal", ""), overwrite=st.WEB_FORCE_OVERWRITE))
                changed += int(apply_field(row, "Сайт", w.get("url", ""), overwrite=st.WEB_FORCE_OVERWRITE))
                changed += int(apply_field(row, "Рейтинг", w.get("rating", ""), overwrite=st.WEB_FORCE_OVERWRITE))
                changed += int(apply_field(row, "Количество отзывов", w.get("review_count", ""), overwrite=st.WEB_FORCE_OVERWRITE))

                tels = pick_n(list(w.get("telephone_list") or []), 3)
                changed += int(apply_field(row, "Телефон 1", tels[0], overwrite=st.WEB_FORCE_OVERWRITE))
                changed += int(apply_field(row, "Телефон 2", tels[1], overwrite=st.WEB_FORCE_OVERWRITE))
                changed += int(apply_field(row, "Телефон 3", tels[2], overwrite=st.WEB_FORCE_OVERWRITE))

                emails = pick_n(list(w.get("email_list") or []), 3)
                changed += int(apply_field(row, "Email 1", emails[0], overwrite=st.WEB_FORCE_OVERWRITE))
                changed += int(apply_field(row, "Email 2", emails[1], overwrite=st.WEB_FORCE_OVERWRITE))
                changed += int(apply_field(row, "Email 3", emails[2], overwrite=st.WEB_FORCE_OVERWRITE))

                # отметка в raw_json
                try:
                    raw = {}
                    try:
                        raw = json.loads(row.get("raw_json") or "{}")
                    except Exception:
                        raw = {"raw_json_parse_error": True}
                    raw["web_enrich"] = {"ok": True, "changed_fields": changed, "jsonld_blocks": w.get("jsonld_blocks")}
                    row["raw_json"] = json.dumps(raw, ensure_ascii=False)
                except Exception:
                    pass

                stats["success"] += 1
                log(f"[ENRICH][WEB] ok oid={oid} changed_fields={changed}")
                time.sleep(st.SLEEP_SEC)

            except Exception as e:
                stats["failed"] += 1
                stats["errors"].append(f"oid={oid}: {e}")
                log(f"[ENRICH][WEB] FAIL oid={oid}: {e}")

    return rows, stats
