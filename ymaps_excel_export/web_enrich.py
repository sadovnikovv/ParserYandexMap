# -*- coding: utf-8 -*-

from __future__ import annotations

import json
import re
import time
from typing import Any, Dict, List, Tuple

import requests
from bs4 import BeautifulSoup

from .config import Settings
from .models import Company
from .selenium_pool import SeleniumPool
from .utils import dedup_keep_order, json_dumps_safe, log, pick_n, safe_str

_PHONE_RE = re.compile(r"\+?\d[\d\s().-]{6,}\d")

_RE_RATING_VALUE = re.compile(r'"ratingValue"\s*:\s*"?([0-9]+(?:[.,][0-9]+)?)"?', re.I)
_RE_REVIEW_COUNT = re.compile(r'"reviewCount"\s*:\s*"?(\d{1,10})"?', re.I)
_RE_REVIEWS_ALT = re.compile(r'"reviewsCount"\s*:\s*"?(\d{1,10})"?', re.I)

# НОВОЕ: кол-во оценок (ratingCount / ratingsCount)
_RE_RATING_COUNT = re.compile(r'"ratingCount"\s*:\s*"?(\d{1,10})"?', re.I)
_RE_RATINGS_COUNT_ALT = re.compile(r'"ratingsCount"\s*:\s*"?(\d{1,10})"?', re.I)

_RE_RATING_ALT = re.compile(r'"rating"\s*:\s*"?([0-9]+(?:[.,][0-9]+)?)"?', re.I)

_RE_HOURS_TEXT_1 = re.compile(r'"Hours"\s*:\s*\{[^{}]*"text"\s*:\s*"([^"]{3,200})"', re.I)
_RE_HOURS_TEXT_2 = re.compile(r'"hours"\s*:\s*\{[^{}]*"text"\s*:\s*"([^"]{3,200})"', re.I)


def normalize_phone_ru(s: str) -> str:
    s = safe_str(s)
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


def _format_rating_1(x: Any) -> str:
    s = safe_str(x).replace(",", ".")
    if not s:
        return ""
    try:
        v = float(s)
    except Exception:
        return ""
    return f"{v:.1f}".replace(".", ",")


def _digits(s: str) -> str:
    s = safe_str(s)
    return "".join(ch for ch in s if ch.isdigit())


def extract_jsonld_blocks(html: str) -> List[Any]:
    blocks = re.findall(
        r"<script[^>]+type=['\"]application/ld\+json['\"][^>]*>(.*?)</script>",
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


def extract_embedded_json_objects(html: str) -> List[Any]:
    """
    Часто у Яндекса встречаются <script type="application/json">...</script>
    и/или window.__SOME_STATE__ = {...}
    """
    out: List[Any] = []

    for b in re.findall(
        r"<script[^>]+type=['\"]application/json['\"][^>]*>(.*?)</script>",
        html or "",
        flags=re.DOTALL | re.IGNORECASE,
    ):
        b = (b or "").strip()
        if not b:
            continue
        try:
            out.append(json.loads(b))
        except Exception:
            continue

    for b in re.findall(
        r"window\.__[A-Z0-9_]{3,}\s*=\s*({.*?})\s*;\s*",
        html or "",
        flags=re.DOTALL,
    ):
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


def parse_rating_counts_from_jsonld(html: str) -> Tuple[str, str, str]:
    """
    Возвращает:
      rating_value, rating_count(оценки), review_count(отзывы)
    """
    rating_value = ""
    rating_count = ""
    review_count = ""

    for o in extract_jsonld_blocks(html):
        for agg in walk_find(o, "aggregateRating"):
            if not isinstance(agg, dict):
                continue

            if not rating_value and agg.get("ratingValue") is not None:
                rating_value = _format_rating_1(agg.get("ratingValue"))

            if not rating_count and agg.get("ratingCount") is not None:
                rating_count = safe_str(agg.get("ratingCount"))

            if not review_count and agg.get("reviewCount") is not None:
                review_count = safe_str(agg.get("reviewCount"))

    # Regex fallback
    if not rating_value:
        m = _RE_RATING_VALUE.search(html or "") or _RE_RATING_ALT.search(html or "")
        if m:
            rating_value = _format_rating_1(m.group(1))

    if not rating_count:
        m = _RE_RATING_COUNT.search(html or "") or _RE_RATINGS_COUNT_ALT.search(html or "")
        if m:
            rating_count = safe_str(m.group(1))

    if not review_count:
        m = _RE_REVIEW_COUNT.search(html or "") or _RE_REVIEWS_ALT.search(html or "")
        if m:
            review_count = safe_str(m.group(1))

    return rating_value, rating_count, review_count


def parse_rating_counts_from_embedded_json(html: str) -> Tuple[str, str, str]:
    """
    Возвращает:
      rating_value, rating_count(оценки), review_count(отзывы)
    """
    rating_value = ""
    rating_count = ""
    review_count = ""

    for obj in extract_embedded_json_objects(html):
        if not rating_value:
            for v in walk_find(obj, "ratingValue"):
                s = _format_rating_1(v)
                if s:
                    rating_value = s
                    break

        if not review_count:
            for v in walk_find(obj, "reviewCount"):
                s = safe_str(v)
                if s.isdigit():
                    review_count = s
                    break

        if not review_count:
            for v in walk_find(obj, "reviewsCount"):
                s = safe_str(v)
                if s.isdigit():
                    review_count = s
                    break

        if not rating_count:
            for v in walk_find(obj, "ratingCount"):
                s = safe_str(v)
                if s.isdigit():
                    rating_count = s
                    break

        if not rating_count:
            for v in walk_find(obj, "ratingsCount"):
                s = safe_str(v)
                if s.isdigit():
                    rating_count = s
                    break

        # иногда rating лежит просто как "rating"
        if not rating_value:
            for v in walk_find(obj, "rating"):
                s = _format_rating_1(v)
                if s:
                    rating_value = s
                    break

        if rating_value and rating_count and review_count:
            return rating_value, rating_count, review_count

    # Regex fallback (на всякий)
    if not rating_value:
        m = _RE_RATING_VALUE.search(html or "") or _RE_RATING_ALT.search(html or "")
        if m:
            rating_value = _format_rating_1(m.group(1))

    if not rating_count:
        m = _RE_RATING_COUNT.search(html or "") or _RE_RATINGS_COUNT_ALT.search(html or "")
        if m:
            rating_count = safe_str(m.group(1))

    if not review_count:
        m = _RE_REVIEW_COUNT.search(html or "") or _RE_REVIEWS_ALT.search(html or "")
        if m:
            review_count = safe_str(m.group(1))

    return rating_value, rating_count, review_count


def parse_counts_from_dom(html: str) -> Tuple[str, str]:
    """
    DOM-эвристика:
      - business-rating-amount-view может содержать "XXXX оценок" и/или "YY отзывов"
      - иногда встречается "(5725)" — трактуем как оценки
    """
    soup = BeautifulSoup(html or "", "html.parser")

    rating_count = ""
    review_count = ""

    for el in soup.find_all(True, class_=lambda c: isinstance(c, str) and "business-rating-amount-view" in c):
        t = safe_str(el.get_text(" ", strip=True)).lower()
        d = _digits(t)
        if not d:
            continue
        if ("оцен" in t) and (not rating_count):
            rating_count = d
        if ("отзыв" in t) and (not review_count):
            review_count = d

    if not rating_count:
        el2 = soup.select_one(".business-rating-with-text-view__count")
        if el2:
            d2 = _digits(el2.get_text(" ", strip=True))
            if d2:
                rating_count = d2

    return rating_count, review_count


def parse_worktime_from_html(html: str) -> str:
    m = _RE_HOURS_TEXT_1.search(html or "") or _RE_HOURS_TEXT_2.search(html or "")
    if m:
        return safe_str(m.group(1))

    soup = BeautifulSoup(html or "", "html.parser")
    for sel in (
        ".business-working-status-view__text",
        ".business-working-status-view",
        "[class*=business-working-status-view]",
    ):
        n = soup.select_one(sel)
        if n is not None:
            t = safe_str(n.get_text(" ", strip=True))
            if t and len(t) < 200:
                return t
    return ""


def parse_web_contacts_fast(html: str) -> Dict[str, Any]:
    soup = BeautifulSoup(html or "", "html.parser")

    # Phones
    tels: List[str] = []
    for a in soup.select('a[href^="tel:"]'):
        href = safe_str(a.get("href"))
        tel = href.replace("tel:", "").strip()
        if tel:
            tels.append(normalize_phone_ru(tel))

    for node in soup.select('[itemprop="telephone"]'):
        t = safe_str(node.get_text(" ", strip=True))
        if t:
            tels.append(normalize_phone_ru(t))

    meta_desc = soup.select_one('meta[name="description"][content]')
    if meta_desc is not None:
        content = safe_str(meta_desc.get("content"))
        for m in _PHONE_RE.findall(content):
            tels.append(normalize_phone_ru(m))

    # Emails
    emails: List[str] = []
    for a in soup.select('a[href^="mailto:"]'):
        href = safe_str(a.get("href"))
        em = href.replace("mailto:", "").strip()
        if em:
            emails.append(em)

    # Site
    site = ""
    aurl = soup.select_one('a[itemprop="url"][href]')
    if aurl is not None:
        site = safe_str(aurl.get("href"))

    return {
        "telephones": dedup_keep_order([x for x in tels if safe_str(x)]),
        "emails": dedup_keep_order([x for x in emails if safe_str(x)]),
        "site": site,
    }


def requests_is_blocked(final_url: str, html: str) -> bool:
    u = safe_str(final_url).lower()
    if "showcaptcha" in u:
        return True

    soup = BeautifulSoup(html or "", "html.parser")
    if soup.select_one("form[action*='showcaptcha']") is not None:
        return True
    if soup.select_one("iframe[src*='captcha'], iframe[src*='showcaptcha']") is not None:
        return True

    txt = safe_str(soup.get_text(" ", strip=True)).lower()
    if "подтвердите, что запросы отправляли вы" in txt:
        return True

    return False


def http_get_org_page(session: requests.Session, oid: str, timeout_sec: int) -> Tuple[str, str]:
    url = f"https://yandex.ru/maps/org/{oid}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36",
        "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
    }

    last_exc: Exception | None = None
    for attempt in range(1, 4):
        try:
            r = session.get(url, headers=headers, timeout=timeout_sec, allow_redirects=True)
            html = r.text or ""
            final_url = r.url or url

            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(0.6 * attempt)
                continue

            if r.status_code >= 400:
                raise RuntimeError(f"WEB HTTP {r.status_code} final_url={final_url}")

            return final_url, html

        except (requests.Timeout, requests.ConnectionError) as e:
            last_exc = e
            time.sleep(0.6 * attempt)

    raise RuntimeError(f"requests failed oid={oid} err={last_exc}")


def set_if_needed(c: Company, attr: str, value: str, overwrite: bool) -> int:
    val = safe_str(value)
    if not val:
        return 0
    cur = safe_str(getattr(c, attr))
    if cur and not overwrite:
        return 0
    setattr(c, attr, val)
    return 1


def enrich_company_from_web(st: Settings, c: Company, pool: SeleniumPool, session: requests.Session) -> Tuple[Company, Dict[str, Any]]:
    stats: Dict[str, Any] = {"mode": "WEB"}

    oid = safe_str(getattr(c, "ID", ""))
    if not oid.isdigit():
        stats["skipped"] = "nondigit_oid"
        return c, stats

    used_selenium = False

    final_url, html = http_get_org_page(session, oid, timeout_sec=st.WEB_TIMEOUT_SEC)
    if requests_is_blocked(final_url, html):
        html = pool.get_page_html(f"https://yandex.ru/maps/org/{oid}")
        used_selenium = True

    contacts = parse_web_contacts_fast(html)

    # 1) JSON-LD
    rating_value, rating_count, review_count = parse_rating_counts_from_jsonld(html)

    # 2) embedded-json fallback
    if not rating_value or not rating_count or not review_count:
        r2, rc2, rv2 = parse_rating_counts_from_embedded_json(html)
        rating_value = rating_value or r2
        rating_count = rating_count or rc2
        review_count = review_count or rv2

    # 3) DOM fallback (русские "оценок/отзывов" + "(число)")
    dom_rating_count, dom_review_count = parse_counts_from_dom(html)
    rating_count = rating_count or dom_rating_count
    review_count = review_count or dom_review_count

    worktime = parse_worktime_from_html(html)

    changed = 0
    overwrite = bool(st.WEB_FORCE_OVERWRITE)

    changed += set_if_needed(c, "Сайт", safe_str(contacts.get("site")), overwrite=overwrite)

    tels = pick_n(contacts.get("telephones") or [], st.MAX_PHONES)
    changed += set_if_needed(c, "Телефон_1", tels[0], overwrite=overwrite)
    changed += set_if_needed(c, "Телефон_2", tels[1], overwrite=overwrite)
    changed += set_if_needed(c, "Телефон_3", tels[2], overwrite=overwrite)

    emails = pick_n(contacts.get("emails") or [], st.MAX_EMAILS)
    changed += set_if_needed(c, "Email_1", emails[0], overwrite=overwrite)
    changed += set_if_needed(c, "Email_2", emails[1], overwrite=overwrite)
    changed += set_if_needed(c, "Email_3", emails[2], overwrite=overwrite)

    changed += set_if_needed(c, "Рейтинг", rating_value, overwrite=overwrite)
    changed += set_if_needed(c, "Количество_оценок", rating_count, overwrite=overwrite)   # <-- НОВОЕ
    changed += set_if_needed(c, "Количество_отзывов", review_count, overwrite=overwrite)

    # Режим работы НЕ перезаписываем агрессивно (обычно он уже есть из выдачи)
    changed += set_if_needed(c, "Режим_работы", worktime, overwrite=False)

    # raw_json trace
    try:
        raw = {}
        try:
            raw = json.loads(safe_str(getattr(c, "raw_json", "")) or "{}")
        except Exception:
            raw = {"raw_json_parse_error": True, "raw_json_raw": safe_str(getattr(c, "raw_json", ""))[:200]}

        if not isinstance(raw, dict):
            raw = {"raw_json_not_dict": True}

        raw["web_enrich"] = {
            "ok": True,
            "used_selenium": used_selenium,
            "final_url": final_url,
            "rating": rating_value,
            "rating_count": rating_count,
            "review_count": review_count,
            "changed": changed,
        }
        setattr(c, "raw_json", json_dumps_safe(raw))
    except Exception:
        pass

    stats["ok"] = True
    stats["used_selenium"] = used_selenium
    stats["changed"] = changed
    return c, stats


def enrich_companies_web(st: Settings, companies: List[Company], pool: SeleniumPool) -> Dict[str, Any]:
    stats: Dict[str, Any] = {"attempted": 0, "success": 0, "failed": 0, "skipped": 0, "errors": []}
    done = 0

    with requests.Session() as session:
        for i, c in enumerate(companies, start=1):
            if st.WEB_MAX_ITEMS > 0 and done >= st.WEB_MAX_ITEMS:
                stats["skipped"] += 1
                continue

            oid = safe_str(getattr(c, "ID", ""))
            if not oid:
                stats["skipped"] += 1
                continue

            done += 1
            stats["attempted"] += 1

            if st.VERBOSE:
                log(f"[ENRICH] {i}/{len(companies)} oid={oid} mode=WEB")

            try:
                newc, _ = enrich_company_from_web(st, c, pool, session)
                companies[i - 1] = newc
                stats["success"] += 1
            except Exception as e:
                stats["failed"] += 1
                stats["errors"].append(f"{oid}: {e}")

            if st.SLEEP_SEC > 0:
                time.sleep(float(st.SLEEP_SEC))

    return stats
