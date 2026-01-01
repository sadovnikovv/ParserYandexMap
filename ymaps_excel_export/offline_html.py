# -*- coding: utf-8 -*-

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Tuple

from bs4 import BeautifulSoup

from .models import Company
from .utils import ANSI_RESET, ANSI_YELLOW, json_dumps_safe, log, safe_str


def iter_offline_html_files(input_path: str) -> List[Path]:
    p = Path(input_path)
    if p.is_dir():
        return sorted(p.glob("*.html"))
    if p.is_file():
        return [p]
    return []


def offline_html_is_scrolled_to_end(html: str) -> bool:
    """
    Грубая эвристика "страница пролистана до конца".
    """
    return 'class="add-business-view"' in (html or "")


def _text(el) -> str:
    if not el:
        return ""
    return safe_str(el.get_text(" ", strip=True))


def _digits(s: str) -> str:
    s = safe_str(s)
    return "".join(ch for ch in s if ch.isdigit())


def parse_side_panel_items(html: str) -> List[Dict[str, Any]]:
    """
    Парсит боковой список выдачи Яндекс.Карт из HTML.

    ВАЖНО:
    - В выдаче рядом с рейтингом чаще всего отображается именно количество ОЦЕНОК.
    - Количество отзывов чаще доступно на карточке организации (WEB-enrich).
    """
    soup = BeautifulSoup(html or "", "html.parser")
    nodes = soup.select('[data-object="search-list-item"][data-id]')
    items: List[Dict[str, Any]] = []

    for n in nodes:
        oid = safe_str(n.get("data-id"))
        coords = safe_str(n.get("data-coordinates"))  # "lon,lat"
        lon, lat = "", ""
        if coords and "," in coords:
            a, b = coords.split(",", 1)
            lon, lat = safe_str(a), safe_str(b)

        def first_by_class_contains(tag, needle: str):
            return tag.find(True, class_=lambda c: isinstance(c, str) and needle in c)

        title = _text(first_by_class_contains(n, "search-business-snippet-view__title")) or \
                _text(first_by_class_contains(n, "search-business-snippet-viewtitle"))

        address = _text(first_by_class_contains(n, "search-business-snippet-view__address")) or \
                  _text(first_by_class_contains(n, "search-business-snippet-viewaddress"))

        category = _text(first_by_class_contains(n, "search-business-snippet-view__category")) or \
                   _text(first_by_class_contains(n, "search-business-snippet-viewcategory"))

        worktime = _text(first_by_class_contains(n, "business-working-status-view"))

        rating_el = first_by_class_contains(n, "business-rating-badge-view__rating-text") or \
                    first_by_class_contains(n, "business-rating-badge-viewrating-text")
        rating = _text(rating_el).replace(",", ".")

        # --- Количество оценок / отзывов (если встречается в выдаче)
        rating_count = ""
        review_count = ""

        # Часто встречается span business-rating-amount-view _summary:
        # "598 оценок" / "123 отзыва"
        for el in n.find_all(True, class_=lambda c: isinstance(c, str) and "business-rating-amount-view" in c):
            t = _text(el).lower()
            d = _digits(t)
            if not d:
                continue
            if ("оцен" in t) and (not rating_count):
                rating_count = d
            if ("отзыв" in t) and (not review_count):
                review_count = d

        # Если текст без слов (например "(5725)") — считаем это количеством оценок.
        if not rating_count:
            cnt_el = first_by_class_contains(n, "business-rating-amount-view")
            cnt_text = _text(cnt_el)
            d = _digits(cnt_text)
            if d:
                rating_count = d

        href = ""
        a_overlay = n.select_one("a.link-overlay[href]")
        if a_overlay:
            href = safe_str(a_overlay.get("href"))
            if href.startswith("/"):
                href = "https://yandex.ru" + href

        items.append({
            "oid": oid,
            "title": title,
            "address": address,
            "worktime": worktime,
            "category": category,
            "rating": rating,
            "rating_count": rating_count,
            "review_count": review_count,
            "lon": lon,
            "lat": lat,
            "href": href,
        })

    # uniq by oid
    seen = set()
    out: List[Dict[str, Any]] = []
    for it in items:
        oid = safe_str(it.get("oid"))
        if not oid or oid in seen:
            continue
        seen.add(oid)
        out.append(it)

    return out


def build_companies_from_offline_html(html: str, source_name: str) -> Tuple[List[Company], Dict[str, Any]]:
    warnings: List[str] = []

    if not offline_html_is_scrolled_to_end(html):
        msg = "Похоже, выдача НЕ пролистана до конца (нет блока add-business-view)."
        log(f"[OFFLINE_HTML][WARN] {source_name}: {ANSI_YELLOW}{msg}{ANSI_RESET}")
        warnings.append(f"{source_name}: {msg}")

    items = parse_side_panel_items(html)
    log(f"[OFFLINE_HTML] {source_name}: items={len(items)}")

    companies: List[Company] = []
    for it in items:
        oid = safe_str(it.get("oid"))
        c = Company(
            ID=oid,
            Название=safe_str(it.get("title")),
            Адрес=safe_str(it.get("address")),
            Долгота=safe_str(it.get("lon")),
            Широта=safe_str(it.get("lat")),
            Режим_работы=safe_str(it.get("worktime")),
            Рейтинг=safe_str(it.get("rating")),
            Количество_оценок=safe_str(it.get("rating_count")),    # <-- ВАЖНО: оценки
            Количество_отзывов=safe_str(it.get("review_count")),    # обычно пусто и добирается WEB-enrich
            Категория_1=safe_str(it.get("category")),
            uri=f"ymapsbm1://org?oid={oid}" if oid else "",
            raw_json=json_dumps_safe({"source": source_name, "item": it}),
        )
        companies.append(c)

    meta = {"warnings": warnings, "items": len(items), "source": source_name}
    return companies, meta


def read_offline_input(input_path: str) -> Tuple[List[Company], Dict[str, Any], str]:
    files = iter_offline_html_files(input_path)
    if not files:
        return [], {"warnings": [f"OFFLINE_HTML_INPUT not found or no *.html: {input_path}"]}, "offline_html_missing"

    warnings: List[str] = []
    all_companies: List[Company] = []
    per_source: List[Dict[str, Any]] = []

    for fp in files:
        try:
            html = fp.read_text(encoding="utf-8", errors="ignore")
        except Exception as e:
            warnings.append(f"{fp.name}: read_error: {e}")
            continue

        companies, meta = build_companies_from_offline_html(html, source_name=fp.name)
        all_companies.extend(companies)
        per_source.append(meta)

    # uniq by ID
    seen = set()
    uniq: List[Company] = []
    for c in all_companies:
        oid = safe_str(c.ID)
        if not oid or oid in seen:
            continue
        seen.add(oid)
        uniq.append(c)

    meta_all = {"warnings": warnings, "sources": per_source, "rows": len(uniq)}
    return uniq, meta_all, ""
