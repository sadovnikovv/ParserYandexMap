# -*- coding: utf-8 -*-

from pathlib import Path
from typing import Any, Dict, List, Tuple

from bs4 import BeautifulSoup

from utils import safe_str, json_dumps_safe, log, ANSI_YELLOW, ANSI_RESET


def iter_offline_html_files(input_path: str) -> List[Path]:
    p = Path(input_path)
    if p.is_dir():
        return sorted(p.glob("*.html"))
    if p.is_file():
        return [p]
    return []


def offline_html_is_scrolled_to_end(html: str) -> bool:
    # очень грубый маркер, но полезен как предупреждение
    return 'class="add-business-view"' in (html or "")


def _text(el) -> str:
    if not el:
        return ""
    return safe_str(el.get_text(" ", strip=True))


def parse_side_panel_items(html: str) -> List[Dict[str, Any]]:
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

        title = _text(first_by_class_contains(n, "search-business-snippet-view__title")) or _text(first_by_class_contains(n, "search-business-snippet-viewtitle"))
        address = _text(first_by_class_contains(n, "search-business-snippet-view__address")) or _text(first_by_class_contains(n, "search-business-snippet-viewaddress"))
        category = _text(first_by_class_contains(n, "search-business-snippet-view__category")) or _text(first_by_class_contains(n, "search-business-snippet-viewcategory"))
        worktime = _text(first_by_class_contains(n, "business-working-status-view"))

        rating = ""
        r_el = first_by_class_contains(n, "business-rating-badge-view__rating-text") or first_by_class_contains(n, "business-rating-badge-viewrating-text")
        rating = _text(r_el).replace(",", ".")

        count = ""
        cnt_el = first_by_class_contains(n, "business-rating-amount-view")
        cnt_text = _text(cnt_el)
        # "27 отзывов" -> 27
        digits = "".join([ch for ch in cnt_text if ch.isdigit()])
        if digits:
            count = digits

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
            "count": count,
            "lon": lon,
            "lat": lat,
            "href": href,
        })

    # uniq by oid
    seen = set()
    out = []
    for it in items:
        oid = safe_str(it.get("oid"))
        if not oid or oid in seen:
            continue
        seen.add(oid)
        out.append(it)
    return out


def build_rows_from_offline_html(html: str, source_name: str) -> Tuple[List[dict], Dict[str, Any]]:
    warnings: List[str] = []
    if not offline_html_is_scrolled_to_end(html):
        msg = "похоже, выдача НЕ пролистана до конца (нет блока add-business-view)."
        log(f"[OFFLINE_HTML][WARN] {source_name}: {ANSI_YELLOW}{msg}{ANSI_RESET}")
        warnings.append(f"{source_name}: {msg}")

    items = parse_side_panel_items(html)
    log(f"[OFFLINE_HTML] {source_name}: боковой список items={len(items)}")

    rows: List[dict] = []
    for it in items:
        oid = safe_str(it.get("oid"))
        row = {
            "ID": oid,
            "Название": safe_str(it.get("title")),
            "Адрес": safe_str(it.get("address")),
            "Индекс": "",
            "Долгота": safe_str(it.get("lon")),
            "Широта": safe_str(it.get("lat")),
            "Сайт": "",
            "Телефон 1": "", "Телефон 2": "", "Телефон 3": "",
            "Email 1": "", "Email 2": "", "Email 3": "",
            "Режим работы": safe_str(it.get("worktime")),
            "Рейтинг": safe_str(it.get("rating")),
            "Количество отзывов": safe_str(it.get("count")),
            "Категория 1": safe_str(it.get("category")),
            "Категория 2": "",
            "Категория 3": "",
            "Особенности": "",
            "uri": f"ymapsbm1://org?oid={oid}",
            "Факс 1": "", "Факс 2": "", "Факс 3": "",
            "Категории (прочие)": "",
            "raw_json": json_dumps_safe({"source": source_name, "item": it}),
        }
        rows.append(row)

    meta = {"warnings": warnings, "items": len(items), "source": source_name}
    return rows, meta


def read_offline_input(input_path: str) -> Tuple[List[dict], Dict[str, Any], str]:
    files = iter_offline_html_files(input_path)
    if not files:
        return [], {"warnings": [f"OFFLINE_HTML_INPUT not found or no *.html: {input_path}"]}, "offline_html_missing"

    warnings: List[str] = []
    all_rows: List[dict] = []
    per_source: List[Dict[str, Any]] = []

    for fp in files:
        try:
            html = fp.read_text(encoding="utf-8", errors="ignore")
        except Exception as e:
            warnings.append(f"{fp.name}: read_error: {e}")
            continue

        rows, meta = build_rows_from_offline_html(html, source_name=fp.name)
        all_rows.extend(rows)
        per_source.append(meta)

    # uniq
    seen = set()
    uniq: List[dict] = []
    for r in all_rows:
        oid = safe_str(r.get("ID"))
        if not oid or oid in seen:
            continue
        seen.add(oid)
        uniq.append(r)

    meta_all = {"warnings": warnings, "sources": per_source, "rows": len(uniq)}
    return uniq, meta_all, ""
