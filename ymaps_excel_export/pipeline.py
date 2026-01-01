# pipeline.py
# -*- coding: utf-8 -*-

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List

from .config import Settings
from .excel_writer import save_to_excel
from .models import Company, RunResult
from .offline_html import read_offline_input
from .selenium_manual_maps import collect_companies_from_selenium_live_maps
from .selenium_pool import SeleniumPool
from .utils import bbox_from_center_diameter_km, now_iso_local, now_str_for_filename, safe_str
from .web_enrich import enrich_companies_web
from .yandex_api import fetch_by_uri, search_bbox


def _build_request_meta(st: Settings) -> Dict[str, Any]:
    return {
        "request_time": now_iso_local(),
        "MODE": st.MODE,
        "OFFLINE_HTML_INPUT": st.OFFLINE_HTML_INPUT,
        "OFFLINE_ENRICH_MODE": st.OFFLINE_ENRICH_MODE,
        "TEXT": st.TEXT,
        "LANG": st.LANG,
        "CENTER_LON": st.CENTER_LON,
        "CENTER_LAT": st.CENTER_LAT,
        "DIAMETER_KM": st.DIAMETER_KM,
        "RESULTS_PER_PAGE": st.RESULTS_PER_PAGE,
        "MAX_SKIP": st.MAX_SKIP,
        "SLEEP_SEC": st.SLEEP_SEC,
        "ENABLE_URI_REQUERY": st.ENABLE_URI_REQUERY,
        "ENABLE_WEB_FALLBACK_FOR_RATING": st.ENABLE_WEB_FALLBACK_FOR_RATING,
        "WEB_FORCE_OVERWRITE": st.WEB_FORCE_OVERWRITE,
        "WEB_MAX_ITEMS": st.WEB_MAX_ITEMS,
        "SELENIUM_HEADLESS": st.SELENIUM_HEADLESS,
        "apikey_present": bool(st.YMAPIKEY),
    }


def _apply_uri_requery_if_needed(st: Settings, companies: List[Company]) -> Dict[str, Any]:
    """
    Логика:
    - Если ENABLE_URI_REQUERY=True и есть YMAPIKEY, делаем fetch_by_uri по c.uri
    - Если в ответе пришёл feature, парсим его и заполняем пустые поля
      (только дополнение пустых, без переписывания).
    """
    if not st.ENABLE_URI_REQUERY:
        return {"enabled": False}

    if not st.YMAPIKEY:
        return {"enabled": True, "skipped": "YMAPIKEY empty"}

    changed = 0
    attempted = 0
    errors: List[str] = []

    from .yandex_api import company_from_feature

    for c in companies:
        uri = safe_str(c.uri)
        if not uri:
            continue

        attempted += 1

        try:
            j = fetch_by_uri(st, uri=uri)
            feats = j.get("features") or []
            f0 = feats[0] if isinstance(feats, list) and feats else None
            if not isinstance(f0, dict):
                continue

            c2 = company_from_feature(f0, st)
            if not c2:
                continue

            def fill_if_empty(attr: str) -> None:
                nonlocal changed
                if safe_str(getattr(c, attr)) or not safe_str(getattr(c2, attr)):
                    return
                setattr(c, attr, getattr(c2, attr))  # type: ignore[misc]
                changed += 1

            for attr in (
                "Индекс",
                "Сайт",
                "Телефон_1",
                "Телефон_2",
                "Телефон_3",
                "Email_1",
                "Email_2",
                "Email_3",
                "Режим_работы",
                "Рейтинг",
                "Количество_оценок",     # <-- НОВОЕ
                "Количество_отзывов",
                "Категория_1",
                "Категория_2",
                "Категория_3",
                "Особенности",
                "Факс_1",
                "Факс_2",
                "Факс_3",
                "Категории_прочие",
            ):
                fill_if_empty(attr)

        except Exception as e:
            errors.append(f"{c.ID}: {e}")

    return {"enabled": True, "attempted": attempted, "changed_fields": changed, "errors": errors}


def run(st: Settings) -> RunResult:
    outdir = Path(st.OUT_DIR)
    outdir.mkdir(parents=True, exist_ok=True)

    outname = f"{st.OUT_PREFIX}_{now_str_for_filename()}.xlsx"
    outpath = str(outdir / outname)

    request_meta = _build_request_meta(st)
    companies: List[Company] = []

    # --- Источник данных (MODE) ---
    if st.MODE == "ONLINEAPI":
        bbox = bbox_from_center_diameter_km(st.CENTER_LON, st.CENTER_LAT, st.DIAMETER_KM)
        request_meta["bbox"] = bbox

        companies, api_meta, err = search_bbox(st, bbox=bbox)
        request_meta["api_meta"] = api_meta

        # ВАЖНО: НЕ выходим. Сохраняем то, что успели собрать.
        if err:
            request_meta["error"] = err
            request_meta["partial"] = True

    elif st.MODE == "OFFLINEHTML":
        companies, offline_meta, err = read_offline_input(st.OFFLINE_HTML_INPUT)
        request_meta["offline_meta"] = offline_meta

        if err:
            request_meta["error"] = err
            request_meta["saved"] = outpath
            request_meta["rows"] = 0
            save_to_excel(st, [], outpath, request_meta)
            return RunResult(companies=[], request_meta=request_meta)

    elif st.MODE == "SELENIUM":
        companies, selenium_meta, err = collect_companies_from_selenium_live_maps(st)
        request_meta["selenium_meta"] = selenium_meta

        if err:
            request_meta["error"] = err
            request_meta["saved"] = outpath
            request_meta["rows"] = 0
            save_to_excel(st, [], outpath, request_meta)
            return RunResult(companies=[], request_meta=request_meta)

    else:
        request_meta["error"] = f"Unsupported MODE: {st.MODE}"
        request_meta["saved"] = outpath
        request_meta["rows"] = 0
        save_to_excel(st, [], outpath, request_meta)
        return RunResult(companies=[], request_meta=request_meta)

    # --- Enrich-цепочка ---
    enrich_stats: Dict[str, Any] = {}

    if st.OFFLINE_ENRICH_MODE in ("API", "APIWEB"):
        enrich_stats["uri_requery"] = _apply_uri_requery_if_needed(st, companies)

    if st.OFFLINE_ENRICH_MODE in ("WEB", "APIWEB"):
        pool = SeleniumPool(st, keep_chrome_open=(st.MODE == "SELENIUM" and st.SELENIUM_KEEP_CHROME_OPEN))
        try:
            enrich_stats["web"] = enrich_companies_web(st, companies, pool)
        finally:
            if not (st.MODE == "SELENIUM" and st.SELENIUM_KEEP_CHROME_OPEN):
                pool.close()

    request_meta["enrich_stats"] = enrich_stats

    # --- Сохраняем Excel в любом случае ---
    save_to_excel(st, companies, outpath, request_meta)
    request_meta["saved"] = outpath
    request_meta["rows"] = len(companies)

    return RunResult(companies=companies, request_meta=request_meta)
