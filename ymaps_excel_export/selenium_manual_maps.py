# -*- coding: utf-8 -*-

from __future__ import annotations

import time
from typing import Any, Dict, List, Tuple

from .config import Settings
from .models import Company
from .offline_html import build_companies_from_offline_html
from .selenium_pool import SeleniumPool, selenium_is_blocked
from .utils import log, safe_str


def _is_yandex_maps_url(url: str) -> bool:
    u = safe_str(url).lower()
    return ("yandex" in u) and ("/maps" in u or "maps.yandex" in u)


def _safe_input(prompt: str) -> None:
    try:
        input(prompt)
    except EOFError:
        log("[SELENIUM][WARN] input() -> EOFError (нет stdin). Продолжаем без ожидания Enter.")


def _ensure_maps_tab(driver, url: str) -> Dict[str, Any]:
    meta: Dict[str, Any] = {"found_existing_tab": False, "opened_new_tab": False, "tab_url": ""}

    # 1) найти уже открытую вкладку Я.Карт
    try:
        for h in list(driver.window_handles):
            driver.switch_to.window(h)
            cur = safe_str(driver.current_url)
            if _is_yandex_maps_url(cur):
                meta["found_existing_tab"] = True
                meta["tab_url"] = cur
                return meta
    except Exception:
        pass

    # 2) открыть, если не нашли
    try:
        driver.execute_script("window.open(arguments[0], '_blank');", url)
        time.sleep(0.2)
        driver.switch_to.window(driver.window_handles[-1])
        meta["opened_new_tab"] = True
    except Exception:
        pass

    try:
        driver.get(url)
    except Exception:
        pass

    meta["tab_url"] = safe_str(getattr(driver, "current_url", ""))
    return meta


def _js_bool(driver, script: str) -> bool:
    try:
        return bool(driver.execute_script(script))
    except Exception:
        return False


def _js_int(driver, script: str) -> int:
    try:
        v = driver.execute_script(script)
        return int(v) if v is not None else 0
    except Exception:
        return 0


def _find_scrollable_container_js(list_item_css: str) -> str:
    # Пытаемся найти ближайший скроллящийся контейнер вокруг элемента выдачи.
    return f"""
    (function() {{
      const item = document.querySelector({list_item_css!r});
      if (!item) return document.scrollingElement || document.documentElement;

      let el = item;
      for (let i = 0; i < 16; i++) {{
        if (!el) break;
        const st = window.getComputedStyle(el);
        const oy = (st && st.overflowY) ? st.overflowY : '';
        if (el.scrollHeight > el.clientHeight + 20 && (oy === 'auto' || oy === 'scroll')) {{
          return el;
        }}
        el = el.parentElement;
      }}
      return document.scrollingElement || document.documentElement;
    }})()
    """


def _scroll_side_panel_to_end(driver, st: Settings) -> Dict[str, Any]:
    meta: Dict[str, Any] = {
        "enabled": True,
        "loops": 0,
        "stable_rounds": 0,
        "items_before": 0,
        "items_after": 0,
        "end_marker_found": False,
        "elapsed_sec": 0.0,
    }

    t0 = time.time()
    meta["items_before"] = _js_int(driver, f"return document.querySelectorAll({st.SELENIUM_LIST_ITEM_CSS!r}).length;")
    scroll_el_js = _find_scrollable_container_js(st.SELENIUM_LIST_ITEM_CSS)

    stable = 0
    last_count = meta["items_before"]
    last_height = _js_int(driver, "return document.body ? document.body.scrollHeight : 0;")

    while time.time() - t0 < float(st.SELENIUM_SCROLL_MAX_SEC):
        meta["loops"] += 1

        # “конец списка” (маркер как в offline_html эвристике)
        if _js_bool(driver, f"return !!document.querySelector({st.SELENIUM_END_MARKER_CSS!r});"):
            meta["end_marker_found"] = True
            break

        # скроллим контейнер в самый низ
        driver.execute_script(
            f"""
            const el = {scroll_el_js};
            if (el) el.scrollTop = el.scrollHeight;
            """
        )

        time.sleep(float(st.SELENIUM_SCROLL_STEP_SEC))

        cnt = _js_int(driver, f"return document.querySelectorAll({st.SELENIUM_LIST_ITEM_CSS!r}).length;")
        h = _js_int(driver, "return document.body ? document.body.scrollHeight : 0;")

        if cnt <= last_count and h <= last_height:
            stable += 1
        else:
            stable = 0

        meta["stable_rounds"] = stable
        last_count = max(last_count, cnt)
        last_height = max(last_height, h)

        if stable >= int(st.SELENIUM_SCROLL_STABLE_ROUNDS):
            break

    meta["items_after"] = _js_int(driver, f"return document.querySelectorAll({st.SELENIUM_LIST_ITEM_CSS!r}).length;")
    meta["elapsed_sec"] = round(time.time() - t0, 3)
    meta["end_marker_found"] = bool(
        meta["end_marker_found"]
        or _js_bool(driver, f"return !!document.querySelector({st.SELENIUM_END_MARKER_CSS!r});")
    )
    return meta


def collect_companies_from_selenium_live_maps(st: Settings) -> Tuple[List[Company], Dict[str, Any], str]:
    request_meta: Dict[str, Any] = {"mode": "SELENIUM", "start_url": st.SELENIUM_START_URL}

    pool = SeleniumPool(st, keep_chrome_open=bool(st.SELENIUM_KEEP_CHROME_OPEN))
    pool.ensure()

    assert pool.driver is not None
    driver = pool.driver

    try:
        request_meta["maps_tab"] = _ensure_maps_tab(driver, st.SELENIUM_START_URL)
        time.sleep(float(st.SELENIUM_PAGE_WAIT_SEC))

        if selenium_is_blocked(driver):
            log(f"[SELENIUM] challenge on {safe_str(driver.current_url)}")
            _safe_input("[SELENIUM] Решите проверку в окне Chrome и нажмите Enter...")
            time.sleep(1.0)
            if selenium_is_blocked(driver):
                return [], {**request_meta, "error": "challenge_still_present"}, "challenge_still_present"

        if st.SELENIUM_WAIT_FOR_ENTER:
            _safe_input(
                "[SELENIUM] Откройте Яндекс.Карты, выполните поиск и дождитесь списка организаций слева.\n"
                "Когда будете готовы — нажмите Enter для выгрузки..."
            )

        # Проверить/доскроллить до конца
        if st.SELENIUM_SCROLL_TO_END:
            request_meta["scroll"] = _scroll_side_panel_to_end(driver, st)
        else:
            request_meta["scroll"] = {"enabled": False}

        html = driver.page_source or ""
        companies, parse_meta = build_companies_from_offline_html(html, source_name="SELENIUM_LIVE")
        request_meta["parse"] = parse_meta
        request_meta["rows_initial"] = len(companies)

        # ВАЖНО: Chrome/вкладку не закрываем (по требованию).
        return companies, request_meta, ""

    except Exception as e:
        request_meta["error"] = safe_str(e)
        return [], request_meta, safe_str(e)
