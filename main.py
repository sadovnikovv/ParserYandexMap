#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path

from settings import Settings
from utils import now_iso_local, now_str_for_filename, log
from offline_html import read_offline_input
from excel_writer import save_to_excel
from chrome_pool import SeleniumPool
from web_enrich import enrich_rows_web


def main():
    st = Settings()

    out_name = f"{st.OUT_PREFIX}_{now_str_for_filename()}.xlsx"
    out_path = str(Path(st.OUT_DIR) / out_name)

    log(f"out={out_name} mode={st.MODE} input={st.OFFLINE_HTML_INPUT} enrich={st.OFFLINE_ENRICH_MODE}")

    request_meta = {
        "request_time": now_iso_local(),
        "mode": st.MODE,
        "OFFLINE_HTML_INPUT": st.OFFLINE_HTML_INPUT,
        "OFFLINE_ENRICH_MODE": st.OFFLINE_ENRICH_MODE,
        "WEB_FORCE_OVERWRITE": st.WEB_FORCE_OVERWRITE,
        "CHROME_EXE": st.CHROME_EXE,
        "DEBUG_HOST": st.DEBUG_HOST,
        "DEBUG_PORT": st.DEBUG_PORT,
        "CHROME_PROFILE_DIR": st.CHROME_PROFILE_DIR,
        "SELENIUM_HEADLESS": st.SELENIUM_HEADLESS,
        "CLOSE_EXISTING_DEBUG_CHROME": st.CLOSE_EXISTING_DEBUG_CHROME,
    }

    if st.MODE != "OFFLINE_HTML":
        request_meta["error"] = f"Unsupported MODE in this build: {st.MODE}"
        save_to_excel([], out_path, request_meta)
        log(f"saved: {out_path} rows=0")
        return

    rows, offline_meta, err = read_offline_input(st.OFFLINE_HTML_INPUT)
    request_meta["offline_meta"] = offline_meta
    if err:
        request_meta["error"] = err
        save_to_excel([], out_path, request_meta)
        log(f"saved: {out_path} rows=0")
        return

    if st.OFFLINE_ENRICH_MODE.upper() == "WEB" and rows:
        pool = SeleniumPool(st)  # ленивый: Chrome поднимется только при реальной нужде
        try:
            rows, web_stats = enrich_rows_web(st, rows, pool)
            request_meta["enrich_stats"] = {"web": web_stats}
        finally:
            pool.close()

    save_to_excel(rows, out_path, request_meta)
    log(f"saved: {out_path} rows={len(rows)}")


if __name__ == "__main__":
    main()
