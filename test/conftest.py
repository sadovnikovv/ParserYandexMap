import os
from dataclasses import replace

import pytest

from ymaps_excel_export.config import Settings
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))



@pytest.fixture()
def st_base(tmp_path, monkeypatch) -> Settings:
    """
    Базовые настройки для тестов:
    - OUT_DIR -> tmp
    - быстрые таймауты/паузы
    """
    # На всякий случай чистим ключ, чтобы тесты явно задавали YMAPIKEY где нужно
    monkeypatch.delenv("YM_API_KEY", raising=False)
    monkeypatch.delenv("YMAPIKEY", raising=False)

    st = Settings(
        MODE="ONLINEAPI",
        OFFLINE_HTML_INPUT="./html",
        OFFLINE_ENRICH_MODE="WEB",

        YMAPIKEY="",
        TEXT="Металлообработка",
        LANG="ru_RU",

        CENTER_LON=37.6173,
        CENTER_LAT=55.7558,
        DIAMETER_KM=2.0,

        RESULTS_PER_PAGE=50,
        MAX_SKIP=100,
        STRICT_BBOX=True,

        SLEEP_SEC=0.0,

        ENABLE_URI_REQUERY=True,
        ENABLE_WEB_FALLBACK_FOR_RATING=True,

        MAX_PHONES=3,
        MAX_EMAILS=3,
        MAX_FAXES=3,
        MAX_CATEGORIES_MAIN=3,

        OUT_DIR=str(tmp_path),
        OUT_PREFIX="out",

        WEB_FORCE_OVERWRITE=True,
        WEB_MAX_ITEMS=0,
        WEB_TIMEOUT_SEC=2,

        CHROME_EXE=r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        DEBUG_HOST="127.0.0.1",
        DEBUG_PORT=9222,
        CHROME_PROFILE_DIR=str(tmp_path / "ChromeProfile9222"),
        CHROME_START_TIMEOUT_SEC=2,
        SELENIUM_HEADLESS=True,
        SELENIUM_PAGE_WAIT_SEC=0.0,
        SELENIUM_WAIT_CONTACTS_SEC=1,
        CLOSE_EXISTING_DEBUG_CHROME=False,

        VERBOSE=False,
    )
    return st
