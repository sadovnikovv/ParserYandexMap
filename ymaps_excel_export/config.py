# -*- coding: utf-8 -*-
"""
file:/C:/Users/user/PycharmProjects/ParserYandexMap_main/ymaps_excel_export/config.py

Настройки проекта.

ВАЖНОЕ ПРАВИЛО:
- В .env допускаются ТОЛЬКО: YM_API_KEY или YMAPIKEY
- Никаких MODE / таймаутов / лимитов и т.п. в .env хранить нельзя.

Все остальные параметры меняются:
- либо прямо в этом файле (дефолты),
- либо через переменные окружения (но не через .env).
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List

from dotenv import load_dotenv

from .utils import env_bool01, env_float, env_int, env_str

# project root = папка на уровень выше пакета ymaps_excel_export
PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"


def _validate_env_only_api_key(env_path: Path) -> None:
    """
    Жёстко запрещаем хранить любые параметры кроме API key в .env.
    """
    if not env_path.exists() or not env_path.is_file():
        return

    allowed = {"YM_API_KEY", "YMAPIKEY"}
    bad: List[str] = []

    text = env_path.read_text(encoding="utf-8", errors="ignore")
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k = line.split("=", 1)[0].strip()
        if k and k not in allowed:
            bad.append(k)

    if bad:
        bad_keys = ", ".join(sorted(set(bad)))
        raise RuntimeError(
            "Запрещено хранить параметры конфигурации в .env. "
            "Допускаются только YM_API_KEY/YMAPIKEY. "
            f"Найдены лишние ключи: {bad_keys}"
        )


@dataclass(frozen=True)
class Settings:
    """
    Главная конфигурация.

    MODE:
    - ONLINEAPI   : поиск организаций через Yandex Search Maps API (bbox)
    - OFFLINEHTML : парсинг сохранённых html-файлов выдачи
    - SELENIUM    : живой режим — пользователь открывает Я.Карты в Chrome и жмёт Enter

    OFFLINE_ENRICH_MODE:
    - NONE  : не дозаполнять
    - API   : дозаполнение через uri-requery (нужен YMAPIKEY)
    - WEB   : дозаполнение через web-карточку (requests -> selenium fallback)
    - APIWEB: сначала API, потом WEB
    """

    # ---------------------------
    # Режим
    # ---------------------------
    MODE: str = "OFFLINEHTML"

    # ---------------------------
    # OFFLINEHTML / SELENIUM (вход)
    # ---------------------------
    OFFLINE_HTML_INPUT: str = "./html"
    OFFLINE_ENRICH_MODE: str = "WEB"

    # ---------------------------
    # ONLINEAPI (поиск)
    # ---------------------------
    YMAPIKEY: str = ""  # Берётся из .env (YM_API_KEY / YMAPIKEY) или из окружения.
    TEXT: str = "Где поесть"
    LANG: str = "ru_RU"
    CENTER_LON: float = 37.6173
    CENTER_LAT: float = 55.7558
    DIAMETER_KM: float = 40.0
    RESULTS_PER_PAGE: int = 50
    MAX_SKIP: int = 1000
    STRICT_BBOX: bool = True

    # ---------------------------
    # Паузы / лимиты
    # ---------------------------
    SLEEP_SEC: float = 0.5

    # ---------------------------
    # Enrich поведение
    # ---------------------------
    ENABLE_URI_REQUERY: bool = True
    ENABLE_WEB_FALLBACK_FOR_RATING: bool = True

    # ---------------------------
    # Ограничители колонок (Excel)
    # ---------------------------
    MAX_PHONES: int = 3
    MAX_EMAILS: int = 3
    MAX_FAXES: int = 3
    MAX_CATEGORIES_MAIN: int = 3

    # ---------------------------
    # Выход (Excel)
    # ---------------------------
    OUT_DIR: str = "./results"
    OUT_PREFIX: str = "out"

    # ---------------------------
    # WEB enrich
    # ---------------------------
    WEB_FORCE_OVERWRITE: bool = True
    WEB_MAX_ITEMS: int = 0
    WEB_TIMEOUT_SEC: int = 12

    # ---------------------------
    # Selenium (fallback и SELENIUM-режим)
    # ---------------------------
    CHROME_EXE: str = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    DEBUG_HOST: str = "127.0.0.1"
    DEBUG_PORT: int = 9222
    CHROME_PROFILE_DIR: str = str((PROJECT_ROOT / "ChromeProfile9222").resolve())
    CHROME_START_TIMEOUT_SEC: int = 15

    SELENIUM_HEADLESS: bool = False
    SELENIUM_PAGE_WAIT_SEC: float = 1.0
    SELENIUM_WAIT_CONTACTS_SEC: int = 12
    CLOSE_EXISTING_DEBUG_CHROME: bool = False

    # ---------------------------
    # SELENIUM live mode
    # ---------------------------
    SELENIUM_START_URL: str = "https://yandex.ru/maps/"
    SELENIUM_WAIT_FOR_ENTER: bool = True

    SELENIUM_SCROLL_TO_END: bool = True
    SELENIUM_SCROLL_MAX_SEC: float = 180.0
    SELENIUM_SCROLL_STEP_SEC: float = 0.7
    SELENIUM_SCROLL_STABLE_ROUNDS: int = 8

    SELENIUM_LIST_ITEM_CSS: str = '[data-object="search-list-item"][data-id]'
    SELENIUM_END_MARKER_CSS: str = '[class*="add-business-view"]'

    SELENIUM_OPEN_URL_IN_NEW_TAB: bool = True
    SELENIUM_RETURN_TO_ORIGINAL_TAB: bool = True
    SELENIUM_KEEP_CHROME_OPEN: bool = True

    # ---------------------------
    # Логи
    # ---------------------------
    VERBOSE: bool = True

    # ---------------------------
    # Excel headers
    # ---------------------------
    HEADERS: List[str] = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "HEADERS",
            [
                "ID",
                "Название",
                "Адрес",
                "Индекс",
                "Долгота",
                "Широта",
                "Сайт",
                "Телефон 1",
                "Телефон 2",
                "Телефон 3",
                "Email 1",
                "Email 2",
                "Email 3",
                "Режим работы",
                "Рейтинг",
                "Количество оценок",   # <-- НОВАЯ КОЛОНКА
                "Количество отзывов",
                "Категория 1",
                "Категория 2",
                "Категория 3",
                "Особенности",
                "uri",
                "Факс 1",
                "Факс 2",
                "Факс 3",
                "Категории (прочие)",
                "raw_json",
            ],
        )

    @classmethod
    def from_env(cls) -> "Settings":
        """
        Загружает .env (ТОЛЬКО ключ API) и читает переменные окружения.
        """
        _validate_env_only_api_key(ENV_PATH)

        # .env может содержать только ключ, но load_dotenv пусть загрузит его в окружение
        load_dotenv(dotenv_path=ENV_PATH, override=True)

        # Приоритет: YM_API_KEY, затем YMAPIKEY (для совместимости)
        api_key = env_str("YM_API_KEY", env_str("YMAPIKEY", ""))

        return cls(
            MODE=env_str("MODE", cls.MODE),

            OFFLINE_HTML_INPUT=env_str("OFFLINE_HTML_INPUT", cls.OFFLINE_HTML_INPUT),
            OFFLINE_ENRICH_MODE=env_str("OFFLINE_ENRICH_MODE", cls.OFFLINE_ENRICH_MODE),

            YMAPIKEY=api_key,
            TEXT=env_str("TEXT", cls.TEXT),
            LANG=env_str("LANG", cls.LANG),
            CENTER_LON=env_float("CENTER_LON", cls.CENTER_LON),
            CENTER_LAT=env_float("CENTER_LAT", cls.CENTER_LAT),
            DIAMETER_KM=env_float("DIAMETER_KM", cls.DIAMETER_KM),
            RESULTS_PER_PAGE=env_int("RESULTS_PER_PAGE", cls.RESULTS_PER_PAGE),
            MAX_SKIP=env_int("MAX_SKIP", cls.MAX_SKIP),
            STRICT_BBOX=env_bool01("STRICT_BBOX", cls.STRICT_BBOX),

            SLEEP_SEC=env_float("SLEEP_SEC", cls.SLEEP_SEC),

            ENABLE_URI_REQUERY=env_bool01("ENABLE_URI_REQUERY", cls.ENABLE_URI_REQUERY),
            ENABLE_WEB_FALLBACK_FOR_RATING=env_bool01(
                "ENABLE_WEB_FALLBACK_FOR_RATING", cls.ENABLE_WEB_FALLBACK_FOR_RATING
            ),

            MAX_PHONES=env_int("MAX_PHONES", cls.MAX_PHONES),
            MAX_EMAILS=env_int("MAX_EMAILS", cls.MAX_EMAILS),
            MAX_FAXES=env_int("MAX_FAXES", cls.MAX_FAXES),
            MAX_CATEGORIES_MAIN=env_int("MAX_CATEGORIES_MAIN", cls.MAX_CATEGORIES_MAIN),

            OUT_DIR=env_str("OUT_DIR", cls.OUT_DIR),
            OUT_PREFIX=env_str("OUT_PREFIX", cls.OUT_PREFIX),

            WEB_FORCE_OVERWRITE=env_bool01("WEB_FORCE_OVERWRITE", cls.WEB_FORCE_OVERWRITE),
            WEB_MAX_ITEMS=env_int("WEB_MAX_ITEMS", cls.WEB_MAX_ITEMS),
            WEB_TIMEOUT_SEC=env_int("WEB_TIMEOUT_SEC", cls.WEB_TIMEOUT_SEC),

            CHROME_EXE=env_str("CHROME_EXE", cls.CHROME_EXE),
            DEBUG_HOST=env_str("DEBUG_HOST", cls.DEBUG_HOST),
            DEBUG_PORT=env_int("DEBUG_PORT", cls.DEBUG_PORT),
            CHROME_PROFILE_DIR=env_str("CHROME_PROFILE_DIR", cls.CHROME_PROFILE_DIR),
            CHROME_START_TIMEOUT_SEC=env_int("CHROME_START_TIMEOUT_SEC", cls.CHROME_START_TIMEOUT_SEC),

            SELENIUM_HEADLESS=env_bool01("SELENIUM_HEADLESS", cls.SELENIUM_HEADLESS),
            SELENIUM_PAGE_WAIT_SEC=env_float("SELENIUM_PAGE_WAIT_SEC", cls.SELENIUM_PAGE_WAIT_SEC),
            SELENIUM_WAIT_CONTACTS_SEC=env_int("SELENIUM_WAIT_CONTACTS_SEC", cls.SELENIUM_WAIT_CONTACTS_SEC),
            CLOSE_EXISTING_DEBUG_CHROME=env_bool01("CLOSE_EXISTING_DEBUG_CHROME", cls.CLOSE_EXISTING_DEBUG_CHROME),

            SELENIUM_START_URL=env_str("SELENIUM_START_URL", cls.SELENIUM_START_URL),
            SELENIUM_WAIT_FOR_ENTER=env_bool01("SELENIUM_WAIT_FOR_ENTER", cls.SELENIUM_WAIT_FOR_ENTER),

            SELENIUM_SCROLL_TO_END=env_bool01("SELENIUM_SCROLL_TO_END", cls.SELENIUM_SCROLL_TO_END),
            SELENIUM_SCROLL_MAX_SEC=env_float("SELENIUM_SCROLL_MAX_SEC", cls.SELENIUM_SCROLL_MAX_SEC),
            SELENIUM_SCROLL_STEP_SEC=env_float("SELENIUM_SCROLL_STEP_SEC", cls.SELENIUM_SCROLL_STEP_SEC),
            SELENIUM_SCROLL_STABLE_ROUNDS=env_int("SELENIUM_SCROLL_STABLE_ROUNDS", cls.SELENIUM_SCROLL_STABLE_ROUNDS),

            SELENIUM_LIST_ITEM_CSS=env_str("SELENIUM_LIST_ITEM_CSS", cls.SELENIUM_LIST_ITEM_CSS),
            SELENIUM_END_MARKER_CSS=env_str("SELENIUM_END_MARKER_CSS", cls.SELENIUM_END_MARKER_CSS),

            SELENIUM_OPEN_URL_IN_NEW_TAB=env_bool01("SELENIUM_OPEN_URL_IN_NEW_TAB", cls.SELENIUM_OPEN_URL_IN_NEW_TAB),
            SELENIUM_RETURN_TO_ORIGINAL_TAB=env_bool01(
                "SELENIUM_RETURN_TO_ORIGINAL_TAB", cls.SELENIUM_RETURN_TO_ORIGINAL_TAB
            ),
            SELENIUM_KEEP_CHROME_OPEN=env_bool01("SELENIUM_KEEP_CHROME_OPEN", cls.SELENIUM_KEEP_CHROME_OPEN),

            VERBOSE=env_bool01("VERBOSE", cls.VERBOSE),
        )
