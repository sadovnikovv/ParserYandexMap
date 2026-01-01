# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Literal

from dotenv import load_dotenv

from .utils import env_bool01, env_float, env_int, env_str

# project root = папка на уровень выше пакета ymaps_excel_export
PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"

# ВАЖНО: не перетираем переменные окружения по умолчанию
load_dotenv(dotenv_path=ENV_PATH, override=False)

Mode = Literal["ONLINEAPI", "OFFLINEHTML"]
EnrichMode = Literal["NONE", "API", "WEB", "APIWEB"]


@dataclass(frozen=True)
class Settings:
    """
    Главная конфигурация.

    Как править:
    - Самый удобный путь: создать .env по образцу .env.example и менять там.
    - Параметры и допустимые значения описаны прямо здесь.
    """

    # --- Режим ---
    MODE: Mode = "OFFLINEHTML"

    # --- OFFLINEHTML ---
    OFFLINE_HTML_INPUT: str = "./html"
    # NONE - не дозаполнять
    # API - дозаполнять через uri-requery (нужен YMAPIKEY)
    # WEB - дозаполнять через web-карточку (requests -> selenium fallback)
    # APIWEB - сначала API, потом WEB (для пустых полей)
    OFFLINE_ENRICH_MODE: EnrichMode = "WEB"

    # --- ONLINEAPI ---
    YMAPIKEY: str = ""  # ключ Yandex Search Maps API
    TEXT: str = "Металлообработка"
    LANG: str = "ru_RU"
    CENTER_LON: float = 37.6173
    CENTER_LAT: float = 55.7558
    DIAMETER_KM: float = 10.0
    RESULTS_PER_PAGE: int = 50
    MAX_SKIP: int = 1000
    STRICT_BBOX: bool = True  # rspn=1

    # --- Общие лимиты/паузы ---
    # (ускорено; при 429 можно увеличить)
    SLEEP_SEC: float = 0.05

    # --- Флаги enrich ---
    ENABLE_URI_REQUERY: bool = True
    ENABLE_WEB_FALLBACK_FOR_RATING: bool = True

    # --- Ограничители колонок ---
    MAX_PHONES: int = 3
    MAX_EMAILS: int = 3
    MAX_FAXES: int = 3
    MAX_CATEGORIES_MAIN: int = 3

    # --- Выход ---
    OUT_DIR: str = "./results"
    OUT_PREFIX: str = "out"

    # --- WEB enrich ---
    WEB_FORCE_OVERWRITE: bool = True
    WEB_MAX_ITEMS: int = 0  # 0 = без лимита
    WEB_TIMEOUT_SEC: int = 12

    # --- Selenium (fallback) ---
    CHROME_EXE: str = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    DEBUG_HOST: str = "127.0.0.1"
    DEBUG_PORT: int = 9222
    CHROME_PROFILE_DIR: str = str((PROJECT_ROOT / "ChromeProfile9222").resolve())
    CHROME_START_TIMEOUT_SEC: int = 15
    SELENIUM_HEADLESS: bool = False
    SELENIUM_PAGE_WAIT_SEC: float = 1.0
    SELENIUM_WAIT_CONTACTS_SEC: int = 12
    CLOSE_EXISTING_DEBUG_CHROME: bool = False

    # --- Логи ---
    VERBOSE: bool = True

    HEADERS: List[str] = None  # будет заполнено в __post_init__

    def __post_init__(self) -> None:
        object.__setattr__(self, "HEADERS", [
            "ID", "Название", "Адрес", "Индекс", "Долгота", "Широта", "Сайт",
            "Телефон 1", "Телефон 2", "Телефон 3",
            "Email 1", "Email 2", "Email 3",
            "Режим работы", "Рейтинг", "Количество отзывов",
            "Категория 1", "Категория 2", "Категория 3",
            "Особенности", "uri",
            "Факс 1", "Факс 2", "Факс 3",
            "Категории (прочие)", "raw_json",
        ])

    @classmethod
    def from_env(cls) -> "Settings":
        """
        Загружает .env и читает переменные окружения.

        Поддерживает:
        - YM_API_KEY (предпочтительно)
        - YMAPIKEY (совместимость)
        """
        load_dotenv(dotenv_path=ENV_PATH, override=True)

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
            ENABLE_WEB_FALLBACK_FOR_RATING=env_bool01("ENABLE_WEB_FALLBACK_FOR_RATING", cls.ENABLE_WEB_FALLBACK_FOR_RATING),

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

            VERBOSE=env_bool01("VERBOSE", cls.VERBOSE),
        )
