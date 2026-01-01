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
# (если .env отсутствует — просто ничего не загрузится, дефолты останутся)
load_dotenv(dotenv_path=ENV_PATH, override=False)

# Режим работы:
# - ONLINEAPI   : получение списка организаций через Yandex Search Maps API по bbox
# - OFFLINEHTML : парсинг уже сохранённых html-файлов выдачи Я.Карт (боковая панель)
# - SELENIUM    : ручной режим: пользователь открывает Я.Карты в Chrome, нажимает Enter -> программа
#                проверяет/доскролливает боковую панель и парсит DOM как OFFLINEHTML
Mode = Literal["ONLINEAPI", "OFFLINEHTML", "SELENIUM"]

# Режим дозаполнения (enrich) после получения списка организаций:
# - NONE  : не дозаполнять
# - API   : дозаполнять через uri-requery (нужен YMAPIKEY)
# - WEB   : дозаполнять через web-карточку (requests -> selenium fallback)
# - APIWEB: сначала API, потом WEB (для пустых полей)
EnrichMode = Literal["NONE", "API", "WEB", "APIWEB"]


def _validate_env_only_api_key(env_path: Path) -> None:
    """
    Жёсткое правило проекта:
    - В .env допускаются только YM_API_KEY или YMAPIKEY.
    - Никакие MODE/MAX_SKIP/RESULTS_PER_PAGE/... в .env хранить нельзя.
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
    Главная конфигурация проекта.

    Принцип:
    - Все поля имеют разумные значения по умолчанию (работает даже без .env).
    - При наличии переменных окружения/.env они переопределяют значения.
    - Но по правилам проекта: в .env держим только ключ API (YM_API_KEY/YMAPIKEY).

    Советы по диапазонам:
    - Любые "лимиты" (MAX_SKIP/WEB_MAX_ITEMS) лучше повышать постепенно и смотреть на 429/5xx.
    - При 504/5xx чаще помогает увеличить SLEEP_SEC и WEB_TIMEOUT_SEC и повторить.
    """

    # ---------------------------
    # Режим
    # ---------------------------

    # MODE: ONLINEAPI/OFFLINEHTML/SELENIUM
    MODE: Mode = "SELENIUM"

    # ---------------------------
    # OFFLINEHTML (вход)
    # ---------------------------

    # Путь к папке с *.html или к одному html-файлу.
    # Пример: "./html"
    OFFLINE_HTML_INPUT: str = "./html"

    # Режим дозаполнения результатов (см. EnrichMode выше).
    OFFLINE_ENRICH_MODE: EnrichMode = "WEB"

    # ---------------------------
    # ONLINEAPI (поиск)
    # ---------------------------

    # Ключ Yandex Search Maps API:
    # - рекомендуется хранить в .env как YM_API_KEY=...
    # - или YMAPIKEY=... (для совместимости)
    YMAPIKEY: str = ""

    # TEXT: поисковый запрос (не пустой).
    # Примеры: "Металлообработка", "Стоматология", "Ремонт телефонов"
    TEXT: str = "Где поесть"

    # LANG: обычно "ru_RU".
    LANG: str = "ru_RU"

    # CENTER_LON/CENTER_LAT: центр bbox (долгота/широта).
    # Диапазоны:
    # - lon: [-180..180]
    # - lat: [-90..90]
    CENTER_LON: float = 37.6173
    CENTER_LAT: float = 55.7558

    # DIAMETER_KM: диаметр зоны поиска (км), bbox строится как квадрат по диаметру.
    # Практический диапазон: 1..200 (чем больше — тем больше результатов/времени).
    DIAMETER_KM: float = 40.0

    # RESULTS_PER_PAGE: сколько результатов просить за запрос.
    # Практический диапазон: 1..50 (50 — обычно максимум и оптимально по скорости).
    RESULTS_PER_PAGE: int = 50

    # MAX_SKIP: верхняя граница skip (псевдо-лимит страниц).
    # Практический диапазон: 0..1000 (выше — упирается в ограничения API; нужен split bbox).
    MAX_SKIP: int = 1000

    # STRICT_BBOX:
    # - True  -> rspn=1 (строго в bbox)
    # - False -> rspn=0 (может выходить за bbox)
    STRICT_BBOX: bool = True

    # ---------------------------
    # Общие лимиты/паузы
    # ---------------------------

    # SLEEP_SEC: пауза между страницами API и/или между сетевыми запросами.
    # Практический диапазон: 0.0..2.0
    # При 429/частых 5xx увеличивать (например 0.25..1.0).
    SLEEP_SEC: float = 0.5

    # ---------------------------
    # Enrich: флаги поведения
    # ---------------------------

    # ENABLE_URI_REQUERY:
    # - True: разрешить enrich через API по uri (только если есть YMAPIKEY)
    # - False: полностью отключить uri-requery
    ENABLE_URI_REQUERY: bool = True

    # ENABLE_WEB_FALLBACK_FOR_RATING:
    # - True: если рейтинга/отзывов нет из API/HTML, можно попытаться добрать из web
    # - False: не лазить в web ради рейтинга/отзывов
    ENABLE_WEB_FALLBACK_FOR_RATING: bool = True

    # ---------------------------
    # Ограничители колонок (Excel)
    # ---------------------------

    # MAX_PHONES/MAX_EMAILS/MAX_FAXES:
    # сколько телефонов/почт/факсов раскладывать по колонкам.
    # Практический диапазон: 1..10 (но Excel обычно удобнее 1..5).
    MAX_PHONES: int = 3
    MAX_EMAILS: int = 3
    MAX_FAXES: int = 3

    # MAX_CATEGORIES_MAIN:
    # сколько категорий писать в "Категория 1..N", остальные пойдут в "Категории (прочие)".
    # Практический диапазон: 1..10
    MAX_CATEGORIES_MAIN: int = 3

    # ---------------------------
    # Выход (Excel)
    # ---------------------------

    # OUT_DIR: папка для результата.
    OUT_DIR: str = "./results"

    # OUT_PREFIX: префикс имени файла.
    OUT_PREFIX: str = "out"

    # ---------------------------
    # WEB enrich
    # ---------------------------

    # WEB_FORCE_OVERWRITE:
    # - True: web-enrich может перезаписывать поля (даже если заполнены)
    # - False: web-enrich заполняет только пустые
    WEB_FORCE_OVERWRITE: bool = True

    # WEB_MAX_ITEMS:
    # - 0  : без лимита (обработать все компании)
    # - >0 : ограничить количество обогащаемых карточек
    WEB_MAX_ITEMS: int = 0

    # WEB_TIMEOUT_SEC:
    # таймаут HTTP запросов (requests) и некоторых ожиданий.
    # Практический диапазон: 5..60
    WEB_TIMEOUT_SEC: int = 12

    # ---------------------------
    # Selenium (fallback и SELENIUM-режим)
    # ---------------------------

    # CHROME_EXE: путь к chrome.exe.
    CHROME_EXE: str = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    # DEBUG_HOST/DEBUG_PORT:
    # remote-debugging для подключения Selenium к существующему профилю Chrome.
    DEBUG_HOST: str = "127.0.0.1"
    DEBUG_PORT: int = 9222

    # CHROME_PROFILE_DIR:
    # user-data-dir для “отдельного” профиля debug-Chrome.
    CHROME_PROFILE_DIR: str = str((PROJECT_ROOT / "ChromeProfile9222").resolve())

    # CHROME_START_TIMEOUT_SEC:
    # сколько ждать, пока debug-Chrome поднимет порт.
    CHROME_START_TIMEOUT_SEC: int = 15

    # SELENIUM_HEADLESS:
    # - True  -> headless режим
    # - False -> обычное окно
    SELENIUM_HEADLESS: bool = False

    # SELENIUM_PAGE_WAIT_SEC:
    SELENIUM_PAGE_WAIT_SEC: float = 1.0

    # SELENIUM_WAIT_CONTACTS_SEC:
    SELENIUM_WAIT_CONTACTS_SEC: int = 12

    # CLOSE_EXISTING_DEBUG_CHROME:
    CLOSE_EXISTING_DEBUG_CHROME: bool = False

    # ---------------------------
    # SELENIUM live mode (ручной режим)
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
    HEADERS: List[str] = None  # будет заполнено в __post_init__

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "HEADERS",
            [
                "ID", "Название", "Адрес", "Индекс", "Долгота", "Широта", "Сайт",
                "Телефон 1", "Телефон 2", "Телефон 3",
                "Email 1", "Email 2", "Email 3",
                "Режим работы", "Рейтинг", "Количество отзывов",
                "Категория 1", "Категория 2", "Категория 3",
                "Особенности", "uri",
                "Факс 1", "Факс 2", "Факс 3",
                "Категории (прочие)", "raw_json",
            ],
        )

    @classmethod
    def from_env(cls) -> "Settings":
        """
        Загружает .env и читает переменные окружения.
        Поддерживает:
        - YM_API_KEY (предпочтительно)
        - YMAPIKEY (совместимость)

        Важно: если переменные не заданы, берутся дефолты из Settings.
        """
        # Запрещаем любые параметры в .env кроме ключа:
        _validate_env_only_api_key(ENV_PATH)

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
            CHROME_START_TIMEOUT_SEC=env_int(
                "CHROME_START_TIMEOUT_SEC", cls.CHROME_START_TIMEOUT_SEC
            ),
            SELENIUM_HEADLESS=env_bool01("SELENIUM_HEADLESS", cls.SELENIUM_HEADLESS),
            SELENIUM_PAGE_WAIT_SEC=env_float(
                "SELENIUM_PAGE_WAIT_SEC", cls.SELENIUM_PAGE_WAIT_SEC
            ),
            SELENIUM_WAIT_CONTACTS_SEC=env_int(
                "SELENIUM_WAIT_CONTACTS_SEC", cls.SELENIUM_WAIT_CONTACTS_SEC
            ),
            CLOSE_EXISTING_DEBUG_CHROME=env_bool01(
                "CLOSE_EXISTING_DEBUG_CHROME", cls.CLOSE_EXISTING_DEBUG_CHROME
            ),
            SELENIUM_START_URL=env_str("SELENIUM_START_URL", cls.SELENIUM_START_URL),
            SELENIUM_WAIT_FOR_ENTER=env_bool01(
                "SELENIUM_WAIT_FOR_ENTER", cls.SELENIUM_WAIT_FOR_ENTER
            ),
            SELENIUM_SCROLL_TO_END=env_bool01(
                "SELENIUM_SCROLL_TO_END", cls.SELENIUM_SCROLL_TO_END
            ),
            SELENIUM_SCROLL_MAX_SEC=env_float(
                "SELENIUM_SCROLL_MAX_SEC", cls.SELENIUM_SCROLL_MAX_SEC
            ),
            SELENIUM_SCROLL_STEP_SEC=env_float(
                "SELENIUM_SCROLL_STEP_SEC", cls.SELENIUM_SCROLL_STEP_SEC
            ),
            SELENIUM_SCROLL_STABLE_ROUNDS=env_int(
                "SELENIUM_SCROLL_STABLE_ROUNDS", cls.SELENIUM_SCROLL_STABLE_ROUNDS
            ),
            SELENIUM_LIST_ITEM_CSS=env_str(
                "SELENIUM_LIST_ITEM_CSS", cls.SELENIUM_LIST_ITEM_CSS
            ),
            SELENIUM_END_MARKER_CSS=env_str(
                "SELENIUM_END_MARKER_CSS", cls.SELENIUM_END_MARKER_CSS
            ),
            SELENIUM_OPEN_URL_IN_NEW_TAB=env_bool01(
                "SELENIUM_OPEN_URL_IN_NEW_TAB", cls.SELENIUM_OPEN_URL_IN_NEW_TAB
            ),
            SELENIUM_RETURN_TO_ORIGINAL_TAB=env_bool01(
                "SELENIUM_RETURN_TO_ORIGINAL_TAB", cls.SELENIUM_RETURN_TO_ORIGINAL_TAB
            ),
            SELENIUM_KEEP_CHROME_OPEN=env_bool01(
                "SELENIUM_KEEP_CHROME_OPEN", cls.SELENIUM_KEEP_CHROME_OPEN
            ),
            VERBOSE=env_bool01("VERBOSE", cls.VERBOSE),
        )
