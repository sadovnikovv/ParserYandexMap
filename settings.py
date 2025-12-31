# -*- coding: utf-8 -*-

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    # Основной режим
    MODE: str = "OFFLINE_HTML"             # пока реализован OFFLINE_HTML
    OFFLINE_HTML_INPUT: str = "./html"     # папка с .html или путь к одному .html
    OFFLINE_ENRICH_MODE: str = "WEB"       # NONE | WEB

    # Выходной файл
    OUT_DIR: str = "./results"
    OUT_PREFIX: str = "out"

    # WEB enrich
    WEB_FORCE_OVERWRITE: bool = True
    WEB_MAX_ITEMS: int = 300
    SLEEP_SEC: float = 0.12

    # Selenium / Chrome
    CHROME_EXE: str = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    DEBUG_HOST: str = "127.0.0.1"
    DEBUG_PORT: int = 9222
    CHROME_PROFILE_DIR: str = str((Path(__file__).resolve().parent / "ChromeProfile9222").resolve())
    CHROME_START_TIMEOUT_SEC: int = 15

    SELENIUM_HEADLESS: bool = False
    SELENIUM_PAGE_WAIT_SEC: float = 1.0
    SELENIUM_WAIT_CONTACTS_SEC: int = 12

    # Если Chrome на порту 9222 был запущен заранее, по умолчанию НЕ закрываем его в конце,
    # чтобы случайно не убить вашу сессию.
    CLOSE_EXISTING_DEBUG_CHROME: bool = False

    # Логи
    VERBOSE: bool = True


HEADERS = [
    "ID", "Название", "Адрес", "Индекс", "Долгота", "Широта", "Сайт",
    "Телефон 1", "Телефон 2", "Телефон 3",
    "Email 1", "Email 2", "Email 3",
    "Режим работы", "Рейтинг", "Количество отзывов",
    "Категория 1", "Категория 2", "Категория 3",
    "Особенности", "uri",
    "Факс 1", "Факс 2", "Факс 3",
    "Категории (прочие)", "raw_json",
]
