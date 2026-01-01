# -*- coding: utf-8 -*-

from __future__ import annotations

import subprocess
import time
from pathlib import Path
from typing import Optional

import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

from .config import Settings
from .utils import log, safe_str


def _debug_port_url(st: Settings, path: str) -> str:
    return f"http://{st.DEBUG_HOST}:{st.DEBUG_PORT}{path}"


def is_debug_chrome_alive(st: Settings, timeout_sec: float = 0.8) -> bool:
    try:
        r = requests.get(_debug_port_url(st, "/json/version"), timeout=timeout_sec)
        return r.status_code == 200
    except Exception:
        return False


def start_debug_chrome(st: Settings) -> subprocess.Popen:
    chrome_path = Path(st.CHROME_EXE)
    if not chrome_path.exists():
        raise RuntimeError(f"chrome.exe not found: {st.CHROME_EXE}")

    Path(st.CHROME_PROFILE_DIR).mkdir(parents=True, exist_ok=True)

    args = [
        str(chrome_path),
        f"--remote-debugging-address={st.DEBUG_HOST}",
        f"--remote-debugging-port={st.DEBUG_PORT}",
        f"--user-data-dir={st.CHROME_PROFILE_DIR}",
        "--no-first-run",
        "--no-default-browser-check",
        "about:blank",
    ]

    log(f"[CHROME] starting: {' '.join(args)}")
    return subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, close_fds=False)


def wait_debug_chrome(st: Settings) -> bool:
    t0 = time.time()
    while time.time() - t0 < st.CHROME_START_TIMEOUT_SEC:
        if is_debug_chrome_alive(st, timeout_sec=0.8):
            return True
        time.sleep(0.2)
    return False


def selenium_is_blocked(driver: webdriver.Chrome) -> bool:
    u = (driver.current_url or "").lower()
    if "showcaptcha" in u:
        return True
    try:
        if driver.find_elements(By.CSS_SELECTOR, "form[action*='showcaptcha']"):
            return True
        if driver.find_elements(By.CSS_SELECTOR, "iframe[src*='captcha'], iframe[src*='showcaptcha']"):
            return True
        body_text = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
        if "подтвердите, что запросы отправляли вы" in body_text:
            return True
    except Exception:
        return False
    return False


class SeleniumPool:
    """
    Selenium-пул:
    - поднимает Chrome с remote debugging (или аттачится к существующему);
    - умеет вручную “переждать” капчу;
    - отдаёт page_source после ожидания блока контактов;
    - в режиме keep_chrome_open=True не закрывает Chrome (для MODE=SELENIUM).
    """

    def __init__(self, st: Settings, *, keep_chrome_open: bool = False):
        self.st = st
        self.keep_chrome_open = keep_chrome_open

        self.proc: Optional[subprocess.Popen] = None
        self.started_by_us: bool = False
        self.driver: Optional[webdriver.Chrome] = None

    def ensure(self) -> None:
        if self.driver:
            return

        already = is_debug_chrome_alive(self.st)
        if already:
            self.started_by_us = False
            log(f"[CHROME] debug port {self.st.DEBUG_PORT} already alive, will attach.")
        else:
            self.proc = start_debug_chrome(self.st)
            self.started_by_us = True
            if not wait_debug_chrome(self.st):
                self.close()
                raise RuntimeError(f"debug chrome did not become ready on port {self.st.DEBUG_PORT}")

        opt = ChromeOptions()
        if self.st.SELENIUM_HEADLESS:
            opt.add_argument("--headless=new")

        opt.add_experimental_option("debuggerAddress", f"{self.st.DEBUG_HOST}:{self.st.DEBUG_PORT}")
        self.driver = webdriver.Chrome(options=opt)

    def _safe_input(self, prompt: str) -> None:
        try:
            input(prompt)
        except EOFError:
            log("[SELENIUM][WARN] input() -> EOFError (нет stdin). Продолжаем без ожидания Enter.")

    def get_page_html(self, url: str) -> str:
        self.ensure()
        assert self.driver is not None
        d = self.driver

        original_handle = None
        try:
            original_handle = d.current_window_handle
        except Exception:
            original_handle = None

        opened_new_tab = False
        if self.st.SELENIUM_OPEN_URL_IN_NEW_TAB:
            try:
                d.execute_script("window.open('about:blank', '_blank');")
                time.sleep(0.2)
                d.switch_to.window(d.window_handles[-1])
                opened_new_tab = True
            except Exception:
                opened_new_tab = False

        d.get(url)
        time.sleep(float(self.st.SELENIUM_PAGE_WAIT_SEC))

        if selenium_is_blocked(d):
            log(f"[SELENIUM] challenge for url={d.current_url}")
            self._safe_input("[SELENIUM] Решите проверку в окне Chrome и нажмите Enter...")
            time.sleep(1.0)
            if selenium_is_blocked(d):
                raise RuntimeError(f"Challenge still present: {d.current_url}")

        try:
            WebDriverWait(d, int(self.st.SELENIUM_WAIT_CONTACTS_SEC)).until(
                lambda x: x.find_elements(
                    By.CSS_SELECTOR,
                    '[itemprop="telephone"], a[itemprop="url"], .orgpage-phones-view',
                )
            )
        except Exception:
            pass

        html = d.page_source or ""

        # Возврат обратно, чтобы не ломать вкладку Я.Карт
        if opened_new_tab and self.st.SELENIUM_RETURN_TO_ORIGINAL_TAB and original_handle:
            try:
                d.close()
            except Exception:
                pass
            try:
                d.switch_to.window(original_handle)
            except Exception:
                pass

        return html

    def close(self) -> None:
        # В SELENIUM режиме Chrome не закрываем (пользователь закрывает сам)
        if self.keep_chrome_open:
            self.driver = None
            self.proc = None
            return

        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

        if self.started_by_us and self.proc:
            try:
                self.proc.terminate()
                self.proc.wait(timeout=5)
            except Exception:
                try:
                    self.proc.kill()
                except Exception:
                    pass
            self.proc = None

        elif (not self.started_by_us) and self.st.CLOSE_EXISTING_DEBUG_CHROME:
            log("[CHROME][WARN] CLOSE_EXISTING_DEBUG_CHROME=True, но безопасно закрыть чужой Chrome нельзя.")
