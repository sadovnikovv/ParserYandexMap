#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Yandex Maps (API Поиска по организациям) -> Excel (1 файл на запуск).

Требования:
- Все параметры задаются в шапке (константы ниже).
- Никакого ввода с консоли.
- Один xlsx-файл, имя включает дату/время.
- Внутри файла: лист "Организации" (вся доступная инфа) + лист "Запрос" (параметры/время).
- Логи в консоль лаконичные: 1-я строка (файл+параметры), далее короткий прогресс.
"""

import json
import math
import time
from datetime import datetime
from typing import Optional, List, Dict, Any

import requests
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment


# ===================== ШАПКА: МЕНЯЕТЕ ТОЛЬКО ЭТО =====================

API_KEY = "77c0977c-5d69-45fb-84c7-44afcce951cb"

TEXT = "Металлопрокат"
LANG = "ru_RU"

# Центр области поиска (Москва)
CENTER_LON = 37.6173   # долгота
CENTER_LAT = 55.7558   # широта

# Диаметр области в км (например 17.2 = радиус 8.6 км)
DIAMETER_KM = 40

# Пагинация/ограничения API
RESULTS_PER_PAGE = 50     # max 50
MAX_SKIP = 1000           # max 1000
SLEEP_SEC = 0.25

# Куда сохранять
OUT_DIR = "."             # например r"C:\Users\user\Desktop"
OUT_PREFIX = "out"        # будет out_YYYY-MM-DD_HH-MM-SS.xlsx

# Лаконичный режим логов
LOG_EVERY_PAGE = True     # True: печатать страницы; False: только старт/итог


# ===================== ВСПОМОГАТЕЛЬНОЕ =====================

def now_str_for_filename() -> str:
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


def now_iso_local() -> str:
    return datetime.now().astimezone().replace(microsecond=0).isoformat()


def bbox_from_center_diameter_km(center_lon: float, center_lat: float, diameter_km: float) -> str:
    """
    Считает bbox вокруг центра по диаметру в км.
    Возвращает: "lon1,lat1~lon2,lat2".

    Важно: bbox у Яндекса задаётся как lon,lat~lon,lat (сначала долгота, затем широта).
    """
    if diameter_km <= 0:
        raise ValueError("DIAMETER_KM должен быть > 0")

    radius_km = diameter_km / 2.0

    km_per_deg_lat = 110.574
    km_per_deg_lon = 111.320 * math.cos(math.radians(center_lat))
    if abs(km_per_deg_lon) < 1e-9:
        km_per_deg_lon = 1e-9

    dlat = radius_km / km_per_deg_lat
    dlon = radius_km / km_per_deg_lon

    lon1 = center_lon - dlon
    lon2 = center_lon + dlon
    lat1 = center_lat - dlat
    lat2 = center_lat + dlat

    lon1 = max(-180.0, min(180.0, lon1))
    lon2 = max(-180.0, min(180.0, lon2))
    lat1 = max(-90.0, min(90.0, lat1))
    lat2 = max(-90.0, min(90.0, lat2))

    return f"{lon1:.6f},{lat1:.6f}~{lon2:.6f},{lat2:.6f}"


def safe_join(items: List[str]) -> str:
    items = [x.strip() for x in items if isinstance(x, str) and x.strip()]
    seen = set()
    out = []
    for x in items:
        if x not in seen:
            out.append(x)
            seen.add(x)
    return "; ".join(out)


# ===================== ПАРСИНГ ОТВЕТА =====================

def extract_company_info(feature: dict) -> Optional[dict]:
    """
    Достаём максимум полезного + raw_json, чтобы потом можно было извлечь любые доп. поля.
    """
    try:
        props = feature.get("properties", {}) or {}
        meta = props.get("CompanyMetaData", {}) or {}

        coords = (feature.get("geometry", {}) or {}).get("coordinates", []) or []
        lon = coords[0] if len(coords) >= 1 else ""
        lat = coords[1] if len(coords) >= 2 else ""

        categories = meta.get("Categories", []) or []
        category_names = safe_join([c.get("name", "") for c in categories if isinstance(c, dict)])
        category_classes = safe_join([c.get("class", "") for c in categories if isinstance(c, dict)])

        phones, faxes, emails, other_contacts = [], [], [], []
        for contact in (meta.get("Phones", []) or []):
            if not isinstance(contact, dict):
                continue
            formatted = (contact.get("formatted", "") or "").strip()
            ctype = (contact.get("type", "") or "").lower().strip()
            if not formatted:
                continue

            if ctype == "email":
                emails.append(formatted)
            elif ctype == "phone":
                phones.append(formatted)
            elif ctype == "fax":
                faxes.append(formatted)
            else:
                other_contacts.append(f"{ctype}:{formatted}" if ctype else formatted)

        hours_text = ""
        hours_av = (meta.get("Hours", {}) or {}).get("Availabilities", []) or []
        if hours_av and isinstance(hours_av[0], dict):
            hours_text = (hours_av[0].get("text", "") or "").strip()

        rating = meta.get("rating", "")
        review_count = meta.get("review_count", "")

        return {
            "ID": meta.get("id", ""),
            "Название": meta.get("name", ""),
            "Адрес": meta.get("address", ""),
            "Описание": meta.get("description", ""),
            "Email": safe_join(emails),
            "Сайт": meta.get("url", ""),
            "Телефоны": safe_join(phones),
            "Факсы": safe_join(faxes),
            "Прочие контакты": safe_join(other_contacts),
            "Рейтинг": rating,
            "Количество отзывов": review_count,
            "Категории": category_names,
            "Классы категорий": category_classes,
            "Режим работы": hours_text,
            "Долгота": lon,
            "Широта": lat,
            "uri": props.get("uri", ""),
            "raw_json": json.dumps(feature, ensure_ascii=False),
        }
    except Exception:
        return None


# ===================== API =====================

def ymaps_search_page(session: requests.Session, bbox: str, skip: int) -> List[dict]:
    url = "https://search-maps.yandex.ru/v1/"
    params = {
        "apikey": API_KEY,
        "text": TEXT,
        "lang": LANG,
        "type": "biz",
        "bbox": bbox,
        "rspn": 1,
        "results": RESULTS_PER_PAGE,
        "skip": skip,
    }

    # Ретраи для временных проблем
    retry_statuses = {429, 500, 502, 503, 504}
    max_retries = 6
    backoff = 1.0
    last_err = None

    for _ in range(max_retries):
        try:
            r = session.get(url, params=params, timeout=25)

            if r.status_code in retry_statuses:
                body = (r.text or "").strip().replace("\n", " ")
                last_err = f"{r.status_code} {r.reason}: {body[:300]}"
                time.sleep(backoff)
                backoff *= 2
                continue

            if r.status_code >= 400:
                body = (r.text or "").strip().replace("\n", " ")
                raise requests.HTTPError(f"{r.status_code} {r.reason}: {body[:2000]}", response=r)

            data = r.json()
            features = data.get("features", []) or []

            rows = []
            for f in features:
                row = extract_company_info(f)
                if row:
                    rows.append(row)
            return rows

        except (requests.Timeout, requests.ConnectionError) as e:
            last_err = str(e)
            time.sleep(backoff)
            backoff *= 2

    raise requests.HTTPError(f"retry_failed: {last_err}")


def fetch_all(bbox: str):
    """
    Возвращает (rows, error_text).
    Если в середине падает сеть/504/429 и т.п. — НЕ теряем уже собранные строки.
    Также: если пришло меньше RESULTS_PER_PAGE — считаем это последней страницей и завершаем без лишнего запроса.
    """
    all_rows: List[dict] = []
    seen_ids = set()
    error_text = ""

    with requests.Session() as session:
        skip = 0
        while skip <= MAX_SKIP:
            try:
                rows = ymaps_search_page(session, bbox=bbox, skip=skip)
            except Exception as e:
                error_text = str(e)
                break

            if not rows:
                break

            for row in rows:
                org_id = row.get("ID", "")
                if org_id:
                    if org_id in seen_ids:
                        continue
                    seen_ids.add(org_id)
                all_rows.append(row)

            if LOG_EVERY_PAGE:
                print(f"skip={skip}: +{len(rows)} (total={len(all_rows)})")

            # КЛЮЧЕВОЕ ИЗМЕНЕНИЕ: неполная страница => это конец выдачи
            if len(rows) < RESULTS_PER_PAGE:
                break

            skip += RESULTS_PER_PAGE
            time.sleep(SLEEP_SEC)

    return all_rows, error_text


# ===================== EXCEL =====================

def write_request_sheet(ws, request_meta: Dict[str, Any]):
    ws.title = "Запрос"
    ws.append(["Параметр", "Значение"])

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for k, v in request_meta.items():
        ws.append([k, str(v)])

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 80
    ws.freeze_panes = "A2"


def write_companies_sheet(ws, companies: List[dict]):
    ws.title = "Организации"

    headers_set = set()
    for c in companies:
        headers_set.update(c.keys())

    preferred = [
        "ID", "Название", "Адрес", "Описание",
        "Email", "Сайт", "Телефоны", "Факсы", "Прочие контакты",
        "Рейтинг", "Количество отзывов",
        "Категории", "Классы категорий",
        "Режим работы",
        "Долгота", "Широта",
        "uri", "raw_json",
    ]
    headers = [h for h in preferred if h in headers_set] + sorted([h for h in headers_set if h not in preferred])

    ws.append(headers)

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for c in companies:
        ws.append([c.get(h, "") for h in headers])

    for col_num, header in enumerate(headers, 1):
        col_letter = get_column_letter(col_num)
        max_len = len(str(header)) + 2
        for row in ws.iter_rows(min_row=1, min_col=col_num, max_col=col_num):
            for cell in row:
                if cell.value is None:
                    continue
                max_len = max(max_len, len(str(cell.value)) + 1)
        ws.column_dimensions[col_letter].width = min(max_len, 60)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def save_to_excel(companies: List[dict], out_path: str, request_meta: Dict[str, Any]):
    wb = openpyxl.Workbook()

    # 1-й лист (active) -> Организации
    ws_org = wb.active
    write_companies_sheet(ws_org, companies)

    # 2-й лист -> Запрос
    ws_req = wb.create_sheet("Запрос")
    write_request_sheet(ws_req, request_meta)

    wb.active = 0
    wb.save(out_path)


# ===================== MAIN =====================

def main():
    request_time = now_iso_local()
    bbox = bbox_from_center_diameter_km(CENTER_LON, CENTER_LAT, DIAMETER_KM)

    out_name = f"{OUT_PREFIX}_{now_str_for_filename()}.xlsx"
    out_path = f"{OUT_DIR.rstrip('/\\\\')}/{out_name}"

    print(f"{out_name} | text='{TEXT}' | center={CENTER_LON},{CENTER_LAT} | diameter_km={DIAMETER_KM} | lang={LANG}")

    request_meta = {
        "request_time": request_time,
        "text": TEXT,
        "lang": LANG,
        "center_lon": CENTER_LON,
        "center_lat": CENTER_LAT,
        "diameter_km": DIAMETER_KM,
        "bbox": bbox,
        "results_per_page": RESULTS_PER_PAGE,
        "max_skip": MAX_SKIP,
        "sleep_sec": SLEEP_SEC,
    }

    companies, err = fetch_all(bbox=bbox)

    if err:
        print(f"ERROR: {err}")
        request_meta["error"] = err

    print(f"done: rows={len(companies)}")

    save_to_excel(companies, out_path, request_meta)
    print("saved")


if __name__ == "__main__":
    main()
