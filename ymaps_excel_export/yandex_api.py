# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import time
from typing import Any, Dict, List, Optional, Tuple

import requests

from .config import Settings
from .models import Company
from .utils import dedup_keep_order, json_dumps_safe, log, pick_n, safe_join, safe_str

YMAPS_SEARCH_URL = "https://search-maps.yandex.ru/v1"
_RETRY_STATUSES = {429, 500, 502, 503, 504}


def _format_rating_1(x: Any) -> str:
    """
    Нормализация рейтинга: 1 знак после запятой, как "4,9".
    Убирает float-хвосты типа 4.900000095367432.
    """
    s = safe_str(x).replace(",", ".")
    if not s:
        return ""
    try:
        v = float(s)
    except Exception:
        return ""
    return f"{v:.1f}".replace(".", ",")


def _human_api_hint(status: int, message: str) -> str:
    """
    Подсказки пользователю для типовых кодов ошибок API (400/403/429) и 5xx.
    Коды и примеры формата ошибок описаны в документации. [page:1]
    """
    msg = (message or "").lower()

    if status == 400:
        return (
            "400 Bad Request: отсутствует обязательный параметр или неверное значение. "
            "Проверьте TEXT (не пустой), bbox, results/skip и наличие apikey."
        )
    if status == 403:
        return (
            "403 Forbidden: неверный apikey или нет доступа к API. "
            "Проверьте YM_API_KEY в .env и что ключ имеет доступ к 'Поиск по организациям'."
        )
    if status == 429:
        return (
            "429 Too Many Requests: слишком много запросов за короткое время. "
            "Увеличьте SLEEP_SEC, уменьшите объём/частоту запросов и повторите позже."
        )
    if status >= 500:
        return "5xx: ошибка на стороне сервера. Подождите и повторите позже."

    if "timeout" in msg:
        return "Таймаут: проверьте интернет/прокси/VPN и WEB_TIMEOUT_SEC."
    return "Неизвестная ошибка API. Смотрите status/message и тело ответа."


def _raise_api_http_error(r: requests.Response) -> None:
    """
    Пытаемся распарсить стандартный JSON ошибки:
    {"statusCode":403,"error":"Forbidden","message":"Invalid apikey"} [page:1]
    """
    text = r.text or ""
    status = r.status_code
    error = safe_str(r.reason)
    message = safe_str(text[:300])

    try:
        j = r.json()
        if isinstance(j, dict):
            status = int(j.get("statusCode") or status)
            error = safe_str(j.get("error") or error)
            message = safe_str(j.get("message") or message)
    except Exception:
        pass

    hint = _human_api_hint(status, message)
    raise requests.HTTPError(f"Yandex API error: HTTP {status} {error}: {message}. Hint: {hint}")


def _get_json_with_retries(session: requests.Session, *, params: Dict[str, Any], timeout_sec: int) -> Dict[str, Any]:
    backoff = 1.0
    last_err = None

    for attempt in range(1, 7):
        try:
            r = session.get(YMAPS_SEARCH_URL, params=params, timeout=timeout_sec)

            if r.status_code in _RETRY_STATUSES:
                last_err = f"{r.status_code} {r.reason} {safe_str((r.text or '')[:300])}"
                time.sleep(backoff)
                backoff *= 2
                continue

            if r.status_code >= 400:
                _raise_api_http_error(r)

            return r.json()

        except (requests.Timeout, requests.ConnectionError) as e:
            last_err = str(e)
            time.sleep(backoff)
            backoff *= 2

    raise requests.HTTPError(f"retry_failed: {last_err}")


def _days_ranges_ru(days: List[str]) -> str:
    order = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
    idx = sorted(set(order.index(d) for d in days if d in order))
    if not idx:
        return ""

    ranges: List[Tuple[int, int]] = []
    start = prev = idx[0]
    for i in idx[1:]:
        if i == prev + 1:
            prev = i
            continue
        ranges.append((start, prev))
        start = prev = i
    ranges.append((start, prev))

    parts = []
    for a, b in ranges:
        if a == b:
            parts.append(order[a])
        else:
            parts.append(f"{order[a]}-{order[b]}")
    return ", ".join(parts)


def _normalize_hhmm(t: str) -> str:
    t = safe_str(t)  # В API иногда приходит "09:00:00"
    if len(t) >= 5 and t[2] == ":":
        return t[:5]
    return t


def parse_contacts_meta(meta: Dict[str, Any]) -> Tuple[List[str], List[str], List[str]]:
    phones: List[str] = []
    emails: List[str] = []
    faxes: List[str] = []

    for c in meta.get("Phones") or []:
        if not isinstance(c, dict):
            continue
        formatted = safe_str(c.get("formatted"))
        ctype = safe_str(c.get("type")).lower()
        if not formatted:
            continue
        if ctype == "email":
            emails.append(formatted)
        elif ctype == "fax":
            faxes.append(formatted)
        else:
            phones.append(formatted)

    return dedup_keep_order(phones), dedup_keep_order(emails), dedup_keep_order(faxes)


def parse_categories_meta(meta: Dict[str, Any]) -> List[str]:
    names: List[str] = []
    for c in meta.get("Categories") or []:
        if not isinstance(c, dict):
            continue
        n = safe_str(c.get("name"))
        if n:
            names.append(n)
    return dedup_keep_order(names)


def parse_address_meta(meta: Dict[str, Any], props: Dict[str, Any]) -> Tuple[str, str]:
    address = safe_str(meta.get("address"))
    postal = ""

    addr_obj = meta.get("Address")
    if isinstance(addr_obj, dict):
        postal = safe_str(addr_obj.get("postalCode") or addr_obj.get("postalcode") or addr_obj.get("post"))
        formatted = safe_str(addr_obj.get("formatted"))
        if not address and formatted:
            address = formatted

    if not address:
        address = safe_str(props.get("description"))

    return address, postal


def parse_hours_meta(meta: Dict[str, Any]) -> str:
    hours = meta.get("Hours")
    if not isinstance(hours, dict):
        return ""

    hourstext = safe_str(hours.get("text"))
    if hourstext:
        return hourstext

    av = hours.get("Availabilities")
    if not isinstance(av, list) or not av:
        return ""

    daymap = {
        "Monday": "Пн",
        "Tuesday": "Вт",
        "Wednesday": "Ср",
        "Thursday": "Чт",
        "Friday": "Пт",
        "Saturday": "Сб",
        "Sunday": "Вс",
    }

    parts: List[str] = []
    for a in av:
        if not isinstance(a, dict):
            continue

        if a.get("TwentyFourHours") is True:
            parts.append("24/7")
            continue

        intervals = a.get("Intervals") or []
        segs: List[str] = []
        if isinstance(intervals, list):
            for inter in intervals:
                if not isinstance(inter, dict):
                    continue
                fr = _normalize_hhmm(inter.get("from"))
                to = _normalize_hhmm(inter.get("to"))
                if fr and to:
                    segs.append(f"{fr}-{to}")

        timestr = ", ".join(segs) if segs else ""

        if a.get("Everyday") is True:
            parts.append(timestr if timestr else "Ежедневно")
            continue

        days: List[str] = []
        for k, ru in daymap.items():
            if a.get(k) is True:
                days.append(ru)

        daysstr = _days_ranges_ru(days)
        if daysstr and timestr:
            parts.append(f"{daysstr} {timestr}")
        elif daysstr:
            parts.append(daysstr)
        elif timestr:
            parts.append(timestr)

    return safe_join(parts)


def parse_features_meta(meta: Dict[str, Any]) -> str:
    feats = meta.get("Features")
    if not isinstance(feats, list) or not feats:
        return ""

    out: List[str] = []
    for f in feats:
        if not isinstance(f, dict):
            continue

        name = safe_str(f.get("name") or f.get("id"))
        value = f.get("value")
        valuestr = ""

        if isinstance(value, bool):
            valuestr = "Да" if value else "Нет"
        elif isinstance(value, list):
            valuestr = safe_join([
                safe_str(x.get("name") or x.get("id")) if isinstance(x, dict) else (
                    "Да" if x is True else "Нет" if x is False else safe_str(x)
                )
                for x in value
            ])
        elif isinstance(value, dict):
            valuestr = safe_str(value.get("name") or value.get("id"))
        else:
            valuestr = safe_str(value)

        if name and valuestr:
            out.append(f"{name}: {valuestr}")
        elif name:
            out.append(name)
        elif valuestr:
            out.append(valuestr)

    return safe_join(out)


def company_from_feature(feature: Dict[str, Any], st: Settings) -> Optional[Company]:
    """
    Переводит один feature из API в Company (Excel-строку).
    """
    try:
        props = feature.get("properties") or {}
        meta = (props.get("CompanyMetaData") or {}) if isinstance(props, dict) else {}

        geom = feature.get("geometry") or {}
        coords = (geom.get("coordinates") or []) if isinstance(geom, dict) else []
        lon = safe_str(coords[0]) if len(coords) >= 1 else ""
        lat = safe_str(coords[1]) if len(coords) >= 2 else ""

        org_id = safe_str(meta.get("id"))
        name = safe_str(meta.get("name") or props.get("name"))

        address, postal = parse_address_meta(meta, props)

        phones, emails, faxes = parse_contacts_meta(meta)
        phones_cols = pick_n(phones, st.MAX_PHONES)
        emails_cols = pick_n(emails, st.MAX_EMAILS)
        faxes_cols = pick_n(faxes, st.MAX_FAXES)

        categories = parse_categories_meta(meta)
        cat_main = categories[:st.MAX_CATEGORIES_MAIN]
        cat_main_cols = pick_n(cat_main, st.MAX_CATEGORIES_MAIN)
        cat_extra = categories[st.MAX_CATEGORIES_MAIN:]
        cat_extra_str = safe_join(cat_extra)

        worktime = parse_hours_meta(meta)
        features_str = parse_features_meta(meta)

        rating = _format_rating_1(meta.get("rating"))
        reviewcount = safe_str(meta.get("reviewCount") or meta.get("reviewcount"))

        uri = safe_str(props.get("uri"))

        return Company(
            ID=org_id,
            Название=name,
            Адрес=address,
            Индекс=postal,
            Долгота=lon,
            Широта=lat,
            Сайт=safe_str(meta.get("url")),

            Телефон_1=phones_cols[0], Телефон_2=phones_cols[1], Телефон_3=phones_cols[2],
            Email_1=emails_cols[0], Email_2=emails_cols[1], Email_3=emails_cols[2],

            Режим_работы=worktime,
            Рейтинг=rating,
            Количество_отзывов=reviewcount,

            Категория_1=cat_main_cols[0], Категория_2=cat_main_cols[1], Категория_3=cat_main_cols[2],
            Особенности=features_str,
            uri=uri,

            Факс_1=faxes_cols[0], Факс_2=faxes_cols[1], Факс_3=faxes_cols[2],
            Категории_прочие=cat_extra_str,
            raw_json=json_dumps_safe(feature),
        )

    except Exception:
        return None


def search_bbox(st: Settings, bbox: str) -> Tuple[List[Company], Dict[str, Any], str]:
    """
    ONLINEAPI: постраничный поиск по bbox.
    """
    if not st.YMAPIKEY:
        return [], {}, "YMAPIKEY is empty"

    out: List[Company] = []
    seen = set()
    err = ""

    params_base = {
        "apikey": st.YMAPIKEY,
        "text": st.TEXT,
        "lang": st.LANG,
        "type": "biz",
        "bbox": bbox,
        "rspn": 1 if st.STRICT_BBOX else 0,
        "results": st.RESULTS_PER_PAGE,
    }

    with requests.Session() as session:
        skip = 0
        while skip <= st.MAX_SKIP:
            params = dict(params_base)
            params["skip"] = skip

            try:
                data = _get_json_with_retries(session, params=params, timeout_sec=st.WEB_TIMEOUT_SEC)
            except Exception as e:
                err = str(e)
                break

            features = data.get("features") or []
            rows: List[Company] = []

            for f in features:
                if not isinstance(f, dict):
                    continue
                c = company_from_feature(f, st)
                if not c or not c.ID:
                    continue
                if c.ID in seen:
                    continue
                seen.add(c.ID)
                rows.append(c)

            out.extend(rows)

            if st.VERBOSE:
                log(f"[API] skip={skip} page_rows={len(rows)} total={len(out)}")

            if len(features) < st.RESULTS_PER_PAGE:
                break

            skip += st.RESULTS_PER_PAGE
            time.sleep(st.SLEEP_SEC)

    meta = {"total": len(out), "unique": len(seen)}
    return out, meta, err


def fetch_by_uri(st: Settings, *, uri: str) -> Dict[str, Any]:
    """
    uri-requery: точечный запрос по uri (для уточнения/дозаполнения).
    """
    if not st.YMAPIKEY:
        raise RuntimeError("YMAPIKEY is empty")

    params = {
        "apikey": st.YMAPIKEY,
        "uri": safe_str(uri),
        "lang": st.LANG,
        "type": "biz",
        "results": 1,
        "skip": 0,
    }

    with requests.Session() as session:
        return _get_json_with_retries(session, params=params, timeout_sec=st.WEB_TIMEOUT_SEC)
