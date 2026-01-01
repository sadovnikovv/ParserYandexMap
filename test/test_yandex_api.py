import json
import time

import pytest

from ymaps_excel_export.yandex_api import company_from_feature, search_bbox, YMAPS_SEARCH_URL


def test_company_from_feature_rating_is_1_decimal_comma(st_base):
    feature = {
        "type": "Feature",
        "geometry": {"type": "Point", "coordinates": [37.0, 55.0]},
        "properties": {
            "CompanyMetaData": {
                "id": "123",
                "name": "Org",
                "rating": 4.900000095367432,
                "reviewCount": 7,
                "url": "https://example.com",
            },
            "uri": "ymapsbm1://org?oid=123",
            "name": "Org",
        },
    }

    c = company_from_feature(feature, st_base)
    assert c is not None
    assert c.ID == "123"
    assert c.Рейтинг == "4,9"
    assert c.Количество_отзывов == "7"


def test_search_bbox_invalid_apikey_returns_human_error(st_base, requests_mock):
    st = st_base.__class__(**{**st_base.__dict__, "YMAPIKEY": "BAD_KEY"})  # frozen dataclass workaround

    requests_mock.get(
        YMAPS_SEARCH_URL,
        status_code=403,
        json={"statusCode": 403, "error": "Forbidden", "message": "Invalid apikey"},
    )

    companies, meta, err = search_bbox(st, bbox="37.0,55.0~38.0,56.0")
    assert companies == []
    assert "HTTP 403" in err
    assert "Invalid apikey" in err
    assert "Hint:" in err


def test_search_bbox_429_then_success(st_base, requests_mock, monkeypatch):
    st = st_base.__class__(**{**st_base.__dict__, "YMAPIKEY": "OK_KEY"})

    # чтобы тест не ждал backoff
    monkeypatch.setattr(time, "sleep", lambda *_: None)

    requests_mock.get(
        YMAPS_SEARCH_URL,
        [
            {"status_code": 429, "text": "Too Many Requests"},
            {"status_code": 200, "json": {"features": []}},
        ],
    )

    companies, meta, err = search_bbox(st, bbox="37.0,55.0~38.0,56.0")
    assert err == ""
    assert companies == []
    assert meta["total"] == 0
