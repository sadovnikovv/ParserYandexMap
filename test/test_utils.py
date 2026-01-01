import pytest

from ymaps_excel_export.utils import bbox_from_center_diameter_km, oid_from_uri, safe_str


def test_safe_str():
    assert safe_str(None) == ""
    assert safe_str("  x  ") == "x"
    assert safe_str(123) == "123"


def test_oid_from_uri():
    assert oid_from_uri("https://yandex.ru/maps/?oid=123") == "123"
    assert oid_from_uri("https://yandex.ru/maps/?a=1&oid=999&b=2") == "999"
    assert oid_from_uri("nope") == ""


def test_bbox_from_center_diameter_km_format():
    bbox = bbox_from_center_diameter_km(37.0, 55.0, 10.0)
    assert "~" in bbox
    a, b = bbox.split("~")
    lon1, lat1 = a.split(",")
    lon2, lat2 = b.split(",")
    float(lon1); float(lat1); float(lon2); float(lat2)


def test_bbox_from_center_diameter_km_invalid():
    with pytest.raises(ValueError):
        bbox_from_center_diameter_km(37.0, 55.0, 0)
