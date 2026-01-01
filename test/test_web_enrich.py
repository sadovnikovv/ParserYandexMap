import pytest

from ymaps_excel_export.models import Company
from ymaps_excel_export.web_enrich import (
    parse_rating_reviews_from_jsonld,
    parse_web_contacts_fast,
    requests_is_blocked,
    enrich_company_from_web,
)


def test_parse_rating_reviews_from_jsonld():
    html = """
    <html><head>
      <script type="application/ld+json">
      {
        "@context": "https://schema.org",
        "@type": "LocalBusiness",
        "aggregateRating": {"ratingValue":"4.9","reviewCount":"12"}
      }
      </script>
    </head><body></body></html>
    """
    rating, count = parse_rating_reviews_from_jsonld(html)
    assert rating == "4,9"
    assert count == "12"


def test_parse_web_contacts_fast_tel_mailto_url():
    html = """
    <html><head>
      <meta name="description" content="Звоните +7 (999) 111-22-33">
    </head>
    <body>
      <a href="tel:+79992223344">Call</a>
      <a href="mailto:test@example.com">Mail</a>
      <a itemprop="url" href="https://site.example">Site</a>
      <span itemprop="telephone">8 (999) 000-00-00</span>
    </body></html>
    """
    d = parse_web_contacts_fast(html)
    assert "test@example.com" in d["emails"]
    assert d["site"] == "https://site.example"
    # нормализация к +7...
    assert any(x.startswith("+7") for x in d["telephones"])


def test_requests_is_blocked_by_url():
    assert requests_is_blocked("https://yandex.ru/showcaptcha?x=1", "<html></html>") is True


def test_enrich_company_from_web_without_selenium(st_base, monkeypatch):
    # Подготовка Company
    c = Company(ID="123", raw_json="{}")

    # HTML без капчи
    html = """
    <html><head>
      <script type="application/ld+json">
      {"aggregateRating":{"ratingValue":"5","reviewCount":"1"}}
      </script>
    </head>
    <body>
      <a href="tel:+79990000000">Call</a>
      <a href="mailto:x@y.ru">Mail</a>
      <a itemprop="url" href="https://example.com">Site</a>
    </body></html>
    """

    # Подменяем http_get_org_page внутри модуля
    import ymaps_excel_export.web_enrich as we

    def fake_get(session, oid, timeout_sec):
        return ("https://yandex.ru/maps/org/123/", html)

    monkeypatch.setattr(we, "http_get_org_page", fake_get)

    class DummyPool:
        def get_page_html(self, url: str) -> str:
            raise AssertionError("Selenium не должен вызываться без капчи")

    pool = DummyPool()

    import requests
    with requests.Session() as session:
        c2, stats = enrich_company_from_web(st_base, c, pool, session)

    assert c2.Телефон_1.startswith("+7")
    assert c2.Email_1 == "x@y.ru"
    assert c2.Сайт == "https://example.com"
    assert c2.Рейтинг in ("5,0", "5,0".replace(".", ","))  # на всякий
    assert c2.Количество_отзывов == "1"
    assert stats["used_selenium"] is False
