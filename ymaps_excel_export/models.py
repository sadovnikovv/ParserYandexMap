# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List


@dataclass
class Company:
    """
    Единая модель строки, которую потом пишем в Excel.

    Важно:
    - Поля названы так же, как колонки в Excel (русские заголовки).
    - raw_json хранит трассировку: откуда что пришло, и "сырой" кусок данных.
    """
    ID: str = ""
    Название: str = ""
    Адрес: str = ""
    Индекс: str = ""
    Долгота: str = ""
    Широта: str = ""
    Сайт: str = ""

    Телефон_1: str = ""
    Телефон_2: str = ""
    Телефон_3: str = ""

    Email_1: str = ""
    Email_2: str = ""
    Email_3: str = ""

    Режим_работы: str = ""
    Рейтинг: str = ""
    Количество_отзывов: str = ""

    Категория_1: str = ""
    Категория_2: str = ""
    Категория_3: str = ""

    Особенности: str = ""
    uri: str = ""

    Факс_1: str = ""
    Факс_2: str = ""
    Факс_3: str = ""

    Категории_прочие: str = ""
    raw_json: str = ""

    def as_excel_row(self) -> Dict[str, Any]:
        """
        Возвращает dict строго под заголовки HEADERS.
        (Внутренние имена со знаком '_' переводим обратно.)
        """
        return {
            "ID": self.ID,
            "Название": self.Название,
            "Адрес": self.Адрес,
            "Индекс": self.Индекс,
            "Долгота": self.Долгота,
            "Широта": self.Широта,
            "Сайт": self.Сайт,
            "Телефон 1": self.Телефон_1,
            "Телефон 2": self.Телефон_2,
            "Телефон 3": self.Телефон_3,
            "Email 1": self.Email_1,
            "Email 2": self.Email_2,
            "Email 3": self.Email_3,
            "Режим работы": self.Режим_работы,
            "Рейтинг": self.Рейтинг,
            "Количество отзывов": self.Количество_отзывов,
            "Категория 1": self.Категория_1,
            "Категория 2": self.Категория_2,
            "Категория 3": self.Категория_3,
            "Особенности": self.Особенности,
            "uri": self.uri,
            "Факс 1": self.Факс_1,
            "Факс 2": self.Факс_2,
            "Факс 3": self.Факс_3,
            "Категории (прочие)": self.Категории_прочие,
            "raw_json": self.raw_json,
        }


@dataclass
class RunResult:
    companies: List[Company] = field(default_factory=list)
    request_meta: Dict[str, Any] = field(default_factory=dict)
