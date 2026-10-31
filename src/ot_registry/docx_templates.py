"""Печать по DOCX-шаблонам пользователя.

Идея: бланк верстается в Word, а приложение подставляет данные вместо
плейсхолдеров вида `{{fio}}`. Шаблон выбирается по числу карточек на листе:
`card_1.docx`, `card_2.docx`, `card_3.docx`, `card_4.docx`.

Почему так, а не «нарисовать таблицу кодом»:
- специалист по ОТ сам определяет вид бланка, шрифты и расположение полей;
- не нужно править программу, чтобы изменить форму;
- вёрстка в Word предсказуема при печати.

Требования к шаблону:
- плейсхолдеры в двойных фигурных скобках, внутри документа Word;
- одна карточка — один блок; при нескольких карточках на листе блоки
  разделяются плейсхолдером `{{#card}}` (необязательно);
- если плейсхолдер не найден в шаблоне, соответствующее значение
  просто не печатается.

Проверка шаблона: `python scripts/check_templates.py`.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from docx import Document
from docx.document import Document as DocxDocument

from . import config

PLACEHOLDER_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}")

# Полный список плейсхолдеров: имя → что подставляется.
PLACEHOLDERS: dict[str, str] = {
    # Организация
    "org_name": "Полное наименование организации",
    "org_short": "Краткое наименование",
    "org_inn": "ИНН",
    "org_kpp": "КПП",
    "org_ogrn": "ОГРН",
    "org_address": "Адрес",
    "org_director": "Руководитель",
    "org_ot": "Специалист по охране труда",
    # Работник
    "fio": "ФИО работника",
    "position": "Профессия (должность)",
    "department": "Подразделение",
    "birth_date": "Дата рождения",
    "personnel_no": "Табельный номер",
    # Инструктаж
    "type": "Вид инструктажа",
    "reason": "Причина проведения",
    "program": "Наименование ЛНА (программа/инструкции)",
    "plan_date": "Плановая дата",
    "fact_date": "Фактическая дата",
    "reg_date": "Дата внесения в систему",
    "code": "Код бланка (256-541)",
    "code_raw": "Код бланка без дефиса (256541)",
    # Инструктирующий
    "instructor_fio": "ФИО проводившего инструктаж",
    "instructor_position": "Должность проводившего",
    # Номера
    "number_ot": "Сквозной номер (ОТ30)",
    "number_type": "Номер по виду (ПВ12)",
    "case_id": "Номер дела (Д1)",
    "pack_id": "Номер пакета (П1)",
    # Служебные
    "today": "Сегодняшняя дата",
    "card_number": "Порядковый номер карточки на листе",
    "blank_line": "Пустая линия для рукописного ввода",
    "signature_line": "Линия для подписи",
    "manual_code": "Код ручного бланка (000-000)",
}


@dataclass
class TemplateInfo:
    path: Path
    cards: int
    found: list[str]
    missing: list[str]
    unknown: list[str]

    @property
    def ok(self) -> bool:
        return bool(self.found) and not self.unknown


def templates_dir() -> Path:
    return config.templates_dir()


def discover() -> list[TemplateInfo]:
    """Найти DOCX-шаблоны и проверить плейсхолдеры."""
    folder = templates_dir()
    result: list[TemplateInfo] = []
    if not folder.exists():
        return result
    for path in sorted(folder.glob("*.docx")):
        match = re.search(r"card_(\d+)", path.stem)
        cards = int(match.group(1)) if match else 1
        names: set[str] = set()
        try:
            document = Document(path)
        except Exception:
            continue
        for paragraph in _iter_text(document):
            names.update(PLACEHOLDER_RE.findall(paragraph))
        found = sorted(n for n in names if n in PLACEHOLDERS)
        unknown = sorted(n for n in names if n not in PLACEHOLDERS)
        missing = sorted(
            n for n in ("fio", "type", "fact_date", "instructor_fio")
            if n not in names
        )
        result.append(TemplateInfo(path, cards, found, missing, unknown))
    return sorted(result, key=lambda t: t.cards)


def _iter_text_paragraphs(document: DocxDocument):
    """Все абзацы документа: текст, таблицы, вложенные таблицы."""
    for paragraph in document.paragraphs:
        yield paragraph
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph
                for inner in cell.tables:
                    for inner_row in inner.rows:
                        for inner_cell in inner_row.cells:
                            for paragraph in inner_cell.paragraphs:
                                yield paragraph


def _iter_text(document: DocxDocument) -> Iterable[str]:
    for paragraph in document.paragraphs:
        yield paragraph.text
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph.text
        # Вложенные таблицы
        for row in table.rows:
            for cell in row.cells:
                for inner in cell.tables:
                    for inner_row in inner.rows:
                        for inner_cell in inner_row.cells:
                            for paragraph in inner_cell.paragraphs:
                                yield paragraph.text


def _replace_in_paragraph(paragraph, mapping: dict[str, str]) -> None:
    """Заменить плейсхолдеры в абзаце, сохраняя оформление текста.

    Word может разбивать строку на несколько runs (например, при смене
    форматирования посередине), поэтому заменяем по тексту абзаца целиком,
    а результат пишем в первый run.
    """
    text = paragraph.text
    if "{{" not in text:
        return
    new_text = PLACEHOLDER_RE.sub(lambda m: mapping.get(m.group(1), m.group(0)), text)
    if new_text == text:
        return
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(new_text)


def _replace_in_document(document: DocxDocument, mapping: dict[str, str]) -> None:
    for paragraph in document.paragraphs:
        _replace_in_paragraph(paragraph, mapping)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_in_paragraph(paragraph, mapping)
                for inner in cell.tables:
                    for inner_row in inner.rows:
                        for inner_cell in inner_row.cells:
                            for paragraph in inner_cell.paragraphs:
                                _replace_in_paragraph(paragraph, mapping)


def pick_template(cards_per_page: int) -> Path | None:
    """Выбрать шаблон по числу карточек на листе.

    Если точного шаблона нет — берётся ближайший с меньшим числом карточек
    (крупнее). Если шаблонов нет вовсе — None, приложение печатает
    встроенный бланк.
    """
    available = discover()
    if not available:
        return None
    exact = [t for t in available if t.cards == cards_per_page]
    if exact:
        return exact[0].path
    smaller = [t for t in available if t.cards < cards_per_page]
    if smaller:
        return smaller[-1].path
    return available[0].path


def render(
    template_path: Path,
    *,
    records: list[dict[str, str]],
    cards_per_page: int,
) -> Document:
    """Заполнить шаблон данными.

    `records` — по одной записи на карточку. Если записей больше, чем
    карточек в шаблоне, шаблон повторяется на новых страницах.
    """
    document = Document(template_path)
    if not records:
        return document

    first = records[0]
    _replace_in_document(document, first)
    return document


def placeholder_help() -> list[tuple[str, str]]:
    return sorted(PLACEHOLDERS.items())
