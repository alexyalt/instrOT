"""Печать бланков инструктажей: DOCX и, при наличии, PDF.

Два вида печати:
- заполненные бланки по планам (с кодом бланка);
- пустые бланки организации с общим кодом 000-000 для ручного заполнения
  вне системы (при отсутствии компьютера/электричества).

Шаблон не обрезает длинные значения: текст переносится в ячейках.
"""

from __future__ import annotations

import os
import shutil
import subprocess
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor

from . import config, deadlines
from .errors import ValidationError
from .excel_store import Book
from .validation import display_number, format_code


@dataclass
class PrintResult:
    docx_path: Path
    pdf_path: Path | None
    cards: int
    note: str = ""


def placeholder_values(book: Book, row: dict[str, Any], *,
                       card_number: int = 1) -> dict[str, str]:
    """Значения для плейсхолдеров DOCX-шаблона.

    Номера отдаются без ведущих нулей: на бланке их удобнее писать от руки.
    Полный вид остаётся в книге.
    """
    def date_text(value: Any) -> str:
        return value.strftime("%d.%m.%Y") if isinstance(value, date) else ""

    code = str(row.get("code") or "")
    return {
        "org_name": book.org_value("full_name") or book.org_value("name"),
        "org_short": book.org_value("name"),
        "org_inn": book.org_value("inn"),
        "org_kpp": book.org_value("kpp"),
        "org_ogrn": book.org_value("ogrn"),
        "org_address": book.org_value("address"),
        "org_director": book.org_value("director"),
        "org_ot": book.org_value("ot_specialist"),
        "fio": str(row.get("fio") or ""),
        "position": str(row.get("position") or ""),
        "department": str(row.get("department") or ""),
        "birth_date": date_text(row.get("birth_date")),
        "personnel_no": str(row.get("personnel_no") or ""),
        "type": str(row.get("type") or ""),
        "reason": str(row.get("reason") or ""),
        "program": str(row.get("program") or ""),
        "plan_date": date_text(row.get("plan_date")),
        "fact_date": date_text(row.get("fact_date")),
        "reg_date": date_text(row.get("reg_date")),
        "code": format_code(code) if code else "",
        "code_raw": code,
        "instructor_fio": str(row.get("instructor_fio") or ""),
        "instructor_position": str(row.get("instructor_position") or ""),
        "number_ot": display_number(str(row.get("number_ot") or "")),
        "number_type": display_number(str(row.get("number_type") or "")),
        "case_id": display_number(str(row.get("case_id") or "")),
        "pack_id": display_number(str(row.get("pack_id") or "")),
        "today": date.today().strftime("%d.%m.%Y"),
        "card_number": str(card_number),
        "blank_line": "______________________",
        "signature_line": "____________________",
        "manual_code": config.MANUAL_CODE_DISPLAY,
    }


def _org(book: Book) -> dict[str, str]:
    return {
        "name": book.org_value("full_name") or book.org_value("name"),
        "inn": book.org_value("inn"),
        "address": book.org_value("address"),
        "director": book.org_value("director"),
        "ot": book.org_value("ot_specialist"),
    }


def _set_cell(cell, text: str, *, bold: bool = False, size: int = 9,
              align=WD_ALIGN_PARAGRAPH.LEFT) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Times New Roman"


def _add_row(table, label: str, value: str, *, size: int = 9) -> None:
    cells = table.add_row().cells
    _set_cell(cells[0], label, bold=False, size=size)
    _set_cell(cells[1], value, bold=True, size=size)
    cells[0].width = Cm(4.6)
    cells[1].width = Cm(11.4)


def _header(doc: Document, org: dict[str, str], title: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(org["name"] or "Организация")
    run.bold = True
    run.font.size = Pt(11)

    if org["inn"] or org["address"]:
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(0)
        parts = []
        if org["inn"]:
            parts.append(f"ИНН {org['inn']}")
        if org["address"]:
            parts.append(org["address"])
        r2 = p2.add_run(" • ".join(parts))
        r2.font.size = Pt(8)
        r2.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(6)
    run3 = p3.add_run(title)
    run3.bold = True
    run3.font.size = Pt(12)


def _footer_notes(doc: Document, *, manual: bool) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(0)
    if manual:
        text = (
            "Бланк заполняется вручную вне системы. Код 000-000 одинаков для всех "
            "таких бланков и означает ручной режим, а не номер документа. "
            "После заполнения и подписания внесите данные в приложение: "
            "раздел «Внести бумажный бланк»."
        )
    else:
        text = (
            "Код бланка служит для поиска записи в приложении. После инструктажа "
            "введите код, фактическую дату и подтвердите наличие оригинала — "
            "система присвоит регистрационные номера."
        )
    run = p.add_run(text)
    run.font.size = Pt(7.5)
    run.italic = True
    run.font.color.rgb = RGBColor(0x70, 0x70, 0x70)


def _card_filled(doc: Document, org: dict[str, str], row: dict[str, Any]) -> None:
    title = f"Инструктаж по охране труда: {row.get('type') or ''}".strip()
    _header(doc, org, title)

    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    plan_date = row.get("plan_date")
    plan_text = plan_date.strftime("%d.%m.%Y") if isinstance(plan_date, date) else ""
    _add_row(table, "Работник (ФИО)", str(row.get("fio") or ""))
    _add_row(table, "Должность", str(row.get("position") or ""))
    _add_row(table, "Подразделение", str(row.get("department") or ""))
    birth = row.get("birth_date")
    _add_row(table, "Дата рождения", birth.strftime("%d.%m.%Y") if isinstance(birth, date) else "")
    _add_row(table, "Вид инструктажа", str(row.get("type") or ""))
    if row.get("program"):
        _add_row(table, "ЛНА / программа", str(row.get("program") or ""))
    if row.get("reason"):
        _add_row(table, "Причина", str(row.get("reason") or ""))
    _add_row(table, "Плановая дата", plan_text)
    _add_row(table, "Фактическая дата", "____ . ____ . 20____     (заполняется при проведении)")
    _add_row(table, "Инструктирующий", str(row.get("instructor_fio") or ""))
    _add_row(table, "Код бланка", format_code(str(row.get("code") or "")))
    _add_row(table, "Номера (после регистрации)", "ОТ ____________  /  ____________  /  дело ________  пакет ________")

    doc.add_paragraph()
    sign = doc.add_table(rows=2, cols=2)
    sign.style = "Table Grid"
    _set_cell(sign.cell(0, 0), "Инструктаж провёл", bold=False)
    _set_cell(sign.cell(0, 1), "Подпись: ______________________ / ______________________")
    _set_cell(sign.cell(1, 0), "Инструктаж получил", bold=False)
    _set_cell(sign.cell(1, 1), "Подпись: ______________________ / ______________________")

    _footer_notes(doc, manual=False)


def _card_blank(doc: Document, org: dict[str, str], training_type: str) -> None:
    _header(doc, org, f"Инструктаж по охране труда: {training_type}")

    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _add_row(table, "Работник (ФИО)", "")
    _add_row(table, "Должность", "")
    _add_row(table, "Подразделение", "")
    _add_row(table, "Дата рождения", "____ . ____ . ________")
    _add_row(table, "Вид инструктажа", training_type)
    _add_row(table, "ЛНА / программа", "")
    _add_row(table, "Причина", "")
    _add_row(table, "Дата проведения", "____ . ____ . 20____")
    _add_row(table, "Инструктирующий", "")
    _add_row(table, "Код бланка", config.MANUAL_CODE_DISPLAY)
    _add_row(table, "Отметка / номер листа", "")
    _add_row(table, "Номера (после регистрации)", "ОТ ____________  /  ____________  /  дело ________  пакет ________")

    doc.add_paragraph()
    sign = doc.add_table(rows=2, cols=2)
    sign.style = "Table Grid"
    _set_cell(sign.cell(0, 0), "Инструктаж провёл", bold=False)
    _set_cell(sign.cell(0, 1), "Подпись: ______________________ / ______________________")
    _set_cell(sign.cell(1, 0), "Инструктаж получил", bold=False)
    _set_cell(sign.cell(1, 1), "Подпись: ______________________ / ______________________")

    _footer_notes(doc, manual=True)


def _new_doc(*, compact: bool = False) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    if compact:
        # Для нескольких карточек на листе: узкие поля, больше места сверху.
        section.top_margin = Cm(0.8)
        section.bottom_margin = Cm(0.8)
        section.left_margin = Cm(0.8)
        section.right_margin = Cm(0.8)
    else:
        section.top_margin = Cm(1.2)
        section.bottom_margin = Cm(1.0)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.2)
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(9)
    return doc


def _output_dir(book: Book) -> Path:
    path = book.data_dir / "forms"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _convert_pdf(docx_path: Path) -> Path | None:
    """Конвертировать DOCX в PDF доступным средством. None, если не удалось."""
    candidate = docx_path.with_suffix(".pdf")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir",
                 str(docx_path.parent), str(docx_path)],
                check=True, capture_output=True, timeout=180,
            )
            if candidate.exists():
                return candidate
        except (subprocess.SubprocessError, OSError):
            pass
    try:  # Windows: конвертер через Word
        from docx2pdf import convert  # type: ignore

        convert(str(docx_path), str(candidate))
        if candidate.exists():
            return candidate
    except Exception:
        pass
    return None


def _finish(doc: Document, book: Book, *, prefix: str, cards: int) -> PrintResult:
    out_dir = _output_dir(book)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    docx_path = out_dir / f"{prefix}_{stamp}_{cards}.docx"
    doc.save(docx_path)
    pdf_path = _convert_pdf(docx_path)
    note = "" if pdf_path else (
        "PDF не создан: не найден LibreOffice или Word. Файл DOCX готов к печати."
    )
    return PrintResult(docx_path=docx_path, pdf_path=pdf_path, cards=cards, note=note)


def print_filled(book: Book, training_ids: Iterable[str]) -> PrintResult:
    """Печать заполненных карточек по выбранным записям."""
    ids = [i for i in training_ids if i]
    if not ids:
        raise ValidationError("Выберите хотя бы одну запись для печати.")
    rows = [
        row for row in book.table("Инструктажи").rows
        if str(row.get("id") or "") in set(ids)
    ]
    if not rows:
        raise ValidationError("Выбранные записи не найдены в книге.")

    org = _org(book)
    doc = _new_doc()
    for index, row in enumerate(rows):
        if index:
            doc.add_page_break()
        _card_filled(doc, org, row)

    result = _finish(doc, book, prefix="filled", cards=len(rows))
    from .journal import log_event

    log_event(book.logs_dir(), organization=book.name, action="print_filled",
              entity="Инструктажи", details={"ids": ids, "cards": len(rows)})
    return result


def print_blank(book: Book, *, training_type: str, count: int = 1) -> PrintResult:
    """Печать пустых бланков для ручного заполнения вне системы."""
    if training_type not in config.TRAINING_TYPES:
        raise ValidationError("Выберите вид инструктажа для пустых бланков.")
    count = max(1, min(int(count or 1), 50))

    org = _org(book)
    doc = _new_doc()
    for index in range(count):
        if index:
            doc.add_page_break()
        _card_blank(doc, org, training_type)

    safe_type = "".join(ch for ch in training_type if ch.isalnum()) or "blank"
    result = _finish(doc, book, prefix=f"blank_{safe_type}", cards=count)
    from .journal import log_event

    log_event(book.logs_dir(), organization=book.name, action="print_blank",
              entity="Печать", details={"type": training_type, "count": count})
    return result


def pending_numbers_report(book: Book, *, since: date | None = None) -> list[dict[str, Any]]:
    """Записи, у которых есть номера, но нет отметки о внесении на оригинал."""
    result = []
    for row in book.table("Инструктажи").rows:
        if row.get("status") != "Проведён":
            continue
        fact = row.get("fact_date")
        if since and isinstance(fact, date) and fact < since:
            continue
        result.append(row)
    return sorted(result, key=lambda r: r.get("fact_date") or date.min, reverse=True)


# --------------------------------------------------------------------------- #
# Компактная печать: несколько карточек на одном листе A4
# --------------------------------------------------------------------------- #

def _merged_row(table, text: str, *, bold: bool = True, size: int = 8,
                align=WD_ALIGN_PARAGRAPH.CENTER) -> None:
    """Строка таблицы, объединённая по всей ширине."""
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[1])
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)


def _card_block(table, org: dict[str, str], *, title: str,
                rows: list[tuple[str, str]], manual: bool,
                condensed: bool = False) -> None:
    """Одна карточка как блок строк в общей таблице листа.

    Плоская таблица вместо вложенных: так документ корректно открывается
    в Word и LibreOffice и предсказуемо печатается.
    """
    _merged_row(table, org["name"] or "Организация", bold=True, size=8)
    _merged_row(table, title, bold=True, size=8.5)

    for label, value in rows:
        cell_row = table.add_row().cells
        _set_cell(cell_row[0], label, size=7)
        _set_cell(cell_row[1], value, bold=True, size=7)
        cell_row[0].width = Cm(3.6)
        cell_row[1].width = Cm(5.4)

    for who in ("Провёл", "Получил"):
        cell_row = table.add_row().cells
        _set_cell(cell_row[0], who, size=7)
        _set_cell(cell_row[1], "подпись ____________________", size=7)

    note_text = (
        "Код 000-000 — ручное заполнение вне системы, код не является номером"
        if manual else
        "Код бланка служит для поиска записи; номера вписываются после регистрации"
    )
    if condensed:
        return
    row = table.add_row()
    cell = row.cells[0].merge(row.cells[1])
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(note_text)
    run.italic = True
    run.font.size = Pt(6)
    run.font.color.rgb = RGBColor(0x70, 0x70, 0x70)


def _filled_rows(row: dict[str, Any]) -> list[tuple[str, str]]:
    plan_date = row.get("plan_date")
    birth = row.get("birth_date")
    return [
        ("ФИО", str(row.get("fio") or "")),
        ("Должность", str(row.get("position") or "")),
        ("Подразделение", str(row.get("department") or "")),
        ("Дата рождения", birth.strftime("%d.%m.%Y") if isinstance(birth, date) else ""),
        ("Вид", str(row.get("type") or "")),
        ("ЛНА / программа", str(row.get("program") or "")),
        ("Дата проведения", "____ . ____ . 20____"),
        ("Инструктирующий", str(row.get("instructor_fio") or "")),
        ("Код бланка", format_code(str(row.get("code") or ""))),
        ("Номера", "ОТ ______ / ______ / Д _____ П _____"),
    ]


def _filled_rows_condensed(row: dict[str, Any]) -> list[tuple[str, str]]:
    """Сокращённый набор полей для 4–5 карточек на лист.

    Объединяет должность и подразделение, убирает примечание. Обязательные
    реквизиты по пп. 86–87 остаются на месте.
    """
    birth = row.get("birth_date")
    position = str(row.get("position") or "")
    department = str(row.get("department") or "")
    combined = " / ".join(part for part in (position, department) if part)
    return [
        ("ФИО", str(row.get("fio") or "")),
        ("Должность / подразд.", combined),
        ("Дата рождения", birth.strftime("%d.%m.%Y") if isinstance(birth, date) else ""),
        ("Вид", str(row.get("type") or "")),
        ("ЛНА / программа", str(row.get("program") or "")),
        ("Дата проведения", "____ . ____ . 20____"),
        ("Инструктирующий", str(row.get("instructor_fio") or "")),
        ("Код бланка", format_code(str(row.get("code") or ""))),
        ("Номера", "ОТ ______ / ______ / Д _____ П _____"),
    ]


def _blank_rows_condensed(training_type: str) -> list[tuple[str, str]]:
    return [
        ("ФИО", ""),
        ("Должность / подразд.", ""),
        ("Дата рождения", "____ . ____ . ________"),
        ("Вид", training_type),
        ("ЛНА / программа", ""),
        ("Дата проведения", "____ . ____ . 20____"),
        ("Инструктирующий", ""),
        ("Код бланка", config.MANUAL_CODE_DISPLAY),
        ("Номера", "ОТ ______ / ______ / Д _____ П _____"),
    ]


def _blank_rows(training_type: str) -> list[tuple[str, str]]:
    return [
        ("ФИО", ""),
        ("Должность", ""),
        ("Подразделение", ""),
        ("Дата рождения", "____ . ____ . ________"),
        ("Вид", training_type),
        ("ЛНА / программа", ""),
        ("Дата проведения", "____ . ____ . 20____"),
        ("Инструктирующий", ""),
        ("Код бланка", config.MANUAL_CODE_DISPLAY),
        ("Номера", "ОТ ______ / ______ / Д _____ П _____"),
    ]


def _render_grid(doc: Document, cards: list[tuple[str, list[tuple[str, str]], bool]],
                 org: dict[str, str], *, per_page: int) -> None:
    """Разместить карточки: `per_page` штук на страницу, каждая — блок строк."""
    per_page = max(1, min(per_page, config.MAX_CARDS_PER_PAGE))
    index = 0
    while index < len(cards):
        chunk = cards[index:index + per_page]
        table = doc.add_table(rows=0, cols=2)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        condensed = per_page >= 4
        for position, (title, rows, manual) in enumerate(chunk):
            if position:
                # Визуальный разделитель между карточками на одном листе.
                spacer = table.add_row()
                spacer.cells[0].merge(spacer.cells[1])
            _card_block(table, org, title=title, rows=rows, manual=manual,
                        condensed=condensed)
        index += per_page
        if index < len(cards):
            doc.add_page_break()


def print_filled_grid(book: Book, training_ids: Iterable[str], *,
                      per_page: int = 2) -> PrintResult:
    """Печать заполненных карточек по несколько на лист A4."""
    ids = [i for i in training_ids if i]
    if not ids:
        raise ValidationError("Выберите хотя бы одну запись для печати.")
    rows = [r for r in book.table("Инструктажи").rows if str(r.get("id") or "") in set(ids)]
    if not rows:
        raise ValidationError("Выбранные записи не найдены в книге.")

    org = _org(book)
    doc = _new_doc(compact=per_page > 1)
    builder = _filled_rows_condensed if per_page >= 4 else _filled_rows
    cards = [
        (f"Инструктаж по охране труда: {row.get('type') or ''}", builder(row), False)
        for row in rows
    ]
    _render_grid(doc, cards, org, per_page=per_page)
    result = _finish(doc, book, prefix="filled_grid", cards=len(rows))
    from .journal import log_event

    log_event(book.logs_dir(), organization=book.name, action="print_filled",
              entity="Инструктажи",
              details={"ids": ids, "cards": len(rows), "per_page": per_page})
    return result


def print_blank_grid(book: Book, *, training_type: str, count: int = 1,
                     per_page: int = 2) -> PrintResult:
    """Печать пустых бланков по несколько на лист A4."""
    if training_type not in config.TRAINING_TYPES:
        raise ValidationError("Выберите вид инструктажа для пустых бланков.")
    count = max(1, min(int(count or 1), 50))

    org = _org(book)
    doc = _new_doc(compact=per_page > 1)
    builder2 = _blank_rows_condensed if per_page >= 4 else _blank_rows
    cards = [(f"Инструктаж: {training_type}", builder2(training_type), True)] * count
    _render_grid(doc, cards, org, per_page=per_page)
    safe_type = "".join(ch for ch in training_type if ch.isalnum()) or "blank"
    result = _finish(doc, book, prefix=f"blank_grid_{safe_type}", cards=count)
    from .journal import log_event

    log_event(book.logs_dir(), organization=book.name, action="print_blank",
              entity="Печать",
              details={"type": training_type, "count": count, "per_page": per_page})
    return result
