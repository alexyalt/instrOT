"""Построение книги Excel: оформление, формулы, проверка данных, защита.

Один модуль отвечает за «шаблон» книги, чтобы новая организация получала
одинаковую структуру без ручной работы и без опечаток в названиях колонок.

Защита листов паролем нужна против случайного изменения структуры, а не как
средство безопасности: это ограничение интерфейса Excel. Приложение пишет
книгу через openpyxl независимо от защиты, а ручной режим снимает защиту
на время редактирования и возвращает её после принятия изменений.
"""

from __future__ import annotations

import secrets
import string
from datetime import date
from typing import Iterable

from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.protection import WorkbookProtection

from . import config
from .schema import (
    INSTRUCTION_SHEET,
    ORG_KEY,
    ORG_SHEET,
    SUMMARY_SHEET,
    Column,
    FormulaColumn,
    Sheet,
    SHEETS,
    column_letter,
    formula_context,
    formula_letter,
)

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
SUBHEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
BODY_FONT = Font(size=10)
NOTE_FONT = Font(size=9, italic=True, color="7F7F7F")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DATE_FORMAT = "DD.MM.YYYY"
FORMULA_FILL = PatternFill("solid", fgColor="F2F7FB")

MAX_FORMULA_ROW = 5000

INSTRUCTION_LINES: tuple[str, ...] = (
    "КАК РАБОТАТЬ С КНИГОЙ",
    "",
    "1. Обычная работа — через приложение instrOT. Не открывайте файл в Excel",
    "   одновременно с работой приложения.",
    "",
    "2. Структура книги защищена паролем: названия колонок и листы изменить нельзя,",
    "   строки добавить нельзя. Это защита от случайной порчи структуры.",
    "   Пароль показывает приложение: раздел «Excel и резерв».",
    "",
    "3. Чтобы править данные вручную: в приложении откройте «Excel и резерв» →",
    "   «Открыть Excel для редактирования». Приложение снимет защиту, сделает",
    "   контрольную копию и заблокирует запись через формы.",
    "   После правок: «Проверить и перечитать» → «Принять изменения».",
    "   Защита вернётся автоматически.",
    "",
    "4. Колонку ID не удаляйте и не меняйте: по ней приложение узнаёт записи.",
    "   Даже если строки отсортированы, связи не теряются.",
    "",
    "5. Даты — в формате ДД.ММ.ГГГГ. Статусы выбирайте из выпадающего списка.",
    "   Ввод постороннего значения Excel не примет.",
    "",
    "6. Зелёные колонки справа («Дней до срока», «Состояние») — формулы.",
    "   Их не нужно заполнять вручную, они пересчитываются при открытии книги.",
    "   Приложение считает сроки самостоятельно и не зависит от этих формул.",
    "",
    "7. На листе «Сводка» — формулы подсчёта по статусам и видам инструктажей.",
    "   Значения появятся после открытия книги в Excel или LibreOffice.",
    "",
    "8. Код бланка 000000 означает бумажный бланк, заполненный вручную вне системы.",
    "   Такой код одинаков у всех ручных бланков и не различает документы.",
    "",
    "9. Выданные номера (ОТ, номер по виду, дело, пакет) не меняйте: они уже",
    "   перенесены на бумажный оригинал.",
    "",
    "10. Регистрация инструктажа означает, что специалист подтвердил наличие",
    "    оригинала с собственноручными подписями. Приложение не проверяет",
    "    подлинность подписей.",
    "",
    "11. Резервные копии создаются автоматически перед каждой записью.",
    "    Они лежат в папке backups рядом с книгой.",
    "",
    "12. Форму бланка, порядок регистрации и периодичность инструктажей",
    "    утверждает работодатель (п. 88 Правил № 2464).",
)

SUMMARY_ROWS: tuple[tuple[str, str, str], ...] = (
    ("Всего запланировано", '=COUNTIF(Инструктажи!$M$2:$M$5000,"Запланирован")',
     "записи со статусом «Запланирован»"),
    ("Просрочено", '=COUNTIFS(Инструктажи!$M$2:$M$5000,"Запланирован",'
                   'Инструктажи!$K$2:$K$5000,"<"&TODAY())',
     "плановая дата раньше сегодняшней"),
    ("На сегодня", '=COUNTIFS(Инструктажи!$M$2:$M$5000,"Запланирован",'
                   'Инструктажи!$K$2:$K$5000,TODAY())',
     "плановая дата совпадает с сегодня"),
    ("Ближайшие 7 дней", '=COUNTIFS(Инструктажи!$M$2:$M$5000,"Запланирован",'
                         'Инструктажи!$K$2:$K$5000,">"&TODAY(),'
                         'Инструктажи!$K$2:$K$5000,"<="&TODAY()+7)',
     "требуют подготовки"),
    ("Ближайшие 14 дней", '=COUNTIFS(Инструктажи!$M$2:$M$5000,"Запланирован",'
                          'Инструктажи!$K$2:$K$5000,">"&TODAY()+7,'
                          'Инструктажи!$K$2:$K$5000,"<="&TODAY()+14)',
     "запас времени"),
    ("Проведено всего", '=COUNTIF(Инструктажи!$M$2:$M$5000,"Проведён")',
     "за всё время"),
    ("Проведено в этом году", '=COUNTIFS(Инструктажи!$M$2:$M$5000,"Проведён",'
                              'Инструктажи!$L$2:$L$5000,">="&DATE(YEAR(TODAY()),1,1))',
     "по дате факта"),
    ("Из них по бумажным бланкам", '=COUNTIFS(Инструктажи!$M$2:$M$5000,"Проведён",'
                                   'Инструктажи!$P$2:$P$5000,"Ручной бланк")',
     "внесены вручную после работы без компьютера"),
    ("Записей без даты плана", '=COUNTIFS(Инструктажи!$M$2:$M$5000,"Запланирован",'
                               'Инструктажи!$K$2:$K$5000,"")',
     "нужно уточнить"),
    ("Работников работает", '=COUNTIF(Работники!$J$2:$J$5000,"Работает")', ""),
    ("Работников уволено", '=COUNTIF(Работники!$J$2:$J$5000,"Уволен")', ""),
    ("", "", ""),
    ("ПО ВИДАМ ИНСТРУКТАЖА", "", "запланировано / проведено"),
)

TYPE_SUMMARY_ROWS: tuple[tuple[str, str], ...] = (
    ("Вводный",
     '=COUNTIFS(Инструктажи!$H$2:$H$5000,"Вводный",Инструктажи!$M$2:$M$5000,"Запланирован")'
     '&" / "&COUNTIFS(Инструктажи!$H$2:$H$5000,"Вводный",Инструктажи!$M$2:$M$5000,"Проведён")'),
    ("Первичный",
     '=COUNTIFS(Инструктажи!$H$2:$H$5000,"Первичный",Инструктажи!$M$2:$M$5000,"Запланирован")'
     '&" / "&COUNTIFS(Инструктажи!$H$2:$H$5000,"Первичный",Инструктажи!$M$2:$M$5000,"Проведён")'),
    ("Повторный",
     '=COUNTIFS(Инструктажи!$H$2:$H$5000,"Повторный",Инструктажи!$M$2:$M$5000,"Запланирован")'
     '&" / "&COUNTIFS(Инструктажи!$H$2:$H$5000,"Повторный",Инструктажи!$M$2:$M$5000,"Проведён")'),
    ("Внеплановый",
     '=COUNTIFS(Инструктажи!$H$2:$H$5000,"Внеплановый",Инструктажи!$M$2:$M$5000,"Запланирован")'
     '&" / "&COUNTIFS(Инструктажи!$H$2:$H$5000,"Внеплановый",Инструктажи!$M$2:$M$5000,"Проведён")'),
    ("Целевой",
     '=COUNTIFS(Инструктажи!$H$2:$H$5000,"Целевой",Инструктажи!$M$2:$M$5000,"Запланирован")'
     '&" / "&COUNTIFS(Инструктажи!$H$2:$H$5000,"Целевой",Инструктажи!$M$2:$M$5000,"Проведён")'),
)

SUMMARY_NOTE = (
    "Значения считаются формулами Excel при открытии книги. "
    "Приложение не изменяет этот лист и показывает сроки независимо от него."
)


def generate_password(length: int = 12) -> str:
    """Пароль защиты структуры книги."""
    alphabet = string.ascii_letters + string.digits
    return "".join(secrets.choice(alphabet) for _ in range(length))


# --------------------------------------------------------------------------- #
# Оформление листов
# --------------------------------------------------------------------------- #

def _style_header(ws: Worksheet, columns: Iterable[Column],
                  formulas: Iterable[FormulaColumn] = ()) -> None:
    idx = 1
    for col in columns:
        cell = ws.cell(row=1, column=idx, value=col.title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER
        cell.protection = Protection(locked=True)
        ws.column_dimensions[get_column_letter(idx)].width = col.width
        idx += 1
    for formula in formulas:
        cell = ws.cell(row=1, column=idx, value=formula.title)
        cell.fill = PatternFill("solid", fgColor="2E7D32")
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER
        cell.protection = Protection(locked=True)
        ws.column_dimensions[get_column_letter(idx)].width = formula.width
        idx += 1
    ws.row_dimensions[1].height = 34
    ws.freeze_panes = "A2"


def _add_validation(ws: Worksheet, sheet: Sheet) -> None:
    """Выпадающие списки для статусов: меньше опечаток при ручной правке."""
    letters = {col.title: column_letter(idx) for idx, col in enumerate(sheet.columns, start=1)}
    for key, values in sheet.validation.items():
        col = next((c for c in sheet.columns if c.key == key), None)
        if col is None:
            continue
        letter = letters.get(col.title)
        if not letter:
            continue
        items = ",".join(values)
        if len(items) > 250:
            continue
        validation = DataValidation(
            type="list",
            formula1=f'"{items}"',
            allow_blank=True,
            showDropDown=False,
            showErrorMessage=True,
            errorTitle="Недопустимое значение",
            error="Выберите значение из списка. Свободный ввод запрещён, "
                  "чтобы не нарушить структуру учёта.",
        )
        validation.add(f"{letter}2:{letter}{MAX_FORMULA_ROW}")
        ws.add_data_validation(validation)


def _add_training_highlighting(ws: Worksheet, sheet: Sheet) -> None:
    """Условное оформление по срокам (просрочено / ближайшие / скоро)."""
    from openpyxl.formatting.rule import FormulaRule

    letters = {col.key: column_letter(idx) for idx, col in enumerate(sheet.columns, start=1)}
    plan = letters.get("plan_date")
    status = letters.get("status")
    if not plan or not status:
        return
    last = get_column_letter(len(sheet.columns) + len(sheet.formulas))
    rng = f"A2:{last}{MAX_FORMULA_ROW}"
    rules = (
        FormulaRule(
            formula=[f'AND(${status}2="Запланирован",${plan}2<>"",${plan}2<TODAY())'],
            fill=PatternFill("solid", fgColor="FFC7CE"),
            font=Font(color="9C0006", size=10),
        ),
        FormulaRule(
            formula=[f'AND(${status}2="Запланирован",${plan}2<>"",${plan}2=TODAY())'],
            fill=PatternFill("solid", fgColor="F8CBAD"),
        ),
        FormulaRule(
            formula=[f'AND(${status}2="Запланирован",${plan}2<>"",${plan}2>TODAY(),'
                     f'${plan}2<=TODAY()+{config.DEFAULT_SOON_DAYS})'],
            fill=PatternFill("solid", fgColor="FFD9B3"),
        ),
        FormulaRule(
            formula=[f'AND(${status}2="Запланирован",${plan}2<>"",'
                     f'${plan}2>TODAY()+{config.DEFAULT_SOON_DAYS},'
                     f'${plan}2<=TODAY()+{config.DEFAULT_UPCOMING_DAYS})'],
            fill=PatternFill("solid", fgColor="FFF2CC"),
        ),
        FormulaRule(
            formula=[f'AND(${status}2="Проведён")'],
            fill=PatternFill("solid", fgColor="E4F4E8"),
        ),
    )
    for rule in rules:
        ws.conditional_formatting.add(rng, rule)


def write_formula_columns(ws: Worksheet, sheet: Sheet, *, last_row: int) -> None:
    """Записать формулы во все строки данных (строки 2..last_row)."""
    if not sheet.formulas:
        return
    context = formula_context(sheet)
    for position, formula in enumerate(sheet.formulas, start=1):
        letter = formula_letter(sheet, position)
        for row in range(2, last_row + 1):
            cell = ws.cell(row=row, column=len(sheet.columns) + position)
            cell.value = formula.template.format(row=row, **context)
            cell.font = BODY_FONT
            cell.fill = FORMULA_FILL
            cell.number_format = formula.number_format
            cell.protection = Protection(locked=True)


def write_instruction_sheet(ws: Worksheet) -> None:
    ws.column_dimensions["A"].width = 112
    for idx, line in enumerate(INSTRUCTION_LINES, start=1):
        cell = ws.cell(row=idx, column=1, value=line)
        cell.font = Font(size=11 if idx == 1 else 10, bold=(idx == 1))
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.protection = Protection(locked=True)


def write_summary_sheet(ws: Worksheet) -> None:
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 58

    for col, title in enumerate(("Показатель", "Значение", "Пояснение"), start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
        cell.border = BORDER
        cell.protection = Protection(locked=True)
    ws.row_dimensions[1].height = 24

    row = 2
    for label, formula, note in SUMMARY_ROWS:
        if not label:
            row += 1
            continue
        is_section = not formula
        cell = ws.cell(row=row, column=1, value=label)
        cell.font = Font(size=10, bold=is_section)
        cell.border = BORDER
        cell.protection = Protection(locked=True)
        if formula:
            value_cell = ws.cell(row=row, column=2, value=formula)
            value_cell.font = Font(size=11, bold=True)
            value_cell.alignment = Alignment(horizontal="center")
            value_cell.border = BORDER
            value_cell.fill = FORMULA_FILL
            value_cell.protection = Protection(locked=True)
        else:
            for col in (1, 2):
                ws.cell(row=row, column=col).fill = SUBHEADER_FILL
            ws.cell(row=row, column=2).border = BORDER
            ws.cell(row=row, column=2).protection = Protection(locked=True)
        note_cell = ws.cell(row=row, column=3, value=note)
        note_cell.font = NOTE_FONT
        note_cell.protection = Protection(locked=True)
        row += 1

    for label, formula in TYPE_SUMMARY_ROWS:
        row += 1
        cell = ws.cell(row=row, column=1, value=f"    {label}")
        cell.font = BODY_FONT
        cell.protection = Protection(locked=True)
        value_cell = ws.cell(row=row, column=2, value=formula)
        value_cell.font = Font(size=10, bold=True)
        value_cell.alignment = Alignment(horizontal="center")
        value_cell.fill = FORMULA_FILL
        value_cell.protection = Protection(locked=True)

    row += 2
    note = ws.cell(row=row, column=1, value=SUMMARY_NOTE)
    note.font = NOTE_FONT
    note.alignment = Alignment(wrap_text=True)
    note.protection = Protection(locked=True)
    ws.freeze_panes = "A2"


def write_org_sheet(ws: Worksheet, org: dict[str, str]) -> None:
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 64
    for col, title in enumerate(("Параметр", "Значение"), start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
        cell.protection = Protection(locked=True)
    row = 2
    for key in ORG_KEY.values():
        key_cell = ws.cell(row=row, column=1, value=key)
        key_cell.font = Font(size=10, bold=True)
        key_cell.border = BORDER
        key_cell.protection = Protection(locked=True)
        value_cell = ws.cell(row=row, column=2, value=org.get(key, ""))
        value_cell.font = BODY_FONT
        value_cell.border = BORDER
        value_cell.protection = Protection(locked=True)
        row += 1
    ws.freeze_panes = "A2"


# --------------------------------------------------------------------------- #
# Защита
# --------------------------------------------------------------------------- #

def protect_workbook(wb, password: str) -> None:
    """Защитить все листы и структуру книги.

    Ячейки блокируются полностью: данные меняются через приложение или через
    ручной режим, который снимает защиту на время редактирования.
    """
    for ws in wb.worksheets:
        ws.protection.sheet = True
        ws.protection.password = password
        ws.protection.formatCells = True
        ws.protection.formatColumns = True
        ws.protection.formatRows = True
        ws.protection.insertColumns = True
        ws.protection.insertRows = True
        ws.protection.deleteColumns = True
        ws.protection.deleteRows = True
        ws.protection.sort = True
        ws.protection.autoFilter = False
        ws.protection.selectLockedCells = False
        ws.protection.selectUnlockedCells = False
    wb.security = WorkbookProtection(
        workbookPassword=password,
        lockStructure=True,
        lockWindows=False,
    )


def unprotect_workbook(wb) -> None:
    """Снять защиту со всех листов и структуры (для ручного режима)."""
    for ws in wb.worksheets:
        ws.protection.sheet = False
        ws.protection.password = None
    wb.security = WorkbookProtection(lockStructure=False)


def is_protected(wb) -> bool:
    return any(getattr(ws.protection, "sheet", False) for ws in wb.worksheets)


def protect_file(path, password: str) -> None:
    """Защитить уже существующий файл (для миграций и ручного режима)."""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    try:
        protect_workbook(wb, password)
        wb.save(path)
    finally:
        wb.close()


def unprotect_file(path) -> None:
    from openpyxl import load_workbook

    wb = load_workbook(path)
    try:
        unprotect_workbook(wb)
        wb.save(path)
    finally:
        wb.close()


# --------------------------------------------------------------------------- #
# Применение оформления к готовой книге
# --------------------------------------------------------------------------- #

def apply_features(wb, *, password: str | None = None, org: dict[str, str] | None = None) -> None:
    """Применить оформление, формулы, проверки и защиту к книге.

    Идемпотентно: можно вызывать для новой книги и для миграции.
    """
    for sheet in SHEETS:
        if sheet.name not in wb.sheetnames:
            ws = wb.create_sheet(sheet.name)
        else:
            ws = wb[sheet.name]

        if sheet.kind == "info":
            write_instruction_sheet(ws)
            continue
        if sheet.kind == "summary":
            write_summary_sheet(ws)
            continue
        if sheet.kind == "org":
            write_org_sheet(ws, org or {})
            continue

        _style_header(ws, sheet.columns, sheet.formulas)
        ws.auto_filter.ref = (
            f"A1:{get_column_letter(len(sheet.columns))}{max(ws.max_row, 2)}"
        )
        last_row = max(ws.max_row, 2)
        write_formula_columns(ws, sheet, last_row=last_row)
        if sheet.name == "Инструктажи":
            _add_training_highlighting(ws, sheet)
        _add_validation(ws, sheet)

    # Стабильный порядок листов как в схеме.
    order = [s.name for s in SHEETS]
    wb._sheets.sort(key=lambda ws: order.index(ws.title) if ws.title in order else 999)

    if password:
        protect_workbook(wb, password)
