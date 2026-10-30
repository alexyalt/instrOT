"""Работа с книгой Excel: чтение, запись, резерв и безопасная замена файла.

Принципы:
- одна книга = одна организация;
- запись только через этот модуль, никогда напрямую из веб-слоя;
- перед каждой записью создаётся резервная копия;
- новая книга сохраняется во временный файл, проверяется и только потом
  заменяет рабочую (атомарная замена на том же томе);
- перед заменой проверяется, что исходный файл не изменился с момента чтения.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import tempfile
import uuid
from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from . import config
from .errors import BookChangedError, BookLockedError, ConfigError, ValidationError
from .locking import FileLock
from .schema import ORG_KEY, SHEETS, Column, Sheet

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
BODY_FONT = Font(size=10)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DATE_FORMAT = "DD.MM.YYYY"

INSTRUCTION_LINES = (
    "Как работать с книгой",
    "",
    "1. Обычная работа — через приложение. Не открывайте файл в Excel одновременно с ним.",
    "2. Чтобы править вручную: выберите «Открыть Excel», после закрытия — «Проверить и перечитать».",
    "3. Колонку ID не удаляйте и не меняйте: по ней приложение узнаёт записи.",
    "4. Даты вводите как ДД.ММ.ГГГГ. Статусы выбирайте из существующих значений.",
    "5. Новые строки добавляйте в конец таблицы, не разрывая её.",
    "6. Код 000000 означает бумажный бланк, заполненный вручную вне системы.",
    "7. Не изменяйте выданные номера (ОТ, по виду, дело, пакет) без необходимости.",
    "8. Резервные копии создаются автоматически перед каждой записью приложения.",
)


def new_id() -> str:
    return uuid.uuid4().hex


def fingerprint(path: Path) -> str:
    """Отпечаток файла: содержимое + размер. Основание для проверки конфликта."""
    path = Path(path)
    if not path.exists():
        return ""
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 256), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _as_date(value: Any) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValidationError(f"Не удалось разобрать дату: «{value}». Используйте ДД.ММ.ГГГГ.")


def _as_int(value: Any) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(float(str(value).replace(",", ".")))
    except (TypeError, ValueError):
        raise ValidationError(f"Ожидалось целое число, получено «{value}».") from None


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    return str(value).strip()


def _coerce_read(col: Column, raw: Any) -> Any:
    if col.type == "date":
        return _as_date(raw)
    if col.type == "int":
        return _as_int(raw)
    return _as_text(raw)


def _coerce_write(col: Column, value: Any) -> Any:
    if col.type == "date":
        return value if isinstance(value, date) else _as_date(value)
    if col.type == "int":
        return _as_int(value)
    return "" if value is None else str(value)


@dataclass
class Table:
    """Набор строк листа: список словарей по ключам колонок."""

    sheet: Sheet
    rows: list[dict[str, Any]] = field(default_factory=list)

    def by_key(self, key: str, value: str) -> dict[str, Any] | None:
        for row in self.rows:
            if str(row.get(key) or "") == value:
                return row
        return None

    def add(self, values: dict[str, Any]) -> dict[str, Any]:
        row = {col.key: None for col in self.sheet.columns}
        row.update(values)
        self.rows.append(row)
        return row


@dataclass
class Book:
    """Открытая книга организации."""

    path: Path
    wb: Workbook
    org: dict[str, str]
    tables: dict[str, Table]
    fingerprint: str
    data_dir: Path

    @property
    def org_uuid(self) -> str:
        return str(self.org.get(ORG_KEY["org_uuid"], ""))

    @property
    def name(self) -> str:
        return str(self.org.get(ORG_KEY["name"], ""))

    def table(self, sheet_name: str) -> Table:
        return self.tables[sheet_name]

    def org_value(self, key: str, default: str = "") -> str:
        return str(self.org.get(ORG_KEY.get(key, key), default))

    def set_org_value(self, key: str, value: Any) -> None:
        if value is None:
            self.org[ORG_KEY.get(key, key)] = ""
        elif isinstance(value, date):
            self.org[ORG_KEY.get(key, key)] = value.strftime("%d.%m.%Y")
        else:
            self.org[ORG_KEY.get(key, key)] = str(value)

    def backups_dir(self) -> Path:
        return self.data_dir / "backups"

    def logs_dir(self) -> Path:
        return self.data_dir / "logs"


# --------------------------------------------------------------------------- #
# Создание книги
# --------------------------------------------------------------------------- #

def _style_header(ws: Worksheet, columns: Iterable[Column]) -> None:
    for idx, col in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(idx)].width = col.width
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"


def _add_training_highlighting(ws: Worksheet, sheet: Sheet) -> None:
    """Подсветка в самом Excel: просрочено и ближайшие сроки.

    Правила не меняют данные, только оформляют строки. Приложение считает
    сроки самостоятельно и не зависит от этих правил.
    """
    keys = [c.key for c in sheet.columns]
    if "plan_date" not in keys or "status" not in keys:
        return
    plan_col = get_column_letter(keys.index("plan_date") + 1)
    status_col = get_column_letter(keys.index("status") + 1)
    last = get_column_letter(len(sheet.columns))
    rng = f"A2:{last}5000"
    rules = (
        FormulaRule(
            formula=[f'AND(${status_col}2="Запланирован",${plan_col}2<>"",${plan_col}2<TODAY())'],
            fill=PatternFill("solid", fgColor="FFC7CE"),
            font=Font(color="9C0006", size=10),
        ),
        FormulaRule(
            formula=[
                f'AND(${status_col}2="Запланирован",${plan_col}2<>"",'
                f'${plan_col}2>=TODAY(),${plan_col}2<=TODAY()+{config.DEFAULT_SOON_DAYS})'
            ],
            fill=PatternFill("solid", fgColor="FFD9B3"),
        ),
        FormulaRule(
            formula=[
                f'AND(${status_col}2="Запланирован",${plan_col}2<>"",'
                f'${plan_col}2>TODAY()+{config.DEFAULT_SOON_DAYS},'
                f'${plan_col}2<=TODAY()+{config.DEFAULT_UPCOMING_DAYS})'
            ],
            fill=PatternFill("solid", fgColor="FFF2CC"),
        ),
    )
    for rule in rules:
        ws.conditional_formatting.add(rng, rule)


def create_book(
    path: Path,
    *,
    data_dir: Path,
    name: str,
    full_name: str = "",
    inn: str = "",
    kpp: str = "",
    ogrn: str = "",
    address: str = "",
    director: str = "",
    ot_specialist: str = "",
) -> Book:
    """Создать новую книгу организации по текущей схеме."""
    path = Path(path)
    if path.exists():
        raise ConfigError(f"Файл уже существует: {path.name}")
    path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)

    for sheet in SHEETS:
        ws = wb.create_sheet(sheet.name)
        for idx, col in enumerate(sheet.columns, start=1):
            ws.cell(row=1, column=idx, value=col.title)
        _style_header(ws, sheet.columns)
        if sheet.name == "Инструктажи":
            _add_training_highlighting(ws, sheet)
        if sheet.name == "Инструкция":
            ws.column_dimensions["A"].width = 110
            for r, line in enumerate(INSTRUCTION_LINES, start=2):
                cell = ws.cell(row=r, column=1, value=line)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.font = Font(size=10, bold=(r == 2))

    book = Book(
        path=path,
        wb=wb,
        org={},
        tables={s.name: Table(s, []) for s in SHEETS},
        fingerprint="",
        data_dir=data_dir,
    )

    now = datetime.now().astimezone()
    org_values: dict[str, Any] = {
        "org_uuid": new_id(),
        "book_uuid": new_id(),
        "schema_version": config.SCHEMA_VERSION,
        "name": name,
        "full_name": full_name or name,
        "inn": inn,
        "kpp": kpp,
        "ogrn": ogrn,
        "address": address,
        "director": director,
        "ot_specialist": ot_specialist,
        "soon_days": config.DEFAULT_SOON_DAYS,
        "upcoming_days": config.DEFAULT_UPCOMING_DAYS,
        "default_period": config.DEFAULT_PERIOD_MONTHS,
        "created_at": now.date(),
        "updated_at": now.date(),
        "counter_ot": 0,
        "counter_type": "{}",
        "counter_case": 0,
        "counter_pack": 0,
    }
    for key, value in org_values.items():
        book.set_org_value(key, value)

    save_book(
        book,
        action="create_book",
        reason="Создание книги организации",
        make_backup=False,
        initial=True,
    )
    return book


# --------------------------------------------------------------------------- #
# Чтение
# --------------------------------------------------------------------------- #

def open_book(path: Path, *, data_dir: Path) -> Book:
    """Прочитать книгу с диска."""
    path = Path(path)
    if not path.exists():
        raise ConfigError(f"Файл книги не найден: {path}")
    try:
        wb = load_workbook(path, data_only=False)
    except Exception as exc:  # openpyxl поднимает разные типы
        raise ConfigError(
            f"Не удалось открыть книгу «{path.name}». Файл повреждён или это не XLSX."
        ) from exc

    missing = [s.name for s in SHEETS if s.name not in wb.sheetnames]
    if missing:
        raise ConfigError(
            "Книга не соответствует формату приложения. "
            f"Отсутствуют листы: {', '.join(missing)}."
        )

    tables: dict[str, Table] = {}
    for sheet in SHEETS:
        ws = wb[sheet.name]
        title_to_col = {col.title: col for col in sheet.columns}
        header: dict[int, Column] = {}
        for idx in range(1, len(sheet.columns) + 1):
            title = str(ws.cell(row=1, column=idx).value or "").strip()
            if not title:
                continue
            if title not in title_to_col:
                raise ConfigError(
                    f"Лист «{sheet.name}» изменён: неизвестная колонка «{title}». "
                    "Восстановите файл из резервной копии."
                )
            header[idx] = title_to_col[title]

        rows: list[dict[str, Any]] = []
        for row_idx in range(2, ws.max_row + 1):
            values: dict[str, Any] = {}
            empty = True
            for idx, col in header.items():
                raw = ws.cell(row=row_idx, column=idx).value
                if raw not in (None, ""):
                    empty = False
                values[col.key] = _coerce_read(col, raw)
            if not empty:
                rows.append(values)
        tables[sheet.name] = Table(sheet, rows)

    org_ws = wb["Организация"]
    org: dict[str, str] = {}
    for row_idx in range(2, org_ws.max_row + 1):
        key = str(org_ws.cell(row=row_idx, column=1).value or "").strip()
        if not key:
            continue
        value = org_ws.cell(row=row_idx, column=2).value
        org[key] = _as_text(value)

    return Book(
        path=path,
        wb=wb,
        org=org,
        tables=tables,
        fingerprint=fingerprint(path),
        data_dir=data_dir,
    )


# --------------------------------------------------------------------------- #
# Сохранение
# --------------------------------------------------------------------------- #

def _write_org(ws: Worksheet, org: dict[str, str]) -> None:
    """Переписать лист «Организация» по фиксированному порядку ключей."""
    existing: dict[str, int] = {}
    for row_idx in range(2, ws.max_row + 1):
        key = ws.cell(row=row_idx, column=1).value
        if key:
            existing[str(key).strip()] = row_idx

    row_idx = 2
    for key in ORG_KEY.values():
        target = existing.get(key, row_idx)
        cell = ws.cell(row=target, column=1, value=key)
        cell.font = Font(size=10, bold=True)
        cell.alignment = Alignment(vertical="top")
        cell.border = BORDER
        value_cell = ws.cell(row=target, column=2, value=org.get(key, ""))
        value_cell.font = BODY_FONT
        value_cell.alignment = Alignment(vertical="top")
        value_cell.border = BORDER
        row_idx = max(target, row_idx) + 1

    # Удалить лишние строки, если ключей стало меньше.
    if ws.max_row > len(ORG_KEY) + 1:
        ws.delete_rows(len(ORG_KEY) + 2, ws.max_row - len(ORG_KEY) - 1)


def _write_table(ws: Worksheet, table: Table) -> None:
    """Переписать строки листа, сохранив оформление шаблонной строки."""
    sheet = table.sheet
    last_row = ws.max_row
    template_row = 2 if last_row >= 2 else 1

    for pos, row in enumerate(table.rows, start=2):
        for col_idx, col in enumerate(sheet.columns, start=1):
            cell = ws.cell(row=pos, column=col_idx)
            if pos > last_row:
                src = ws.cell(row=template_row, column=col_idx)
                if src.has_style:
                    cell._style = copy(src._style)
            cell.value = _coerce_write(col, row.get(col.key))
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=False)
            if col.type == "date":
                cell.number_format = DATE_FORMAT
            elif col.type == "int":
                cell.number_format = "0"

    if last_row > len(table.rows) + 1:
        ws.delete_rows(len(table.rows) + 2, last_row - len(table.rows) - 1)

    last_col = get_column_letter(len(sheet.columns))
    ws.auto_filter.ref = f"A1:{last_col}{max(len(table.rows) + 1, 2)}"


def _validate_saved(path: Path) -> None:
    """Проверить, что сохранённый файл открывается и содержит нужные листы."""
    try:
        wb = load_workbook(path, read_only=True)
    except Exception as exc:
        raise ValidationError(
            "Сохранённый файл не открывается. Запись отменена, данные не изменены."
        ) from exc
    try:
        names = set(wb.sheetnames)
    finally:
        wb.close()
    missing = [s.name for s in SHEETS if s.name not in names]
    if missing:
        raise ValidationError(
            "Сохранённый файл не содержит обязательных листов: " + ", ".join(missing)
        )


def _make_backup(book: Book) -> Path | None:
    """Резервная копия текущего файла перед изменением."""
    if not book.path.exists():
        return None
    backups = book.backups_dir()
    backups.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    target = backups / f"registry_{stamp}.xlsx"
    shutil.copy2(book.path, target)
    rotate_backups(backups)
    return target


def rotate_backups(backups: Path, keep: int | None = None) -> None:
    """Удалить старые копии, оставив последние `keep` штук."""
    keep = config.DEFAULT_BACKUP_KEEP if keep is None else keep
    files = sorted(backups.glob("registry_*.xlsx"))
    for old in files[: max(len(files) - keep, 0)]:
        try:
            old.unlink()
        except OSError:
            pass


def save_book(
    book: Book,
    *,
    action: str,
    reason: str = "",
    entity: str = "",
    entity_id: str = "",
    expected_fingerprint: str | None = None,
    make_backup: bool = True,
    initial: bool = False,
    details: dict[str, Any] | None = None,
) -> None:
    """Сохранить книгу: резерв → временный файл → проверка → замена.

    `expected_fingerprint` — отпечаток на момент чтения. Если файл изменился
    (ручная правка в Excel, другой процесс), запись отменяется.
    """
    path = book.path
    lock_path = path.with_suffix(".lock")
    expected = book.fingerprint if expected_fingerprint is None else expected_fingerprint

    with FileLock(lock_path):
        if not initial and expected:
            current = fingerprint(path) if path.exists() else ""
            if current and current != expected:
                raise BookChangedError(
                    "Файл книги изменён вне приложения (например, в Excel). "
                    "Изменения не сохранены. Откройте форму заново и повторите ввод."
                )

        book.set_org_value("schema_version", config.SCHEMA_VERSION)
        book.set_org_value("updated_at", datetime.now().date())

        _write_org(book.wb["Организация"], book.org)
        for name, table in book.tables.items():
            if name == "Организация":
                continue
            _write_table(book.wb[name], table)

        backup_path: Path | None = None
        if make_backup:
            backup_path = _make_backup(book)

        fd, tmp_name = tempfile.mkstemp(
            prefix=".registry_", suffix=".xlsx", dir=str(path.parent)
        )
        os.close(fd)
        tmp_path = Path(tmp_name)
        try:
            book.wb.save(tmp_path)
            _validate_saved(tmp_path)
            if not initial and expected:
                current = fingerprint(path) if path.exists() else ""
                if current and current != expected:
                    raise BookChangedError(
                        "Файл книги изменился во время сохранения. "
                        "Данные не записаны, повторите попытку."
                    )
            _replace_file(tmp_path, path)
        except Exception:
            tmp_path.unlink(missing_ok=True)
            if backup_path and backup_path.exists() and not path.exists():
                shutil.copy2(backup_path, path)
            raise

        book.fingerprint = fingerprint(path)

    from .journal import log_event

    log_event(
        book.logs_dir(),
        organization=book.name or book.path.stem,
        action=action,
        entity=entity,
        entity_id=entity_id,
        reason=reason,
        details=details or {},
    )


def _replace_file(src: Path, dst: Path) -> None:
    """Атомарная замена файла в пределах одного тома."""
    import time

    for attempt in range(3):
        try:
            os.replace(src, dst)
            return
        except PermissionError:
            if attempt == 2:
                raise BookLockedError(
                    "Не удалось записать файл книги: он занят другой программой. "
                    "Закройте книгу в Excel и повторите сохранение."
                ) from None
            time.sleep(0.4)
