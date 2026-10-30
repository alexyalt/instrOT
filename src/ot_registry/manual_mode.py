"""Ручной режим Excel: открыть книгу, зафиксировать состояние, проверить правки.

Порядок:
1. Пользователь выбирает «Открыть Excel» → создаётся контрольная копия и
   снимок данных, запись через приложение для этой организации блокируется.
2. Правки вносятся в Excel.
3. Пользователь нажимает «Проверить и перечитать» → сравнивается книга с
   контрольной копией по ID, выдаётся отчёт; при ошибках правки не принимаются.
4. Принятие обновляет снимок и снимает блокировку.

Сравнение идёт по постоянным ID, поэтому сортировка строк не создаёт ложных
изменений.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from . import config
from .errors import AppError, ConflictError, NotFoundError, ValidationError
from .excel_store import Book, fingerprint, new_id, open_book, save_book
from .schema import SHEETS, SHEET_BY_NAME

MODE_FILE = "manual_mode.json"
SNAPSHOT_FILE = "manual_snapshot.json"
CHECK_COPY = "manual_before_check.xlsx"

SYSTEM_FIELDS = {
    "Работники": {"id"},
    "Инструктирующие": {"id", "person_id"},
    "Инструктажи": {"id", "number_ot", "number_type", "case_id", "pack_id", "created_at"},
    "История_работников": {"id", "worker_id"},
}


def _mode_path(book: Book) -> Path:
    return book.path.parent / MODE_FILE


def _snapshot_path(book: Book) -> Path:
    return book.path.parent / SNAPSHOT_FILE


@dataclass
class ManualState:
    active: bool = False
    opened_at: str = ""
    check_copy: str = ""
    note: str = ""


def read_state(book: Book) -> ManualState:
    path = _mode_path(book)
    if not path.exists():
        return ManualState()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return ManualState()
    return ManualState(
        active=bool(data.get("active")),
        opened_at=str(data.get("opened_at") or ""),
        check_copy=str(data.get("check_copy") or ""),
        note=str(data.get("note") or ""),
    )


def _write_state(book: Book, state: ManualState) -> None:
    _mode_path(book).write_text(
        json.dumps(
            {
                "active": state.active,
                "opened_at": state.opened_at,
                "check_copy": state.check_copy,
                "note": state.note,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def is_locked(book: Book) -> bool:
    return read_state(book).active


def ensure_unlocked(book: Book) -> None:
    """Запретить запись через приложение, пока идёт ручное редактирование."""
    state = read_state(book)
    if state.active:
        raise ConflictError(
            "Книга открыта для ручного редактирования. Закройте Excel, нажмите "
            "«Проверить и перечитать» и подтвердите изменения — после этого запись "
            "через приложение возобновится."
        )


# --------------------------------------------------------------------------- #
# Снимок данных
# --------------------------------------------------------------------------- #

def _serialize(book: Book) -> dict[str, Any]:
    result: dict[str, Any] = {"taken_at": datetime.now().isoformat(timespec="seconds"),
                              "tables": {}, "org": dict(book.org)}
    for name, table in book.tables.items():
        rows = []
        for row in table.rows:
            clean = {}
            for key, value in row.items():
                if isinstance(value, date):
                    clean[key] = value.isoformat()
                else:
                    clean[key] = "" if value is None else str(value)
            rows.append(clean)
        result["tables"][name] = rows
    return result


def _load_snapshot(book: Book) -> dict[str, Any] | None:
    path = _snapshot_path(book)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _row_key(sheet: str, row: dict[str, Any]) -> str:
    if sheet in ("Работники", "Инструктирующие", "Инструктажи", "История_работников"):
        return str(row.get("id") or "")
    if sheet == "Архив":
        return f"{row.get('case_id','')}-{row.get('pack_id','')}"
    if sheet == "Справочники":
        return f"{row.get('kind','')}:{row.get('value','')}"
    return ""


def _normalize(row: dict[str, Any]) -> dict[str, str]:
    return {
        key: (value.isoformat() if isinstance(value, date) else str(value or ""))
        for key, value in row.items()
    }


# --------------------------------------------------------------------------- #
# Открытие для ручного редактирования
# --------------------------------------------------------------------------- #

def open_for_manual(book: Book) -> dict[str, str]:
    """Создать контрольную копию и снимок, заблокировать запись."""
    state = read_state(book)
    if state.active:
        raise ConflictError(
            "Ручной режим уже включён. Закройте Excel и выполните проверку книги."
        )
    if not book.path.exists():
        raise NotFoundError("Файл книги не найден.")

    folder = book.path.parent
    check_copy = folder / CHECK_COPY
    shutil.copy2(book.path, check_copy)

    snapshot = _serialize(book)
    snapshot["check_copy_fingerprint"] = fingerprint(check_copy)
    _snapshot_path(book).write_text(
        json.dumps(snapshot, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    _write_state(book, ManualState(active=True, opened_at=datetime.now().isoformat(timespec="seconds"),
                                   check_copy=str(check_copy)))
    return {"copy": str(check_copy), "path": str(book.path)}


def launch_excel(book: Book) -> str:
    """Открыть книгу в Excel (или другом приложении по умолчанию)."""
    path = str(book.path)
    try:
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
            return "Excel открыт."
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.Popen([opener, path])
        return "Книга открыта приложением по умолчанию."
    except Exception:
        return f"Не удалось открыть файл автоматически. Откройте вручную: {path}"


# --------------------------------------------------------------------------- #
# Проверка после ручного редактирования
# --------------------------------------------------------------------------- #

@dataclass
class Diff:
    sheet: str
    key: str
    kind: str  # added / removed / changed
    fields: list[str] = field(default_factory=list)
    blocked: bool = False
    note: str = ""

    def describe(self) -> str:
        return {
            "added": "добавлена строка",
            "removed": "удалена строка",
            "changed": "изменены поля: " + ", ".join(self.fields),
        }[self.kind]


def compare_with_snapshot(book: Book, snapshot: dict[str, Any]) -> tuple[list[Diff], list[str]]:
    """Сравнить текущую книгу со снимком по ID.

    Строки без ID (добавленные вручную в Excel) считаются новыми: для них
    формируется отдельный ключ по содержимому, чтобы изменение не терялось.
    """
    diffs: list[Diff] = []
    problems: list[str] = []

    for sheet in SHEETS:
        name = sheet.name
        if name in ("Организация", "Инструкция"):
            continue

        before_rows = snapshot.get("tables", {}).get(name, [])
        before: dict[str, dict[str, str]] = {}
        unkeyed_before: list[dict[str, str]] = []
        for raw in before_rows:
            key = _row_key(name, raw)
            normalized = _normalize(raw)
            if key:
                before[key] = normalized
            else:
                unkeyed_before.append(normalized)

        after: dict[str, dict[str, str]] = {}
        added_unkeyed: list[tuple[int, dict[str, str]]] = []
        for index, row in enumerate(book.tables[name].rows):
            key = _row_key(name, row)
            normalized = _normalize(row)
            if key:
                after[key] = normalized
            else:
                added_unkeyed.append((index, normalized))

        system = SYSTEM_FIELDS.get(name, set())

        # Изменённые существующие строки.
        for key, row in after.items():
            if key not in before:
                diffs.append(Diff(name, key, "added", note="новая строка"))
                continue
            changed: list[str] = []
            blocked_fields: list[str] = []
            for field_key, value in row.items():
                old = str(before[key].get(field_key, "") or "")
                if old != str(value or ""):
                    changed.append(field_key)
                    if field_key in system:
                        blocked_fields.append(field_key)
            if changed:
                diffs.append(
                    Diff(
                        name,
                        key,
                        "changed",
                        fields=changed,
                        blocked=bool(blocked_fields),
                        note=(
                            "нельзя менять служебные поля: " + ", ".join(blocked_fields)
                            if blocked_fields else ""
                        ),
                    )
                )

        # Новые строки, добавленные в Excel без ID.
        for index, row in added_unkeyed:
            non_empty = {k: v for k, v in row.items() if v}
            label = non_empty.get("fio") or non_empty.get("number_ot") or f"строка {index + 2}"
            diffs.append(
                Diff(name, f"new:{index}", "added",
                     note=f"новая строка без ID: {label} — ID будет назначен при принятии")
            )
            if name in ("Инструктажи", "Работники"):
                if not row.get("fio") and name == "Инструктажи":
                    problems.append(
                        f"{name}: в новой строке без ID не указано ФИО — "
                        "заполните её в Excel или удалите."
                    )

        # Пропавшие строки.
        for key, row in before.items():
            if key not in after:
                blocked = name in ("Инструктажи", "История_работников")
                diffs.append(
                    Diff(
                        name, key, "removed", blocked=blocked,
                        note=("удаление проведённых записей и истории недопустимо"
                              if blocked else ""),
                    )
                )

    for item in diffs:
        if item.blocked:
            problems.append(
                f"{item.sheet}: {item.describe()}" + (f" — {item.note}" if item.note else "")
            )
    return diffs, problems


def check_book(book: Book) -> dict[str, Any]:
    """Прочитать книгу заново и сравнить со снимком.

    Возвращает отчёт. Если есть блокирующие ошибки, принятие невозможно.
    """
    state = read_state(book)
    snapshot = _load_snapshot(book)
    if snapshot is None:
        raise ConflictError(
            "Нет контрольного снимка для сверки. Сначала включите ручной режим "
            "или восстановите книгу из резервной копии."
        )

    try:
        fresh = open_book(book.path, data_dir=book.data_dir)
    except AppError as exc:
        return {"ok": False, "fatal": exc.message, "diffs": [], "problems": [exc.message],
                "state": state}

    diffs, problems = compare_with_snapshot(fresh, snapshot)

    # Дополнительные проверки целостности после ручных правок.
    problems.extend(_validate_after_manual(fresh))

    return {
        "ok": not problems,
        "fatal": "",
        "diffs": diffs,
        "problems": problems,
        "state": state,
        "book": fresh,
    }


def _validate_after_manual(book: Book) -> list[str]:
    problems: list[str] = []

    # Уникальность ID и номеров.
    ids = [str(r.get("id") or "") for r in book.table("Инструктажи").rows]
    ids = [i for i in ids if i]
    duplicates = {i for i in ids if ids.count(i) > 1}
    if duplicates:
        problems.append("Инструктажи: повторяются ID строк — " + ", ".join(sorted(duplicates)[:5]))

    numbers = [str(r.get("number_ot") or "") for r in book.table("Инструктажи").rows if r.get("number_ot")]
    dup_numbers = {n for n in numbers if numbers.count(n) > 1}
    if dup_numbers:
        problems.append("Инструктажи: повторяются номера ОТ — " + ", ".join(sorted(dup_numbers)[:5]))

    # Ссылки на работников и инструктирующих.
    worker_ids = {str(r.get("id") or "") for r in book.table("Работники").rows}
    instructor_ids = {str(r.get("id") or "") for r in book.table("Инструктирующие").rows}
    for row in book.table("Инструктажи").rows:
        worker_id = str(row.get("worker_id") or "")
        if worker_id and worker_id not in worker_ids:
            problems.append(
                f"Инструктажи: запись {str(row.get('id'))[:8]} ссылается на отсутствующего работника."
            )
        trainer_id = str(row.get("instructor_id") or "")
        if trainer_id and trainer_id not in instructor_ids:
            problems.append(
                f"Инструктажи: запись {str(row.get('id'))[:8]} ссылается на отсутствующего инструктирующего."
            )

    # Новые строки без ID.
    missing_ids = [r for r in book.table("Инструктажи").rows if not str(r.get("id") or "")]
    if missing_ids:
        problems.append(
            f"Инструктажи: {len(missing_ids)} новых строк без ID. "
            "Сохраните книгу и нажмите «Принять изменения» — ID будут назначены."
        )
    return problems


def accept_changes(book: Book, *, reason: str = "") -> dict[str, Any]:
    """Принять ручные правки: назначить ID новым строкам, обновить снимок."""
    report = check_book(book)
    if report.get("fatal"):
        raise ConflictError(report["fatal"])
    if not report["ok"]:
        raise ValidationError(
            "Книга содержит ошибки. Исправьте их в Excel и проверьте снова.",
            details=report["problems"],
        )

    fresh: Book = report["book"]
    assigned = 0
    for sheet_name in ("Работники", "Инструктирующие", "Инструктажи", "История_работников"):
        table = fresh.tables[sheet_name]
        for row in table.rows:
            if not str(row.get("id") or ""):
                row["id"] = new_id()
                assigned += 1
            if sheet_name == "Инструктирующие" and not str(row.get("person_id") or ""):
                row["person_id"] = row["id"]

    # Блокировка снимается до записи: иначе сохранение отклонит само себя.
    # Если запись не удастся, блокировка возвращается — книга останется
    # в состоянии «требуется проверка».
    previous_state = read_state(fresh)
    _write_state(fresh, ManualState(active=False))
    try:
        save_book(
            fresh,
            action="manual_excel_accept",
            entity="Excel",
            reason=reason or "Принятие ручных правок",
            details={"assigned_ids": assigned},
            make_backup=True,
        )
    except Exception:
        _write_state(fresh, previous_state)
        raise

    snapshot = _serialize(fresh)
    _snapshot_path(fresh).write_text(
        json.dumps(snapshot, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    _write_state(fresh, ManualState(active=False))
    return {"assigned": assigned, "report": report}


def cancel_manual(book: Book, *, restore: bool = False) -> str:
    """Отменить ручной режим. При restore — вернуть контрольную копию."""
    state = read_state(book)
    if not state.active:
        raise ConflictError("Ручной режим не включён.")
    message = "Ручной режим выключен, книга перечитана."
    if restore:
        copy_path = Path(state.check_copy) if state.check_copy else book.path.parent / CHECK_COPY
        if not copy_path.exists():
            raise NotFoundError("Контрольная копия не найдена — восстановление невозможно.")
        backup_dir = book.backups_dir()
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(book.path, backup_dir / f"before_restore_{stamp}.xlsx")
        shutil.copy2(copy_path, book.path)
        message = "Книга восстановлена из контрольной копии, сделанной перед ручным режимом."
    _write_state(book, ManualState(active=False))
    return message
