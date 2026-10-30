"""Сценарии работы: работники, планы, регистрация, бумажные бланки, архив.

Модуль работает с моделью книги (`excel_store.Book`) и не знает о Flask.
Каждая изменяющая функция сама сохраняет книгу через `save_book`, поэтому
веб-слой не должен вызывать сохранение повторно.
"""

from __future__ import annotations

import json
import random
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Iterable

from . import config, deadlines
from .errors import ConflictError, NotFoundError, ValidationError
from .excel_store import Book, new_id, save_book
from .schema import ORG_KEY
from . import validation

# --------------------------------------------------------------------------- #
# Справочные выборки
# --------------------------------------------------------------------------- #


def workers(book: Book, *, only_active: bool = True) -> list[dict[str, Any]]:
    rows = book.table("Работники").rows
    if only_active:
        rows = [r for r in rows if (r.get("status") or "Работает") == "Работает"]
    return sorted(rows, key=lambda r: str(r.get("fio") or "").lower())


def worker_by_id(book: Book, worker_id: str) -> dict[str, Any] | None:
    return book.table("Работники").by_key("id", worker_id)


def instructors(book: Book, *, on: date | None = None) -> list[dict[str, Any]]:
    rows = list(book.table("Инструктирующие").rows)
    if on is not None:
        rows = [r for r in rows if deadlines.instructor_valid(r.get("valid_to"), on)]
    return sorted(rows, key=lambda r: str(r.get("fio") or "").lower())


def instructor_by_id(book: Book, instructor_id: str) -> dict[str, Any] | None:
    return book.table("Инструктирующие").by_key("id", instructor_id)


def instructor_types(row: dict[str, Any]) -> set[str]:
    raw = str(row.get("types") or "")
    if not raw.strip():
        return set(config.TRAINING_TYPES)
    return {p.strip() for p in raw.replace(";", ",").split(",") if p.strip()}


def trainings(book: Book) -> list[dict[str, Any]]:
    return list(book.table("Инструктажи").rows)


def planned(book: Book) -> list[dict[str, Any]]:
    return [t for t in trainings(book) if (t.get("status") or "") == "Запланирован"]


def dictionaries(book: Book) -> dict[str, list[str]]:
    """Значения справочников + подсказки из существующих записей."""
    result: dict[str, list[str]] = {kind: [] for kind in ("reason", "position", "department", "program")}
    for row in book.table("Справочники").rows:
        kind = str(row.get("kind") or "")
        value = str(row.get("value") or "").strip()
        if value and kind in result:
            result[kind].append(value)
    for row in book.table("Работники").rows:
        for kind, key in (("position", "position"), ("department", "department")):
            value = str(row.get(key) or "").strip()
            if value:
                result[kind].append(value)
    for row in book.table("Инструктажи").rows:
        for kind, key in (("program", "program"), ("reason", "reason")):
            value = str(row.get(key) or "").strip()
            if value:
                result[kind].append(value)
    for row in book.table("Инструктирующие").rows:
        value = str(row.get("program") or "").strip()
        if value:
            result["program"].append(value)
    return {k: sorted(set(v), key=str.lower) for k, v in result.items()}


def all_positions(book: Book) -> list[str]:
    return dictionaries(book)["position"]


# --------------------------------------------------------------------------- #
# Работники
# --------------------------------------------------------------------------- #

def _history(book: Book, *, worker_id: str, event_type: str, old: str = "",
             new: str = "", reason: str = "", event_date: date | None = None) -> None:
    book.table("История_работников").add(
        {
            "id": new_id(),
            "worker_id": worker_id,
            "event_date": event_date or config.today(),
            "event_type": event_type,
            "old_value": old,
            "new_value": new,
            "reason": reason,
            "created_at": datetime.now().strftime("%d.%m.%Y %H:%M"),
        }
    )


def _find_possible_duplicates(book: Book, fio: str, *, exclude_id: str = "") -> list[str]:
    target = " ".join(fio.lower().split())
    result = []
    for row in book.table("Работники").rows:
        if row.get("id") == exclude_id:
            continue
        if " ".join(str(row.get("fio") or "").lower().split()) == target:
            result.append(str(row.get("id")))
    return result


def add_worker(book: Book, data: dict[str, Any], *, allow_duplicate: bool = False,
               reason: str = "") -> dict[str, Any]:
    clean = validation.validate_worker(data)
    duplicates = _find_possible_duplicates(book, clean["fio"])
    if duplicates and not allow_duplicate:
        raise ConflictError(
            "Работник с таким ФИО уже есть. Подтвердите, что это другой человек, "
            "или выберите существующую запись.",
            details=duplicates,
        )
    row = book.table("Работники").add({**clean, "id": new_id(),
                                       "updated_at": datetime.now().strftime("%d.%m.%Y %H:%M")})
    save_book(book, action="worker_add", entity="Работники", entity_id=row["id"],
              reason=reason or "Добавление работника",
              details={"fio": clean["fio"], "position": clean["position"]})
    return row


def update_worker(book: Book, worker_id: str, data: dict[str, Any], *, reason: str = "") -> dict[str, Any]:
    row = worker_by_id(book, worker_id)
    if row is None:
        raise NotFoundError("Работник не найден.")
    clean = validation.validate_worker(data)

    for field, label in (("position", "Должность"), ("department", "Подразделение")):
        if str(row.get(field) or "") != clean[field]:
            _history(book, worker_id=worker_id, event_type=f"Изменение: {label}",
                     old=str(row.get(field) or ""), new=clean[field], reason=reason)
    if (row.get("status") or "") != clean["status"]:
        _history(book, worker_id=worker_id, event_type="Статус",
                 old=str(row.get("status") or ""), new=clean["status"], reason=reason)

    row.update(clean)
    row["updated_at"] = datetime.now().strftime("%d.%m.%Y %H:%M")
    save_book(book, action="worker_update", entity="Работники", entity_id=worker_id,
              reason=reason or "Изменение данных работника", details={"fio": clean["fio"]})
    return row


def transfer_worker(book: Book, worker_id: str, *, position: str, department: str,
                    manager: str = "", reason: str = "") -> dict[str, Any]:
    row = worker_by_id(book, worker_id)
    if row is None:
        raise NotFoundError("Работник не найден.")
    if not position.strip():
        raise ValidationError("Укажите новую должность.", fields={"position": "Укажите новую должность."})
    old = f"{row.get('position') or ''} / {row.get('department') or ''}".strip(" /")
    new = f"{position} / {department}".strip(" /")
    if old == new:
        raise ValidationError("Новая должность и подразделение совпадают с текущими.")
    row["position"] = position.strip()
    row["department"] = department.strip()
    row["manager"] = manager.strip()
    row["updated_at"] = datetime.now().strftime("%d.%m.%Y %H:%M")
    _history(book, worker_id=worker_id, event_type="Перевод", old=old, new=new, reason=reason)
    save_book(book, action="worker_transfer", entity="Работники", entity_id=worker_id,
              reason=reason or "Перевод", details={"old": old, "new": new})
    return row


def dismiss_worker(book: Book, worker_id: str, *, dismiss_date: date,
                   reason: str = "") -> dict[str, Any]:
    row = worker_by_id(book, worker_id)
    if row is None:
        raise NotFoundError("Работник не найден.")
    if row.get("hire_date") and dismiss_date < row["hire_date"]:
        raise ValidationError("Дата увольнения раньше даты приёма.",
                              fields={"dismiss_date": "Дата увольнения раньше даты приёма."})
    row["status"] = "Уволен"
    row["dismiss_date"] = dismiss_date
    row["updated_at"] = datetime.now().strftime("%d.%m.%Y %H:%M")
    _history(book, worker_id=worker_id, event_type="Увольнение",
             old="Работает", new="Уволен", reason=reason, event_date=dismiss_date)

    affected = []
    for item in planned(book):
        if item.get("worker_id") == worker_id:
            item["status"] = "Отменён"
            affected.append(str(item.get("id")))
    save_book(book, action="worker_dismiss", entity="Работники", entity_id=worker_id,
              reason=reason or "Увольнение",
              details={"cancelled_plans": affected})
    return row


def worker_history(book: Book, worker_id: str) -> list[dict[str, Any]]:
    rows = [r for r in book.table("История_работников").rows if r.get("worker_id") == worker_id]
    return sorted(rows, key=lambda r: r.get("event_date") or date.min, reverse=True)


# --------------------------------------------------------------------------- #
# Инструктирующие
# --------------------------------------------------------------------------- #

def add_instructor(book: Book, data: dict[str, Any], *, reason: str = "") -> dict[str, Any]:
    fio = str(data.get("fio") or "").strip()
    if not fio:
        raise ValidationError("Укажите ФИО инструктирующего.", fields={"fio": "Укажите ФИО."})
    person_id = str(data.get("person_id") or "").strip() or new_id()
    row = book.table("Инструктирующие").add(
        {
            "id": new_id(),
            "person_id": person_id,
            "fio": fio,
            "position": str(data.get("position") or "").strip(),
            "program": str(data.get("program") or "").strip(),
            "types": str(data.get("types") or "").strip(),
            "valid_from": data.get("valid_from") or None,
            "valid_to": data.get("valid_to") or None,
            "note": str(data.get("note") or "").strip(),
        }
    )
    save_book(book, action="instructor_add", entity="Инструктирующие",
              entity_id=row["id"], reason=reason or "Добавление допуска",
              details={"fio": fio})
    return row


# --------------------------------------------------------------------------- #
# Планирование
# --------------------------------------------------------------------------- #

def generate_code(book: Book) -> str:
    """Уникальный код бланка среди незавершённых записей организации."""
    used = {
        str(t.get("code") or "")
        for t in trainings(book)
        if str(t.get("status") or "") == "Запланирован"
    }
    for _ in range(500):
        code = str(random.randint(100000, 999999))
        if code not in used:
            return code
    raise ConflictError("Не удалось подобрать свободный код бланка. Проверьте активные планы.")


def plan_training(book: Book, data: dict[str, Any], *, repeat_from: dict[str, Any] | None = None) -> dict[str, Any]:
    clean = validation.validate_plan(data)
    worker = worker_by_id(book, clean["worker_id"])
    if worker is None:
        raise NotFoundError("Работник не найден.", details=["Обновите страницу и выберите работника заново."])
    if (worker.get("status") or "Работает") != "Работает":
        raise ValidationError("Нельзя планировать инструктаж уволенному работнику.")

    trainer = instructor_by_id(book, clean["instructor_id"])
    if trainer is None:
        raise NotFoundError("Инструктирующий не найден.")
    if clean["type"] not in instructor_types(trainer):
        raise ValidationError(
            f"У инструктирующего не указан допуск по виду «{clean['type']}».",
            fields={"instructor_id": "Выберите инструктирующего с нужным допуском."},
        )
    if not deadlines.instructor_valid(trainer.get("valid_to"), clean["plan_date"]):
        raise ValidationError(
            "Срок действия допуска инструктирующего истекает раньше плановой даты.",
            fields={"instructor_id": "Выберите инструктирующего с действующим допуском."},
        )

    duplicates = [
        t for t in planned(book)
        if t.get("worker_id") == clean["worker_id"]
        and t.get("type") == clean["type"]
        and t.get("plan_date") == clean["plan_date"]
    ]
    if duplicates and not data.get("allow_duplicate"):
        raise ConflictError(
            "На эту дату уже есть такой план. Проверьте существующую запись, "
            "чтобы не создать дубль.",
            details=[str(t.get("id")) for t in duplicates],
        )

    code = generate_code(book)
    row = book.table("Инструктажи").add(
        {
            "id": new_id(),
            "code": code,
            "worker_id": clean["worker_id"],
            "fio": str(worker.get("fio") or ""),
            "position": str(worker.get("position") or ""),
            "department": str(worker.get("department") or ""),
            "type": clean["type"],
            "reason": clean["reason"],
            "program": clean["program"],
            "plan_date": clean["plan_date"],
            "fact_date": None,
            "status": "Запланирован",
            "instructor_id": clean["instructor_id"],
            "instructor_fio": str(trainer.get("fio") or ""),
            "source": "Система",
            "paper_confirmed": "",
            "reg_date": None,
            "number_ot": "",
            "number_type": "",
            "case_id": "",
            "pack_id": "",
            "note": clean["note"],
            "created_at": datetime.now().strftime("%d.%m.%Y %H:%M"),
            "updated_at": datetime.now().strftime("%d.%m.%Y %H:%M"),
        }
    )
    save_book(book, action="plan_create", entity="Инструктажи", entity_id=row["id"],
              reason="Планирование инструктажа" + (" (повторный)" if repeat_from else ""),
              details={"code": code, "type": clean["type"], "worker_id": clean["worker_id"]})
    return row


def plan_batch(book: Book, *, worker_ids: Iterable[str], instructor_id: str,
               plan_date: date, training_type: str, program: str = "",
               reason: str = "") -> list[dict[str, Any]]:
    created: list[dict[str, Any]] = []
    failures: list[str] = []
    for worker_id in worker_ids:
        try:
            created.append(
                plan_training(
                    book,
                    {
                        "worker_id": worker_id,
                        "instructor_id": instructor_id,
                        "plan_date": plan_date,
                        "type": training_type,
                        "program": program,
                        "reason": reason,
                    },
                )
            )
        except Exception as exc:  # собираем ошибки по людям, не теряя остальные
            worker = worker_by_id(book, worker_id)
            name = worker.get("fio") if worker else worker_id
            failures.append(f"{name}: {getattr(exc, 'message', str(exc))}")
    if failures:
        raise ValidationError("Часть планов не создана.", details=failures)
    return created


def cancel_plan(book: Book, training_id: str, *, reason: str) -> dict[str, Any]:
    row = book.table("Инструктажи").by_key("id", training_id)
    if row is None:
        raise NotFoundError("Запись инструктажа не найдена.")
    if row.get("status") != "Запланирован":
        raise ConflictError("Отменить можно только запланированный инструктаж.")
    if not reason.strip():
        raise ValidationError("Укажите причину отмены.", fields={"reason": "Укажите причину."})
    row["status"] = "Отменён"
    row["note"] = (str(row.get("note") or "") + f" | Отменён: {reason}").strip(" |")
    row["updated_at"] = datetime.now().strftime("%d.%m.%Y %H:%M")
    save_book(book, action="plan_cancel", entity="Инструктажи", entity_id=training_id,
              reason=reason)
    return row


# --------------------------------------------------------------------------- #
# Нумерация и архив
# --------------------------------------------------------------------------- #

def _counter(book: Book, key: str) -> int:
    raw = book.org_value(key, "0")
    try:
        return int(raw or 0)
    except ValueError:
        return 0


def _bump_counter(book: Book, key: str) -> int:
    value = _counter(book, key) + 1
    book.set_org_value(key, value)
    return value


def _type_counters(book: Book) -> dict[str, int]:
    raw = book.org_value("counter_type", "{}")
    try:
        data = json.loads(raw) if raw else {}
    except json.JSONDecodeError:
        data = {}
    return {k: int(v) for k, v in data.items()}


def _save_type_counters(book: Book, counters: dict[str, int]) -> None:
    book.set_org_value("counter_type", json.dumps(counters, ensure_ascii=False))


def issue_numbers(book: Book, training_type: str, fact_date: date) -> dict[str, str]:
    """Выдать номера: сквозной, по виду, дело и пакет.

    Дело — по виду и году факта. Внутри дела не больше одного открытого пакета.
    """
    prefix = config.TYPE_PREFIX.get(training_type, "ИН")
    ot = _bump_counter(book, "counter_ot")
    type_counters = _type_counters(book)
    type_counters[training_type] = type_counters.get(training_type, 0) + 1
    _save_type_counters(book, type_counters)

    number_ot = f"ОТ{ot:06d}"
    number_type = f"{prefix}{type_counters[training_type]:06d}"

    case_id, pack_id = current_case_and_pack(book, training_type=training_type, fact_date=fact_date)
    return {
        "number_ot": number_ot,
        "number_type": number_type,
        "case_id": case_id,
        "pack_id": pack_id,
    }


def current_case_and_pack(book: Book, *, training_type: str, fact_date: date) -> tuple[str, str]:
    """Найти открытое дело и открытый пакет либо открыть новые."""
    year = fact_date.year
    archive = book.table("Архив")
    case_row = None
    for row in archive.rows:
        if row.get("type") == training_type and row.get("year") == year and row.get("status") == "Открыт":
            case_row = row
            break

    if case_row is None:
        case_number = _bump_counter(book, "counter_case")
        case_id = f"Д{case_number:04d}"
        case_row = archive.add(
            {
                "case_id": case_id,
                "year": year,
                "type": training_type,
                "pack_id": "",
                "title": f"{training_type} инструктажи, {year}",
                "opened_at": fact_date,
                "closed_at": None,
                "status": "Открыт",
                "note": "",
            }
        )

    pack_row = None
    for row in archive.rows:
        if (row.get("case_id") == case_row.get("case_id")
                and row.get("pack_id")
                and row.get("status") == "Открыт"):
            pack_row = row
            break

    if pack_row is None:
        pack_number = _bump_counter(book, "counter_pack")
        pack_id = f"П{pack_number:04d}"
        pack_row = archive.add(
            {
                "case_id": case_row.get("case_id"),
                "year": year,
                "type": training_type,
                "pack_id": pack_id,
                "title": f"Пакет {pack_id}",
                "opened_at": fact_date,
                "closed_at": None,
                "status": "Открыт",
                "note": "",
            }
        )
        for row in archive.rows:
            if row.get("case_id") == case_row.get("case_id") and not row.get("pack_id"):
                row["pack_id"] = pack_id

    return str(case_row.get("case_id") or ""), str(pack_row.get("pack_id") or "")


def close_pack(book: Book, case_id: str, pack_id: str, *, close_date: date,
               reason: str = "") -> dict[str, Any]:
    archive = book.table("Архив")
    row = None
    for item in archive.rows:
        if item.get("case_id") == case_id and item.get("pack_id") == pack_id:
            row = item
            break
    if row is None:
        raise NotFoundError("Пакет не найден.")
    if row.get("status") == "Закрыт":
        raise ConflictError("Этот пакет уже закрыт.")
    row["status"] = "Закрыт"
    row["closed_at"] = close_date
    save_book(book, action="pack_close", entity="Архив",
              entity_id=f"{case_id}-{pack_id}", reason=reason or "Закрытие пакета")
    return row


def close_case(book: Book, case_id: str, *, close_date: date, reason: str = "") -> None:
    archive = book.table("Архив")
    open_packs = [
        r for r in archive.rows
        if r.get("case_id") == case_id and r.get("pack_id") and r.get("status") == "Открыт"
    ]
    if open_packs:
        raise ConflictError(
            "В деле есть открытые пакеты. Закройте их перед закрытием дела.",
            details=[str(r.get("pack_id")) for r in open_packs],
        )
    for row in archive.rows:
        if row.get("case_id") == case_id and not row.get("pack_id"):
            row["status"] = "Закрыт"
            row["closed_at"] = close_date
    save_book(book, action="case_close", entity="Архив", entity_id=case_id,
              reason=reason or "Закрытие дела")
