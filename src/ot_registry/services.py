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

from . import config, deadlines, rules
from .errors import ConflictError, NotFoundError, ValidationError
from .excel_store import Book, new_id, save_book
from .schema import ORG_KEY
from . import validation

# --------------------------------------------------------------------------- #
# Справочные выборки
# --------------------------------------------------------------------------- #

def _effective_period(book: Book, worker: dict[str, Any]) -> int:
    """Периодичность повторного инструктажа для конкретного работника."""
    if worker.get("period_months"):
        try:
            return int(worker["period_months"])
        except (TypeError, ValueError):
            pass
    return _org_period(book)


def _org_period(book: Book) -> int:
    raw = book.org_value("default_period", str(config.DEFAULT_PERIOD_MONTHS))
    try:
        return int(raw or config.DEFAULT_PERIOD_MONTHS)
    except ValueError:
        return config.DEFAULT_PERIOD_MONTHS



def workers(book: Book, *, only_active: bool = True,
            include_external: bool = False) -> list[dict[str, Any]]:
    """Список работников.

    `only_active` — только действующие (работает, посетитель, подрядчик).
    `include_external` — добавить внешних (посетители, подрядчики); по умолчанию
    они исключаются, потому что им планируется только вводный инструктаж.
    """
    rows = list(book.table("Работники").rows)
    if only_active:
        rows = [r for r in rows if worker_is_active(r)]
    if not include_external:
        rows = [r for r in rows if str(r.get("status") or "") not in config.EXTERNAL_WORKER_STATUSES]
    return sorted(rows, key=lambda r: str(r.get("fio") or "").lower())


def worker_is_active(row: dict[str, Any]) -> bool:
    status = str(row.get("status") or "Работает")
    return status in config.ACTIVE_WORKER_STATUSES


def worker_is_external(row: dict[str, Any]) -> bool:
    return str(row.get("status") or "") in config.EXTERNAL_WORKER_STATUSES


def allowed_types_for_worker(worker: dict[str, Any]) -> list[str]:
    """Какие виды инструктажа допустимы для работника.

    Внешним (посетители, подрядчики) проводится только вводный инструктаж.
    """
    if worker_is_external(worker):
        return ["Вводный"]
    return list(config.TRAINING_TYPES)


def worker_by_id(book: Book, worker_id: str) -> dict[str, Any] | None:
    return book.table("Работники").by_key("id", worker_id)


def instructors(book: Book, *, on: date | None = None,
                include_revoked: bool = False) -> list[dict[str, Any]]:
    """Допуски инструктирующих.

    По умолчанию — только действующие: не аннулированные, со сроком не истёкшим
    и у которых работник не уволен.
    """
    rows = list(book.table("Инструктирующие").rows)
    if not include_revoked:
        rows = [r for r in rows if instructor_is_active(book, r, on=on)]
    return sorted(rows, key=lambda r: (str(r.get("fio") or "").lower(),
                                       str(r.get("program") or "").lower()))


def instructor_is_active(book: Book, row: dict[str, Any],
                         *, on: date | None = None) -> bool:
    """Действует ли допуск: не аннулирован, срок не истёк, человек не уволен."""
    if str(row.get("status") or "Действует") != "Действует":
        return False
    check_date = on or config.today()
    if not deadlines.instructor_valid(row.get("valid_to"), check_date):
        return False
    # Допуск уволенного работника не действует: он не может проводить инструктаж.
    person_id = str(row.get("person_id") or "")
    if person_id:
        for worker in book.table("Работники").rows:
            if str(worker.get("id") or "") == person_id:
                return worker_is_active(worker)
    return True


def instructor_revoke_reasons(book: Book) -> list[str]:
    values = [
        str(r.get("value") or "")
        for r in book.table("Справочники").rows
        if str(r.get("kind") or "") == "revoke" and str(r.get("value") or "").strip()
    ]
    return values or ["Увольнение", "Истёк срок обучения", "Перевод на другую должность",
                      "Ошибка при вводе"]


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
    if not worker_is_active(worker):
        raise ValidationError(
            "Нельзя планировать инструктаж: работник в статусе "
            f"«{worker.get('status') or 'Уволен'}»."
        )
    allowed = allowed_types_for_worker(worker)
    if clean["type"] not in allowed:
        raise ValidationError(
            f"Для статуса «{worker.get('status')}» допустим только "
            f"{', '.join(t.lower() for t in allowed)} инструктаж. "
            "Внешним работникам планируется только вводный инструктаж.",
            fields={"type": "Выберите вводный инструктаж."},
        )
    if not worker.get("birth_date"):
        raise ValidationError(
            "В карточке работника не указана дата рождения. Это обязательный реквизит "
            "регистрации инструктажа (пп. 86, 87 Правил № 2464).",
            fields={"worker_id": "Дополните дату рождения в карточке работника."},
        )

    trainer = instructor_by_id(book, clean["instructor_id"])
    if trainer is None:
        raise NotFoundError("Инструктирующий не найден.")
    if clean["type"] not in instructor_types(trainer):
        raise ValidationError(
            f"У инструктирующего не указан допуск по виду «{clean['type']}».",
            fields={"instructor_id": "Выберите инструктирующего с нужным допуском."},
        )
    if str(trainer.get("status") or "Действует") != "Действует":
        raise ValidationError(
            "Допуск инструктирующего аннулирован.",
            fields={"instructor_id": "Выберите инструктирующего с действующим допуском."},
        )
    if not deadlines.instructor_valid(trainer.get("valid_to"), clean["plan_date"]):
        raise ValidationError(
            "Срок действия допуска инструктирующего истекает раньше плановой даты.",
            fields={"instructor_id": "Выберите инструктирующего с действующим допуском."},
        )

    if clean["type"] == "Повторный":
        # Проверяем фактическую периодичность: индивидуальную или организационную.
        effective = _effective_period(book, worker)
        rule = rules.period_rule("Повторный", organization_period=effective)
        if effective > rules.REPEAT_PERIOD_MONTHS and not str(data.get("period_basis") or "").strip():
            raise ValidationError(
                f"Указана периодичность повторного инструктажа {effective} мес. "
                "Правила № 2464 устанавливают срок не реже одного раза в 6 месяцев, "
                "поэтому для большего срока нужно основание.",
                fields={"period_months": "Укажите основание для периодичности больше "
                                         "6 месяцев или измените срок."},
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
            "birth_date": worker.get("birth_date"),
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


# --------------------------------------------------------------------------- #
# Инструкции и программы по охране труда
# --------------------------------------------------------------------------- #

def programs(book: Book, *, only_active: bool = True) -> list[dict[str, Any]]:
    """Список инструкций и программ.

    По умолчанию — только действующие: недействующие скрыты из подсказок,
    но остаются в книге для истории (проведённые инструктажи на них ссылаются).
    """
    rows = list(book.table("Инструкции").rows)
    if only_active:
        rows = [r for r in rows if str(r.get("status") or "Действует") == "Действует"]
    return sorted(rows, key=lambda r: (str(r.get("number") or ""), str(r.get("title") or "").lower()))


def program_title(row: dict[str, Any]) -> str:
    """Название для подсказки и печати: «ИОТ-001 — Инструкция по охране труда»."""
    number = str(row.get("number") or "").strip()
    title = str(row.get("title") or "").strip()
    if number and title:
        return f"{number} — {title}"
    return title or number


def add_program(book: Book, data: dict[str, Any], *, reason: str = "") -> dict[str, Any]:
    """Добавить инструкцию или программу."""
    title = str(data.get("title") or "").strip()
    if not title:
        raise ValidationError("Укажите название инструкции или программы.",
                              fields={"title": "Укажите название."})
    number = str(data.get("number") or "").strip()
    status = str(data.get("status") or "Действует").strip()
    if status not in config.INSTRUCTOR_STATUSES and status not in ("Действует", "Не действует"):
        status = "Действует"
    row = book.table("Инструкции").add(
        {
            "id": new_id(),
            "number": number,
            "title": title,
            "description": str(data.get("description") or "").strip(),
            "approved_by": str(data.get("approved_by") or "").strip(),
            "approved_at": data.get("approved_at") or None,
            "status": status,
            "replaced_by": str(data.get("replaced_by") or "").strip(),
            "note": str(data.get("note") or "").strip(),
        }
    )
    save_book(book, action="program_add", entity="Инструкции", entity_id=row["id"],
              reason=reason or "Добавлена инструкция/программа",
              details={"number": number, "title": title})
    return row


def update_program(book: Book, program_id: str, data: dict[str, Any], *,
                   reason: str = "") -> dict[str, Any]:
    row = book.table("Инструкции").by_key("id", program_id)
    if row is None:
        raise NotFoundError("Инструкция не найдена.")
    title = str(data.get("title") or "").strip()
    if not title:
        raise ValidationError("Укажите название инструкции или программы.",
                              fields={"title": "Укажите название."})
    before = {k: str(row.get(k) or "") for k in ("number", "title", "status", "approved_by")}
    row["number"] = str(data.get("number") or "").strip()
    row["title"] = title
    row["description"] = str(data.get("description") or "").strip()
    row["approved_by"] = str(data.get("approved_by") or "").strip()
    row["approved_at"] = data.get("approved_at") or None
    row["status"] = str(data.get("status") or "Действует").strip()
    row["replaced_by"] = str(data.get("replaced_by") or "").strip()
    row["note"] = str(data.get("note") or "").strip()
    after = {k: str(row.get(k) or "") for k in before}
    changed = {k: {"before": before[k], "after": after[k]} for k in before if before[k] != after[k]}
    save_book(book, action="program_update", entity="Инструкции", entity_id=program_id,
              reason=reason or "Изменена инструкция/программа", details={"changed": changed})
    return row


def deactivate_program(book: Book, program_id: str, *, reason: str) -> dict[str, Any]:
    """Перевести инструкцию в состояние «Не действует»."""
    row = book.table("Инструкции").by_key("id", program_id)
    if row is None:
        raise NotFoundError("Инструкция не найдена.")
    if not str(reason or "").strip():
        raise ValidationError("Укажите причину: почему инструкция больше не действует.",
                              fields={"reason": "Укажите причину."})
    row["status"] = "Не действует"
    row["note"] = (str(row.get("note") or "") + f" | Не действует: {reason}").strip(" |")
    save_book(book, action="program_deactivate", entity="Инструкции",
              entity_id=program_id, reason=reason)
    return row


# --------------------------------------------------------------------------- #
# Аннулирование допуска инструктирующего
# --------------------------------------------------------------------------- #

def revoke_instructor(book: Book, row_id: str, *, reason: str,
                      revoke_date: date | None = None) -> dict[str, Any]:
    """Аннулировать выданный допуск.

    Запись не удаляется: проведённые инструктажи ссылаются на неё. После
    аннулирования допуск не предлагается при планировании, но остаётся
    в истории и виден при включённом фильтре.
    """
    row = book.table("Инструктирующие").by_key("id", row_id)
    if row is None:
        raise NotFoundError("Допуск не найден.")
    if str(row.get("status") or "Действует") == "Аннулирован":
        raise ConflictError("Этот допуск уже аннулирован.")
    if not str(reason or "").strip():
        raise ValidationError("Укажите причину аннулирования допуска.",
                              fields={"reason": "Укажите причину."})
    row["status"] = "Аннулирован"
    row["revoked_at"] = revoke_date or config.today()
    row["revoke_reason"] = str(reason).strip()
    save_book(book, action="instructor_revoke", entity="Инструктирующие",
              entity_id=row_id, reason=reason,
              details={"fio": row.get("fio"), "program": row.get("program")})
    return row


def restore_instructor(book: Book, row_id: str, *, reason: str = "") -> dict[str, Any]:
    """Вернуть аннулированный допуск в действие (если срок не истёк)."""
    row = book.table("Инструктирующие").by_key("id", row_id)
    if row is None:
        raise NotFoundError("Допуск не найден.")
    if str(row.get("status") or "Действует") == "Действует":
        raise ConflictError("Этот допуск уже действует.")
    valid_to = row.get("valid_to")
    if valid_to and valid_to < config.today():
        raise ValidationError(
            "Срок действия допуска уже истёк — восстановить нельзя. "
            "Добавьте новый допуск с актуальной датой обучения.",
            fields={"valid_to": "Срок истёк, требуется новый допуск."},
        )
    row["status"] = "Действует"
    row["revoked_at"] = None
    row["revoke_reason"] = ""
    save_book(book, action="instructor_restore", entity="Инструктирующие",
              entity_id=row_id, reason=reason or "Восстановление допуска")
    return row


def set_deputy(book: Book, row_id: str, deputy_id: str, *, reason: str = "") -> dict[str, Any]:
    """Назначить заместителя для допуска.

    Заместитель — другая строка допуска с тем же набором видов и программой.
    Нужен, чтобы при увольнении или аннулировании основного допуска
    было видно, кто проводит инструктаж вместо него.
    """
    row = book.table("Инструктирующие").by_key("id", row_id)
    if row is None:
        raise NotFoundError("Допуск не найден.")
    if deputy_id and deputy_id == row_id:
        raise ValidationError("Допуск не может быть заместителем сам себе.",
                              fields={"deputy_id": "Выберите другого инструктирующего."})
    if deputy_id:
        deputy = book.table("Инструктирующие").by_key("id", deputy_id)
        if deputy is None:
            raise NotFoundError("Заместитель не найден.")
        if str(deputy.get("status") or "Действует") != "Действует":
            raise ValidationError(
                "Нельзя назначить заместителем аннулированный допуск.",
                fields={"deputy_id": "Выберите действующий допуск."},
            )
    row["deputy_id"] = deputy_id
    save_book(book, action="instructor_deputy", entity="Инструктирующие",
              entity_id=row_id, reason=reason or "Назначен заместитель",
              details={"deputy_id": deputy_id})
    return row


def instructors_without_deputy(book: Book) -> list[dict[str, Any]]:
    """Действующие допуски без заместителя: предупреждение о риске."""
    result = []
    for row in instructors(book):
        deputy_id = str(row.get("deputy_id") or "")
        if deputy_id:
            deputy = book.table("Инструктирующие").by_key("id", deputy_id)
            if deputy and instructor_is_active(book, deputy):
                continue
        result.append(row)
    return result
