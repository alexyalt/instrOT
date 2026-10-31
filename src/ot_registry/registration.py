"""Регистрация проведённых инструктажей: обычная и по бумажному бланку."""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

from . import config, deadlines, rules, validation
from .errors import ConflictError, NotFoundError, ValidationError
from .excel_store import Book, new_id, save_book
from .services import (
    dictionaries,
    instructor_by_id,
    instructor_types,
    issue_numbers,
    plan_training,
    planned,
    trainings,
    worker_by_id,
    add_worker,
)

MANUAL_SOURCE = "Ручной бланк"


def find_active_by_code(book: Book, code: str) -> dict[str, Any]:
    """Найти запланированный инструктаж по коду бланка."""
    normalized = validation.normalize_code(code)
    if validation.is_manual_code(normalized):
        raise ValidationError(
            "Код 000-000 означает бумажный бланк, заполненный вне системы. "
            "Для него используйте раздел «Внести бумажный бланк»."
        )
    matches = [
        t for t in trainings(book)
        if str(t.get("code") or "") == normalized and t.get("status") == "Запланирован"
    ]
    if not matches:
        old = [t for t in trainings(book) if str(t.get("code") or "") == normalized]
        if old:
            raise ConflictError(
                "Код найден, но инструктаж уже проведён или отменён. "
                "Проверьте историю, чтобы не внести документ второй раз.",
                details=[str(t.get("id")) for t in old],
            )
        raise NotFoundError("Запланированный инструктаж с таким кодом не найден.")
    if len(matches) > 1:
        raise ConflictError(
            "Найдено несколько активных записей с этим кодом. Требуется проверка.",
            details=[str(t.get("id")) for t in matches],
        )
    return matches[0]


def _worker_snapshot(book: Book, worker_id: str) -> dict[str, str]:
    worker = worker_by_id(book, worker_id)
    if worker is None:
        raise NotFoundError("Работник не найден.")
    return {
        "fio": str(worker.get("fio") or ""),
        "position": str(worker.get("position") or ""),
        "department": str(worker.get("department") or ""),
        "birth_date": worker.get("birth_date"),
    }


def _check_instructor(book: Book, instructor_id: str, *, on: date,
                      training_type: str) -> dict[str, Any]:
    trainer = instructor_by_id(book, instructor_id)
    if trainer is None:
        raise NotFoundError("Инструктирующий не найден.")
    if str(trainer.get("status") or "Действует") != "Действует":
        raise ValidationError(
            "Допуск этого инструктирующего аннулирован — проводить инструктаж он не может.",
            fields={"instructor_id": "Выберите другого инструктирующего."},
        )
    if training_type not in instructor_types(trainer):
        raise ValidationError(
            "У инструктирующего нет допуска по этому виду инструктажа.",
            fields={"instructor_id": "Выберите другого инструктирующего."},
        )
    if not deadlines.instructor_valid(trainer.get("valid_to"), on):
        raise ValidationError(
            f"На {on:%d.%m.%Y} допуск инструктирующего уже не действует.",
            fields={"instructor_id": "Выберите инструктирующего с действующим допуском."},
        )
    return trainer


def register_training(book: Book, training_id: str, data: dict[str, Any]) -> dict[str, Any]:
    """Зарегистрировать факт проведения по существующей записи."""
    row = book.table("Инструктажи").by_key("id", training_id)
    if row is None:
        raise NotFoundError("Запись инструктажа не найдена.")
    if row.get("status") == "Проведён":
        raise ConflictError(
            "Инструктаж уже зарегистрирован. Повторная регистрация не выполнена."
        )
    if row.get("status") == "Отменён":
        raise ConflictError("Запись отменена. Создайте новый план.")

    clean = validation.validate_registration(data)
    training_type = str(row.get("type") or "")
    validation.require_birth_date(worker_by_id(book, str(row.get("worker_id") or "")))
    if not row.get("birth_date"):
        worker_row = worker_by_id(book, str(row.get("worker_id") or "")) or {}
        row["birth_date"] = worker_row.get("birth_date")
    trainer = _check_instructor(
        book, clean["instructor_id"], on=clean["fact_date"], training_type=training_type
    )

    numbers = issue_numbers(book, training_type, clean["fact_date"])
    row.update(numbers)
    row["status"] = "Проведён"
    row["fact_date"] = clean["fact_date"]
    row["reg_date"] = config.today()
    row["instructor_id"] = clean["instructor_id"]
    row["instructor_fio"] = str(trainer.get("fio") or "")
    row["paper_confirmed"] = "Да"
    row["source"] = row.get("source") or "Система"
    if clean["note"]:
        row["note"] = (str(row.get("note") or "") + " | " + clean["note"]).strip(" |")
    row["updated_at"] = datetime.now().strftime("%d.%m.%Y %H:%M")

    next_plan = _maybe_plan_repeat(book, row)
    save_book(
        book,
        action="training_register",
        entity="Инструктажи",
        entity_id=training_id,
        reason="Регистрация проведённого инструктажа",
        details={
            "code": row.get("code"),
            "numbers": numbers,
            "next_plan_id": next_plan.get("id") if next_plan else "",
        },
    )
    return {"training": row, "next_plan": next_plan}


def _period_for_worker(book: Book, worker_id: str) -> int:
    worker = worker_by_id(book, worker_id) or {}
    period = worker.get("period_months")
    if period:
        return int(period)
    default = book.org_value("default_period", str(config.DEFAULT_PERIOD_MONTHS))
    try:
        return int(default or config.DEFAULT_PERIOD_MONTHS)
    except ValueError:
        return config.DEFAULT_PERIOD_MONTHS


def _maybe_plan_repeat(book: Book, row: dict[str, Any]) -> dict[str, Any] | None:
    """Для повторного инструктажа предложить следующий срок."""
    if str(row.get("type") or "") != "Повторный":
        return None
    fact_date = row.get("fact_date")
    if not fact_date:
        return None
    period = _period_for_worker(book, str(row.get("worker_id") or ""))
    next_date = rules.next_repeat(fact_date, organization_period=period)
    existing = [
        t for t in planned(book)
        if t.get("worker_id") == row.get("worker_id") and t.get("type") == "Повторный"
    ]
    if existing:
        return None
    return plan_training(
        book,
        {
            "worker_id": row.get("worker_id"),
            "instructor_id": row.get("instructor_id"),
            "plan_date": next_date,
            "type": "Повторный",
            "program": row.get("program"),
            "reason": "",
            "allow_past": False,
        },
        repeat_from=row,
    )


def _similar_trainings(book: Book, *, worker_id: str, training_type: str,
                       fact_date: date) -> list[dict[str, Any]]:
    return [
        t for t in trainings(book)
        if t.get("worker_id") == worker_id
        and t.get("type") == training_type
        and t.get("status") == "Проведён"
        and t.get("fact_date") == fact_date
    ]


def register_manual(book: Book, data: dict[str, Any], *,
                    allow_similar: bool = False) -> dict[str, Any]:
    """Внести бумажный бланк, заполненный вне системы.

    Если указан plan_id — завершает существующий план (код плана сохраняется
    в примечании). Иначе создаёт проведённую запись без плановой даты.
    """
    clean = validation.validate_manual_entry(data)
    training_type = clean["type"]
    fact_date = clean["fact_date"]

    # При связывании с планом инструктирующий берётся из плана: специалист
    # не должен вводить его повторно с бумажного бланка.
    plan_row = None
    if clean["plan_id"]:
        plan_row = book.table("Инструктажи").by_key("id", clean["plan_id"])
        if plan_row is None:
            raise NotFoundError("Выбранный план не найден.")
        if plan_row.get("status") != "Запланирован":
            raise ConflictError("Выбранный план уже завершён или отменён.")
        if not clean["instructor_id"]:
            clean["instructor_id"] = str(plan_row.get("instructor_id") or "")

    trainer = _check_instructor(
        book, clean["instructor_id"], on=fact_date, training_type=training_type
    )

    # Работник должен существовать в книге: создание людей из формы
    # бумажного ввода порождало случайные должности и подразделения.
    worker_id = str(clean["worker_id"] or "")

    snapshot = _worker_snapshot(book, worker_id)
    validation.require_birth_date(worker_by_id(book, worker_id))
    similar = _similar_trainings(
        book, worker_id=worker_id, training_type=training_type, fact_date=fact_date
    )
    if similar and not allow_similar:
        raise ConflictError(
            "Похоже, этот инструктаж уже внесён: совпадают работник, вид и дата. "
            "Проверьте существующую запись. Если инструктажей действительно было два, "
            "подтвердите это.",
            details=[str(t.get("id")) for t in similar],
        )

    numbers = issue_numbers(book, training_type, fact_date)
    now = datetime.now().strftime("%d.%m.%Y %H:%M")

    if plan_row is not None:
        row = plan_row
        plan_code = str(row.get("code") or "")
        row["note"] = (str(row.get("note") or "") + f" | Бумажный бланк; код плана {plan_code}").strip(" |")
        row["code"] = config.MANUAL_CODE
    else:
        row = book.table("Инструктажи").add(
            {
                "id": new_id(),
                "code": config.MANUAL_CODE,
                "worker_id": worker_id,
                "fio": snapshot["fio"],
                "position": snapshot["position"],
                "department": snapshot["department"],
                "type": training_type,
                "reason": clean["reason"],
                "program": clean["program"],
                "birth_date": snapshot["birth_date"],
                "plan_date": None,
                "created_at": now,
            }
        )

    row.update(numbers)
    row.update(
        {
            "status": "Проведён",
            "fact_date": fact_date,
            "reg_date": config.today(),
            "instructor_id": clean["instructor_id"],
            "instructor_fio": str(trainer.get("fio") or ""),
            "paper_confirmed": "Да",
            "source": MANUAL_SOURCE,
            "updated_at": now,
        }
    )
    if clean["manual_mark"]:
        row["note"] = (str(row.get("note") or "") + f" | Отметка на бланке: {clean['manual_mark']}").strip(" |")
    if clean["note"]:
        row["note"] = (str(row.get("note") or "") + " | " + clean["note"]).strip(" |")

    next_plan = _maybe_plan_repeat(book, row)
    save_book(
        book,
        action="manual_register",
        entity="Инструктажи",
        entity_id=str(row.get("id")),
        reason="Внесение бумажного бланка",
        details={
            "numbers": numbers,
            "worker_id": worker_id,
            "linked_plan": bool(plan_row),
            "next_plan_id": (next_plan or {}).get("id", ""),
        },
    )
    return {"training": row, "next_plan": next_plan}


def adjust_repeat_plans(book: Book, *, dry_run: bool = True) -> dict[str, Any]:
    """Проверить повторные планы после массового ввода бумажных документов.

    План, оказавшийся раньше последнего уже проведённого повторного, считается
    устаревшим: он не должен оставаться просроченным.
    """
    last_facts: dict[str, date] = {}
    for row in trainings(book):
        if row.get("status") != "Проведён" or row.get("type") != "Повторный":
            continue
        fact = row.get("fact_date")
        if fact is None:
            continue
        worker_id = str(row.get("worker_id") or "")
        if worker_id not in last_facts or fact > last_facts[worker_id]:
            last_facts[worker_id] = fact

    stale: list[dict[str, Any]] = []
    for row in planned(book):
        if row.get("type") != "Повторный":
            continue
        worker_id = str(row.get("worker_id") or "")
        last = last_facts.get(worker_id)
        plan_date = row.get("plan_date")
        if last and plan_date and plan_date <= last:
            stale.append(
                {
                    "id": str(row.get("id")),
                    "fio": str(row.get("fio") or ""),
                    "plan_date": plan_date,
                    "last_fact": last,
                    "period": _period_for_worker(book, worker_id),
                }
            )

    if not dry_run and stale:
        for item in stale:
            row = book.table("Инструктажи").by_key("id", item["id"])
            if row:
                row["status"] = "Отменён"
                row["note"] = (str(row.get("note") or "") + " | Отменён: устарел после внесения более поздних данных").strip(" |")
        save_book(book, action="repeat_plans_adjust", entity="Инструктажи",
                  reason="Отмена устаревших повторных планов",
                  details={"count": len(stale)})

    return {"stale": stale, "dry_run": dry_run}


def suggest_link_candidates(book: Book, *, worker_id: str, training_type: str) -> list[dict[str, Any]]:
    """Планы, к которым может относиться бумажный бланк."""
    rows = [
        t for t in planned(book)
        if t.get("worker_id") == worker_id and t.get("type") == training_type
    ]
    return sorted(rows, key=lambda r: r.get("plan_date") or date.min)


def suggest_new_worker_data(book: Book) -> dict[str, list[str]]:
    """Подсказки для формы добавления работника (должности, подразделения)."""
    return dictionaries(book)
