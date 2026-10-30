"""Демонстрационные данные для новой книги.

Нужны, чтобы структура книги была видна сразу: как заполняются строки,
как считаются сроки, как выглядят формулы. Все данные синтетические:
реальных работников они не описывают.

Демоданные включаются при создании организации флажком «Добавить примеры
данных». Для рабочей организации их можно не добавлять или удалить.
"""

from __future__ import annotations

import random
from datetime import date, timedelta
from typing import Any

from .config import TRAINING_TYPES
from .excel_store import Book, new_id

# Синтетические ФИО: типичные для отрасли, но не привязанные к людям.
DEMO_PERSONS: tuple[tuple[str, str, str, date], ...] = (
    ("Иванов Иван Иванович", "Дворник", "Участок благоустройства", date(1978, 3, 14)),
    ("Петрова Анна Сергеевна", "Уборщик служебных помещений", "Клининговый участок", date(1985, 7, 22)),
    ("Сидоров Пётр Алексеевич", "Озеленитель", "Участок озеленения", date(1990, 11, 5)),
    ("Кузнецова Мария Викторовна", "Садовник", "Участок озеленения", date(1982, 1, 30)),
    ("Смирнов Николай Петрович", "Мастер участка", "Участок благоустройства", date(1975, 9, 18)),
    ("Волкова Елена Дмитриевна", "Уборщик служебных помещений", "Клининговый участок", date(1993, 4, 8)),
    ("Морозов Сергей Владимирович", "Рабочий зелёного хозяйства", "Участок озеленения", date(1988, 6, 25)),
    ("Фёдорова Ольга Игоревна", "Дворник", "Участок благоустройства", date(1996, 12, 2)),
    ("Никитин Андрей Юрьевич", "Озеленитель", "Участок озеленения", date(1979, 8, 11)),
    ("Егорова Татьяна Николаевна", "Уборщик служебных помещений", "Клининговый участок", date(1987, 2, 17)),
)

DEMO_INSTRUCTORS: tuple[tuple[str, str, str, str, date], ...] = (
    (
        "Соколов Дмитрий Александрович",
        "Специалист по охране труда",
        "Вводный инструктаж по охране труда",
        "Вводный,Первичный,Повторный,Внеплановый,Целевой",
        date(2029, 12, 31),
    ),
    (
        "Соколов Дмитрий Александрович",
        "Специалист по охране труда",
        "Инструкция по охране труда при работе с ручным инструментом",
        "Первичный,Повторный,Внеплановый",
        date(2029, 12, 31),
    ),
    (
        "Смирнов Николай Петрович",
        "Мастер участка",
        "Инструкция по охране труда № 1 (дворник)",
        "Первичный,Повторный,Целевой",
        date(2028, 6, 30),
    ),
    (
        "Кузнецова Мария Викторовна",
        "Садовник",
        "Инструкция по охране труда № 2 (озеленитель)",
        "Первичный,Повторный",
        date(2027, 10, 15),
    ),
)

REASONS = (
    "Изменение технологического процесса",
    "Нарушение требований охраны труда",
    "Перерыв в работе более 30 календарных дней",
    "По требованию должностных лиц органов надзора",
)


def _phone() -> str:
    return ""


def fill_demo_data(book: Book, *, seed: int = 20260101) -> dict[str, int]:
    """Заполнить книгу синтетическими данными.

    Возвращает количество созданных строк по листам. Данные согласованы:
    у каждого работника есть дата рождения, у проведённых инструктажей —
    инструктирующий с действующим допуском, номера уникальны.
    """
    rng = random.Random(seed)
    today = date.today()

    workers_table = book.tables["Работники"]
    instructors_table = book.tables["Инструктирующие"]
    trainings_table = book.tables["Инструктажи"]
    archive_table = book.tables["Архив"]
    history_table = book.tables["История_работников"]

    # --- Инструктирующие и допуски ---
    instructor_rows: list[dict[str, Any]] = []
    for idx, (fio, position, program, types, valid_to) in enumerate(DEMO_INSTRUCTORS):
        person_id = f"demo-person-{idx // 2 + 1}"
        row = instructors_table.add(
            {
                "id": new_id(),
                "person_id": person_id,
                "fio": fio,
                "position": position,
                "program": program,
                "types": types,
                "valid_from": date(2025, 1, 15),
                "valid_to": valid_to,
                "note": "демонстрационные данные",
            }
        )
        instructor_rows.append(row)

    # --- Работники ---
    worker_rows: list[dict[str, Any]] = []
    for idx, (fio, position, department, birth) in enumerate(DEMO_PERSONS):
        hire = date(today.year - rng.randint(0, 6), rng.randint(1, 12), rng.randint(1, 27))
        status = "Работает"
        dismiss = None
        if idx == len(DEMO_PERSONS) - 1:
            status = "Уволен"
            dismiss = today - timedelta(days=rng.randint(20, 120))
        row = workers_table.add(
            {
                "id": new_id(),
                "personnel_no": f"{idx + 1:04d}",
                "fio": fio,
                "birth_date": birth,
                "position": position,
                "department": department,
                "manager": "Смирнов Николай Петрович" if department != "Клининговый участок"
                           else "Петрова Анна Сергеевна",
                "hire_date": hire,
                "dismiss_date": dismiss,
                "status": status,
                "period_months": 6,
                "note": "демонстрационные данные",
                "updated_at": date.today().strftime("%d.%m.%Y"),
            }
        )
        worker_rows.append(row)

    # --- Планы и проведённые инструктажи ---
    active_workers = [r for r in worker_rows if r["status"] == "Работает"]
    active_instructors = [r for r in instructor_rows if r["valid_to"] and r["valid_to"] > today]

    ot_counter = 0
    type_counters: dict[str, int] = {}
    case_counter = 0
    pack_counter = 0
    codes: set[str] = set()

    def next_code() -> str:
        while True:
            code = str(rng.randint(100000, 999999))
            if code not in codes:
                codes.add(code)
                return code

    def case_for(training_type: str, year: int) -> tuple[str, str]:
        nonlocal case_counter, pack_counter
        for row in archive_table.rows:
            if row.get("type") == training_type and row.get("year") == year:
                case_id = str(row.get("case_id") or "")
                for pack in archive_table.rows:
                    if pack.get("case_id") == case_id and pack.get("status") == "Открыт":
                        return case_id, str(pack.get("pack_id") or "")
                pack_counter += 1
                pack_id = f"П{pack_counter:04d}"
                archive_table.add(
                    {
                        "case_id": case_id,
                        "year": year,
                        "type": training_type,
                        "pack_id": pack_id,
                        "title": f"Пакет {pack_id}",
                        "opened_at": today,
                        "closed_at": None,
                        "status": "Открыт",
                        "note": "демонстрационные данные",
                    }
                )
                return case_id, pack_id

        case_counter += 1
        case_id = f"Д{case_counter:04d}"
        pack_counter += 1
        pack_id = f"П{pack_counter:04d}"
        archive_table.add(
            {
                "case_id": case_id,
                "year": year,
                "type": training_type,
                "pack_id": pack_id,
                "title": f"Пакет {pack_id}",
                "opened_at": today,
                "closed_at": None,
                "status": "Открыт",
                "note": "демонстрационные данные",
            }
        )
        return case_id, pack_id

    def add_training(
        worker: dict[str, Any],
        training_type: str,
        *,
        fact: date | None,
        plan: date | None,
        instructor: dict[str, Any],
        program: str,
        reason: str = "",
        source: str = "Система",
    ) -> dict[str, Any]:
        nonlocal ot_counter
        status = "Проведён" if fact else "Запланирован"
        code_value = next_code()
        numbers: dict[str, Any] = {
            "number_ot": "",
            "number_type": "",
            "case_id": "",
            "pack_id": "",
        }
        reg_date = None
        paper = ""
        if fact:
            ot_counter += 1
            prefix = {"Вводный": "ВВ", "Первичный": "ПР", "Повторный": "ПВ",
                      "Внеплановый": "ВН", "Целевой": "ЦЛ"}.get(training_type, "ИН")
            type_counters[training_type] = type_counters.get(training_type, 0) + 1
            case_id, pack_id = case_for(training_type, fact.year)
            numbers = {
                "number_ot": f"ОТ{ot_counter:06d}",
                "number_type": f"{prefix}{type_counters[training_type]:06d}",
                "case_id": case_id,
                "pack_id": pack_id,
            }
            reg_date = fact + timedelta(days=rng.randint(0, 3))
            paper = "Да"
            if source == "Ручной бланк":
                code_value = "000000"

        return trainings_table.add(
            {
                "id": new_id(),
                "code": code_value,
                "worker_id": worker["id"],
                "fio": worker["fio"],
                "position": worker["position"],
                "department": worker["department"],
                "birth_date": worker["birth_date"],
                "type": training_type,
                "reason": reason,
                "program": program,
                "plan_date": plan,
                "fact_date": fact,
                "status": status,
                "instructor_id": instructor["id"],
                "instructor_fio": instructor["fio"],
                "source": source,
                "paper_confirmed": paper,
                "reg_date": reg_date,
                "note": "демонстрационные данные",
                **numbers,
                "created_at": (plan or fact or today).strftime("%d.%m.%Y"),
                "updated_at": today.strftime("%d.%m.%Y"),
            }
        )

    primary_instructor = active_instructors[0]
    secondary = active_instructors[2] if len(active_instructors) > 2 else active_instructors[0]

    # Вводный и первичный — проведены в прошлом, есть у всех работающих.
    for worker in active_workers:
        intro_date = worker["hire_date"] + timedelta(days=rng.randint(0, 2))
        add_training(
            worker, "Вводный",
            fact=intro_date, plan=intro_date,
            instructor=primary_instructor,
            program="Вводный инструктаж по охране труда",
        )
        add_training(
            worker, "Первичный",
            fact=intro_date, plan=intro_date,
            instructor=secondary,
            program="Инструкция по охране труда № 1 (дворник)",
        )

    # Повторные: разные состояния — просрочено, сегодня, ближайшие, позднее.
    offsets = [-25, -6, 0, 3, 7, 12, 20, 45, 90, 150]
    for worker, offset in zip(active_workers, offsets):
        last_fact = today - timedelta(days=180 + rng.randint(0, 40))
        add_training(
            worker, "Повторный",
            fact=last_fact, plan=last_fact,
            instructor=secondary,
            program="Инструкция по охране труда № 1 (дворник)",
        )
        add_training(
            worker, "Повторный",
            fact=None, plan=today + timedelta(days=offset),
            instructor=secondary,
            program="Инструкция по охране труда № 1 (дворник)",
        )

    # Внеплановый и целевой — примеры с причиной.
    add_training(
        active_workers[0], "Внеплановый",
        fact=today - timedelta(days=40), plan=today - timedelta(days=40),
        instructor=primary_instructor,
        program="Инструкция по охране труда при работе с ручным инструментом",
        reason="Нарушение требований охраны труда",
    )
    add_training(
        active_workers[1], "Целевой",
        fact=None, plan=today + timedelta(days=5),
        instructor=primary_instructor,
        program="Инструкция по охране труда при работе с ручным инструментом",
        reason="Перед выполнением разовых работ",
    )

    # Пример бумажного бланка, внесённого вручную.
    add_training(
        active_workers[2], "Повторный",
        fact=today - timedelta(days=12), plan=None,
        instructor=secondary,
        program="Инструкция по охране труда № 2 (озеленитель)",
        source="Ручной бланк",
    )

    # История: перевод одного работника.
    if len(worker_rows) > 2:
        moved = worker_rows[2]
        moved["department"] = "Участок благоустройства"
        moved["position"] = "Дворник"
        history_table.add(
            {
                "id": new_id(),
                "worker_id": moved["id"],
                "event_date": today - timedelta(days=30),
                "event_type": "Перевод",
                "old_value": "Озеленитель / Участок озеленения",
                "new_value": "Дворник / Участок благоустройства",
                "reason": "приказ № 15 от " + (today - timedelta(days=30)).strftime("%d.%m.%Y"),
                "created_at": today.strftime("%d.%m.%Y %H:%M"),
            }
        )

    # Счётчики организации.
    book.set_org_value("counter_ot", ot_counter)
    book.set_org_value("counter_type", __import__("json").dumps(type_counters, ensure_ascii=False))
    book.set_org_value("counter_case", case_counter)
    book.set_org_value("counter_pack", pack_counter)

    return {
        "Работники": len(worker_rows),
        "Инструктирующие": len(instructor_rows),
        "Инструктажи": len(trainings_table.rows),
        "Архив": len(archive_table.rows),
        "История": len(history_table.rows),
    }
