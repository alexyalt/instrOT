"""Описание листов и колонок книги Excel.

Схема — единый источник истины для чтения, записи и валидации книги.
Порядок колонок здесь определяет порядок в файле; менять его можно только
с миграцией (см. docs/WORKBOOK.md).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

ColType = Literal["text", "date", "int", "money"]


@dataclass(frozen=True)
class Column:
    key: str
    title: str
    type: ColType = "text"
    width: int = 18
    required: bool = False
    note: str = ""


@dataclass(frozen=True)
class Sheet:
    name: str
    columns: tuple[Column, ...]
    key: str | None = None  # колонка-идентификатор строки (UUID)
    note: str = ""
    editable_in_excel: bool = True


ORGANIZATION_COLUMNS = (
    Column("key", "Параметр", width=34),
    Column("value", "Значение", width=60),
)

WORKERS_COLUMNS = (
    Column("id", "ID", width=38, note="Постоянный идентификатор, не менять"),
    Column("personnel_no", "Табельный", width=12),
    Column("fio", "ФИО", width=34, required=True),
    Column("birth_date", "Дата рождения", "date", 14),
    Column("position", "Должность", width=26, required=True),
    Column("department", "Подразделение", width=24),
    Column("manager", "Руководитель", width=26),
    Column("hire_date", "Дата приёма", "date", 14),
    Column("dismiss_date", "Дата увольнения", "date", 14),
    Column("status", "Статус", width=12, required=True),
    Column("period_months", "Периодичность, мес", "int", 16),
    Column("note", "Примечание", width=28),
    Column("updated_at", "Изменено", width=20, note="Служебное поле"),
)

INSTRUCTORS_COLUMNS = (
    Column("id", "ID полномочия", width=38, note="Строка допуска, не менять"),
    Column("person_id", "ID человека", width=38, note="Служебное поле"),
    Column("fio", "ФИО", width=32, required=True),
    Column("position", "Должность", width=26),
    Column("program", "Программа / основание", width=32),
    Column("types", "Разрешённые виды", width=30, note="Через запятую"),
    Column("valid_from", "Действует с", "date", 14),
    Column("valid_to", "Действует по", "date", 14),
    Column("note", "Примечание", width=26),
)

TRAININGS_COLUMNS = (
    Column("id", "ID", width=38, note="Постоянный идентификатор, не менять"),
    Column("code", "Код бланка", width=12, note="6 цифр; 000000 — ручной бланк"),
    Column("worker_id", "ID работника", width=38),
    Column("fio", "ФИО", width=30, note="Снимок на момент инструктажа"),
    Column("position", "Должность", width=24, note="Снимок"),
    Column("department", "Подразделение", width=22, note="Снимок"),
    Column("type", "Вид", width=14, required=True),
    Column("reason", "Причина", width=26),
    Column("program", "Программа / инструкции", width=28),
    Column("plan_date", "Дата плана", "date", 14),
    Column("fact_date", "Дата факта", "date", 14),
    Column("status", "Статус", width=14, required=True),
    Column("instructor_id", "ID инструктирующего", width=38),
    Column("instructor_fio", "Инструктирующий", width=28),
    Column("source", "Источник", width=14),
    Column("paper_confirmed", "Бумага подтверждена", width=18),
    Column("reg_date", "Дата внесения", "date", 14),
    Column("number_ot", "Номер ОТ", width=16),
    Column("number_type", "Номер по виду", width=16),
    Column("case_id", "Дело", width=16),
    Column("pack_id", "Пакет", width=16),
    Column("note", "Примечание", width=26),
    Column("created_at", "Создано", width=20, note="Служебное поле"),
    Column("updated_at", "Изменено", width=20, note="Служебное поле"),
)

ARCHIVE_COLUMNS = (
    Column("case_id", "Дело", width=12),
    Column("year", "Год", "int", 8),
    Column("type", "Вид", width=14),
    Column("pack_id", "Пакет", width=12),
    Column("title", "Название", width=34),
    Column("opened_at", "Открыт", "date", 14),
    Column("closed_at", "Закрыт", "date", 14),
    Column("status", "Статус", width=12, note="Открыт / Закрыт"),
    Column("note", "Примечание", width=26),
)

DICTIONARY_COLUMNS = (
    Column("kind", "Справочник", width=20, required=True),
    Column("value", "Значение", width=44, required=True),
    Column("note", "Примечание", width=30),
)

WORKER_HISTORY_COLUMNS = (
    Column("id", "ID записи", width=38),
    Column("worker_id", "ID работника", width=38),
    Column("event_date", "Дата события", "date", 14),
    Column("event_type", "Событие", width=18),
    Column("old_value", "Было", width=30),
    Column("new_value", "Стало", width=30),
    Column("reason", "Основание", width=28),
    Column("created_at", "Создано", width=20),
)

INSTRUCTION_COLUMNS = (
    Column("line", "Инструкция", width=100),
)

SHEETS: tuple[Sheet, ...] = (
    Sheet("Организация", ORGANIZATION_COLUMNS, note="Реквизиты и счётчики", editable_in_excel=False),
    Sheet("Работники", WORKERS_COLUMNS, key="id"),
    Sheet("Инструктирующие", INSTRUCTORS_COLUMNS, key="id"),
    Sheet("Инструктажи", TRAININGS_COLUMNS, key="id"),
    Sheet("Архив", ARCHIVE_COLUMNS, note="Строка — один пакет"),
    Sheet("Справочники", DICTIONARY_COLUMNS),
    Sheet("История_работников", WORKER_HISTORY_COLUMNS, key="id"),
    Sheet("Инструкция", INSTRUCTION_COLUMNS, editable_in_excel=False),
)

SHEET_BY_NAME: dict[str, Sheet] = {s.name: s for s in SHEETS}

# Служебные ключи листа «Организация».
ORG_KEY = {
    "org_uuid": "UUID организации",
    "book_uuid": "UUID книги",
    "schema_version": "Версия схемы",
    "name": "Название",
    "full_name": "Полное наименование",
    "inn": "ИНН",
    "kpp": "КПП",
    "ogrn": "ОГРН",
    "address": "Адрес",
    "director": "Руководитель",
    "ot_specialist": "Специалист по охране труда",
    "soon_days": "Порог ближайших, дней",
    "upcoming_days": "Порог предстоящих, дней",
    "default_period": "Периодичность по умолчанию, мес",
    "custom_forms": "Индивидуальные формы",
    "created_at": "Создана",
    "updated_at": "Обновлена",
    "counter_ot": "Счётчик ОТ",
    "counter_type": "Счётчики по видам",
    "counter_case": "Счётчик дел",
    "counter_pack": "Счётчик пакетов",
}

DICTIONARY_KINDS = {
    "reason": "Причины",
    "position": "Должности",
    "department": "Подразделения",
    "program": "Программы",
    "note": "Прочее",
}
