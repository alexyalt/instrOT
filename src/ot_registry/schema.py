"""Описание листов, колонок и формул книги Excel.

Схема — единый источник истины для чтения, записи и валидации книги.
Порядок колонок здесь определяет порядок в файле; менять его можно только
с миграцией (см. docs/WORKBOOK.md).

Листы делятся на виды:
- data   — таблицы, которые читает и пишет приложение;
- org    — служебные параметры организации (ключ-значение);
- summary — формулы и сводки (приложение не меняет, только создаёт);
- info    — инструкция для человека.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

ColType = Literal["text", "date", "int", "money"]
SheetKind = Literal["data", "org", "summary", "info"]


@dataclass(frozen=True)
class Column:
    key: str
    title: str
    type: ColType = "text"
    width: int = 18
    required: bool = False
    note: str = ""


@dataclass(frozen=True)
class FormulaColumn:
    """Колонка с формулой Excel. Приложение её перезаписывает при сохранении.

    Формула получает подстановки: {row} — номер строки, {plan} — буква колонки
    плановой даты, {status} — буква колонки статуса, {fact} — буква колонки
    фактической даты.
    """

    title: str
    template: str
    width: int = 16
    number_format: str = "General"
    note: str = ""


@dataclass(frozen=True)
class Sheet:
    name: str
    columns: tuple[Column, ...] = ()
    key: str | None = None  # колонка-идентификатор строки (UUID)
    note: str = ""
    kind: SheetKind = "data"
    formulas: tuple[FormulaColumn, ...] = ()
    validation: dict[str, tuple[str, ...]] = field(default_factory=dict)

    @property
    def is_data(self) -> bool:
        return self.kind == "data"


ORGANIZATION_COLUMNS = (
    Column("key", "Параметр", width=34),
    Column("value", "Значение", width=60),
)

WORKERS_COLUMNS = (
    Column("id", "ID", width=38, note="Постоянный идентификатор, не менять"),
    Column("personnel_no", "Табельный", width=12),
    Column("fio", "ФИО", width=34, required=True),
    Column("birth_date", "Дата рождения", "date", 14,
           note="Обязательно: реквизит регистрации (п. 86, 87 Правил № 2464)"),
    Column("position", "Должность", width=26, required=True),
    Column("department", "Подразделение", width=24),
    Column("manager", "Руководитель", width=26),
    Column("hire_date", "Дата приёма", "date", 14),
    Column("dismiss_date", "Дата увольнения", "date", 14),
    Column("status", "Статус", width=12, required=True),
    Column("period_months", "Периодичность, мес", "int", 16,
           note="6 мес. по умолчанию; больше — только с основанием"),
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
    Column("status", "Состояние", width=14, note="Действует / Аннулирован"),
    Column("revoked_at", "Аннулирован", "date", 14),
    Column("revoke_reason", "Причина аннулирования", width=28),
    Column("deputy_id", "ID заместителя", width=38,
           note="Строка допуска, к которой привязан заместитель"),
    Column("note", "Примечание", width=26),
)

TRAININGS_COLUMNS = (
    Column("id", "ID", width=38, note="Постоянный идентификатор, не менять"),
    Column("code", "Код бланка", width=12, note="6 цифр; 000000 — ручной бланк"),
    Column("worker_id", "ID работника", width=38),
    Column("fio", "ФИО", width=30, note="Снимок на момент инструктажа"),
    Column("position", "Должность", width=24, note="Снимок"),
    Column("department", "Подразделение", width=22, note="Снимок"),
    Column("birth_date", "Дата рождения", "date", 14,
           note="Обязательный реквизит (п. 86, 87 Правил № 2464)"),
    Column("type", "Вид", width=14, required=True),
    Column("reason", "Причина", width=26),
    Column("program", "ЛНА / программа", width=28),
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

TRAININGS_FORMULAS = (
    FormulaColumn(
        "Дней до срока",
        '=IF(${status}{row}="Запланирован",IF(${plan}{row}="","",${plan}{row}-TODAY()),"")',
        width=14,
        number_format="0",
        note="Считается Excel; приложение считает сроки независимо",
    ),
    FormulaColumn(
        "Состояние",
        '=IF(${status}{row}<>"Запланирован",${status}{row},'
        'IF(${plan}{row}="","нет даты",'
        'IF(${plan}{row}<TODAY(),"ПРОСРОЧЕНО",'
        'IF(${plan}{row}-TODAY()<=7,"ближайшие",'
        'IF(${plan}{row}-TODAY()<=14,"скоро","позднее")))))',
        width=16,
        note="Текстовая подсказка, дублирует подсветку",
    ),
)

WORKERS_FORMULAS = (
    FormulaColumn(
        "Инструктажей проведено",
        '=COUNTIFS(Инструктажи!$C$2:$C$5000,$A{row},Инструктажи!$M$2:$M$5000,"Проведён")',
        width=16,
        number_format="0",
    ),
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

# Инструкции и программы по охране труда: отдельный лист, потому что у записи
# есть номер, дата и приказ утверждения, а также состояние действия.
PROGRAM_COLUMNS = (
    Column("id", "ID", width=38, note="Постоянный идентификатор, не менять"),
    Column("number", "Номер", width=14, note="например: ИОТ-001"),
    Column("title", "Название", width=52, required=True),
    Column("description", "Краткое описание", width=40),
    Column("approved_by", "Утверждено приказом", width=30),
    Column("approved_at", "Дата утверждения", "date", 16),
    Column("status", "Состояние", width=16, note="Действует / Не действует"),
    Column("replaced_by", "Заменено на", width=24),
    Column("note", "Примечание", width=28),
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
    Column("line", "Инструкция", width=110),
)

SUMMARY_COLUMNS = (
    Column("label", "Показатель", width=42),
    Column("value", "Значение", width=18),
    Column("note", "Пояснение", width=60),
)

SHEETS: tuple[Sheet, ...] = (
    Sheet("Организация", ORGANIZATION_COLUMNS, kind="org",
          note="Реквизиты и счётчики. Приложение поддерживает значения автоматически"),
    Sheet("Работники", WORKERS_COLUMNS, key="id", formulas=WORKERS_FORMULAS,
          validation={"status": ("Работает", "Уволен")}),
    Sheet("Инструктирующие", INSTRUCTORS_COLUMNS, key="id"),
    Sheet("Инструктажи", TRAININGS_COLUMNS, key="id", formulas=TRAININGS_FORMULAS,
          validation={"status": ("Запланирован", "Проведён", "Отменён")}),
    Sheet("Сводка", SUMMARY_COLUMNS, kind="summary",
          note="Формулы Excel: считается при открытии книги"),
    Sheet("Архив", ARCHIVE_COLUMNS, note="Строка — один пакет",
          validation={"status": ("Открыт", "Закрыт")}),
    Sheet("Инструкции", PROGRAM_COLUMNS, key="id",
          validation={"status": ("Действует", "Не действует")}),
    Sheet("Справочники", DICTIONARY_COLUMNS),
    Sheet("История_работников", WORKER_HISTORY_COLUMNS, key="id"),
    Sheet("Инструкция", INSTRUCTION_COLUMNS, kind="info"),
)

SHEET_BY_NAME: dict[str, Sheet] = {s.name: s for s in SHEETS}
DATA_SHEETS: tuple[Sheet, ...] = tuple(s for s in SHEETS if s.is_data)
READABLE_SHEETS: tuple[Sheet, ...] = tuple(s for s in SHEETS if s.kind in ("data", "org"))
ORG_SHEET = "Организация"
SUMMARY_SHEET = "Сводка"
INSTRUCTION_SHEET = "Инструкция"
ARCHIVE_SHEET = "Архив"

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
    "brand_color": "Цвет фирмы",
    "nav_order": "Порядок закладок",
    "created_at": "Создана",
    "updated_at": "Обновлена",
    "counter_ot": "Счётчик ОТ",
    "counter_type": "Счётчики по видам",
    "counter_case": "Счётчик дел",
    "counter_pack": "Счётчик пакетов",
}

DICTIONARY_KINDS = {
    "reason": "Причины",
    "revoke": "Причины аннулирования",
    "position": "Должности",
    "department": "Подразделения",
    "program": "Программы",
    "note": "Прочее",
}

# Базовое наполнение справочников при создании книги.
DEFAULT_DICTIONARIES: tuple[tuple[str, str, str], ...] = (
    ("reason", "Изменение технологического процесса", ""),
    ("reason", "Нарушение требований охраны труда", ""),
    ("reason", "Перерыв в работе более 30 календарных дней", ""),
    ("reason", "По требованию должностных лиц органов надзора", ""),
    ("reason", "Изменение условий труда", ""),
    ("reason", "Перед выполнением разовых работ", ""),
    ("position", "Дворник", ""),
    ("position", "Уборщик служебных помещений", ""),
    ("position", "Озеленитель", ""),
    ("position", "Садовник", ""),
    ("position", "Рабочий зелёного хозяйства", ""),
    ("position", "Мастер участка", ""),
    ("position", "Начальник участка", ""),
    ("position", "Специалист по охране труда", ""),
    ("department", "Администрация", ""),
    ("department", "Участок благоустройства", ""),
    ("department", "Участок озеленения", ""),
    ("department", "Клининговый участок", ""),
    ("department", "Другое", "внешние и временные работники"),
    ("program", "Вводный инструктаж по охране труда", ""),
    ("program", "Инструкция по охране труда № 1 (дворник)", ""),
    ("program", "Инструкция по охране труда № 2 (озеленитель)", ""),
    ("program", "Инструкция по охране труда № 3 (уборщик помещений)", ""),
    ("program", "Инструкция по охране труда при работе с ручным инструментом", ""),
)


def column_letter(index: int) -> str:
    """1 → A, 26 → Z, 27 → AA."""
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def column_letters(sheet: Sheet) -> dict[str, str]:
    """Соответствие ключ колонки → буква столбца (только для колонок данных)."""
    return {col.key: column_letter(idx) for idx, col in enumerate(sheet.columns, start=1)}


def formula_context(sheet: Sheet) -> dict[str, str]:
    """Подстановки для формул: буквы ключевых колонок."""
    letters = column_letters(sheet)
    return {
        "plan": letters.get("plan_date", "K"),
        "status": letters.get("status", "M"),
        "fact": letters.get("fact_date", "L"),
        "code": letters.get("code", "B"),
        "type": letters.get("type", "H"),
    }


def last_letter(sheet: Sheet) -> str:
    return column_letter(len(sheet.columns))


def formula_letter(sheet: Sheet, position: int) -> str:
    """Буква для формулы с учётом того, что они идут после колонок данных."""
    return column_letter(len(sheet.columns) + position)
