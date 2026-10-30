"""Нормативные сроки и требования ПП РФ № 2464 (в части инструктажей).

Источник: Постановление Правительства РФ от 24.12.2021 № 2464
«О порядке обучения по охране труда и проверки знания требований охраны труда»
(Правила обучения по охране труда и проверки знания требований охраны труда).

Здесь собраны только те нормы, которые влияют на расчёт сроков и проверку
полноты регистрации. Текст правил меняется: перед внедрением проверьте
актуальную редакцию.

Ключевые пункты:
- п. 86 — реквизиты регистрации вводного инструктажа;
- п. 87 — реквизиты регистрации инструктажа на рабочем месте
  (первичный, повторный, внеплановый) и целевого;
- п. 88 — порядок регистрации и форма документирования утверждаются
  работодателем; именно этот пункт позволяет вести собственные формы
  вместо журнала;
- повторный инструктаж — не реже одного раза в 6 месяцев.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any

from . import config, deadlines

LAW_TITLE = (
    "Постановление Правительства РФ от 24.12.2021 № 2464 "
    "«О порядке обучения по охране труда и проверки знания требований охраны труда»"
)

LAW_REFERENCE = {
    "registration_order": "п. 88 Правил № 2464",
    "intro_fields": "п. 86 Правил № 2464",
    "workplace_fields": "п. 87 Правил № 2464",
    "repeat_period": "п. 87 Правил № 2464 (повторный — не реже 1 раза в 6 месяцев)",
}

# Периодичность повторного инструктажа по умолчанию, месяцы.
REPEAT_PERIOD_MONTHS = 6

# Реквизиты по видам инструктажа. Используются для проверки полноты
# и для подсказок в интерфейсе.
REQUIRED_FIELDS: dict[str, tuple[str, ...]] = {
    "Вводный": (
        "fact_date",
        "fio",
        "position",
        "birth_date",
        "department",
        "instructor_fio",
        "instructor_position",
        "signatures",
    ),
    "Первичный": (
        "fact_date",
        "fio",
        "position",
        "birth_date",
        "type",
        "instructor_fio",
        "instructor_position",
        "lna",
        "signatures",
    ),
    "Повторный": (
        "fact_date",
        "fio",
        "position",
        "birth_date",
        "type",
        "instructor_fio",
        "instructor_position",
        "lna",
        "signatures",
    ),
    "Внеплановый": (
        "fact_date",
        "fio",
        "position",
        "birth_date",
        "type",
        "reason",
        "instructor_fio",
        "instructor_position",
        "lna",
        "signatures",
    ),
    "Целевой": (
        "fact_date",
        "fio",
        "position",
        "birth_date",
        "type",
        "reason",
        "instructor_fio",
        "instructor_position",
        "lna",
        "signatures",
    ),
}

FIELD_TITLES = {
    "fact_date": "дата проведения",
    "fio": "ФИО работника",
    "position": "профессия (должность) работника",
    "birth_date": "число, месяц, год рождения работника",
    "department": "наименование подразделения",
    "type": "вид инструктажа",
    "reason": "причина проведения",
    "instructor_fio": "ФИО проводившего инструктаж",
    "instructor_position": "профессия (должность) проводившего инструктаж",
    "lna": "наименование локального акта (ЛНА)",
    "signatures": "подписи проводившего и прошедшего инструктаж",
}

TYPE_LABEL = {
    "program": "наименование локального акта (ЛНА), в объёме требований которого проведён инструктаж",
    "reason": "причина проведения инструктажа",
}


@dataclass(frozen=True)
class PeriodRule:
    """Правило периодичности для вида инструктажа."""

    training_type: str
    months: int | None
    basis: str
    note: str = ""


def period_rule(training_type: str, *, organization_period: int | None = None) -> PeriodRule:
    """Периодичность для вида инструктажа.

    Повторный — не реже одного раза в 6 месяцев. Для остальных видов
    повторяемость не устанавливается: они проводятся по событию.
    """
    if training_type == "Повторный":
        months = organization_period or REPEAT_PERIOD_MONTHS
        if months > REPEAT_PERIOD_MONTHS:
            return PeriodRule(
                training_type,
                months,
                LAW_REFERENCE["repeat_period"],
                note=(
                    "Периодичность больше 6 месяцев требует основания: правила "
                    "устанавливают предел «не реже одного раза в 6 месяцев»."
                ),
            )
        return PeriodRule(training_type, months, LAW_REFERENCE["repeat_period"])
    return PeriodRule(
        training_type,
        None,
        LAW_REFERENCE["registration_order"],
        note="Проводится по событию, а не по календарному сроку.",
    )


def next_repeat(fact_date: date, *, organization_period: int | None = None) -> date:
    """Предельная дата следующего повторного инструктажа от фактической даты."""
    months = organization_period or REPEAT_PERIOD_MONTHS
    if months > REPEAT_PERIOD_MONTHS:
        months = REPEAT_PERIOD_MONTHS
    return deadlines.next_repeat_date(fact_date, months)


def check_required_fields(row: dict[str, Any]) -> list[str]:
    """Проверить, какие обязательные реквизиты отсутствуют в проведённой записи."""
    training_type = str(row.get("type") or "")
    required = REQUIRED_FIELDS.get(training_type, ())
    missing: list[str] = []
    for field in required:
        if field in ("signatures",):
            if str(row.get("paper_confirmed") or "") != "Да":
                missing.append(FIELD_TITLES[field])
            continue
        if field == "lna":
            if not str(row.get("program") or "").strip():
                missing.append(FIELD_TITLES[field])
            continue
        if field == "instructor_position":
            # Должность хранится в строке инструктирующего, а не в снимке;
            # отсутствие проверяется отдельной сверкой.
            continue
        if not row.get(field):
            missing.append(FIELD_TITLES[field])
    return missing


def period_warning(organization_period: int | None) -> str:
    if organization_period and organization_period > REPEAT_PERIOD_MONTHS:
        return (
            f"Указана периодичность повторного инструктажа {organization_period} мес. "
            f"Правила № 2464 устанавливают срок не реже одного раза в "
            f"{REPEAT_PERIOD_MONTHS} месяцев. Проверьте основание."
        )
    return ""
