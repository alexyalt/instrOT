"""Расчёт сроков инструктажей и уровней подсветки.

Интерфейс считает сроки самостоятельно и не полагается на формулы Excel:
openpyxl не пересчитывает формулы, а сохранённый кэш может быть устаревшим.
"""

from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import date, timedelta
from enum import IntEnum

from . import config


class DueLevel(IntEnum):
    """Уровень срочности. Больше — важнее."""

    DONE = 0
    NONE = 1
    LATER = 2
    UPCOMING = 3
    SOON = 4
    TODAY = 5
    OVERDUE = 6

    @property
    def css(self) -> str:
        return {
            DueLevel.DONE: "due-done",
            DueLevel.NONE: "due-none",
            DueLevel.LATER: "due-later",
            DueLevel.UPCOMING: "due-upcoming",
            DueLevel.SOON: "due-soon",
            DueLevel.TODAY: "due-today",
            DueLevel.OVERDUE: "due-overdue",
        }[self]

    @property
    def label(self) -> str:
        return {
            DueLevel.DONE: "Проведён",
            DueLevel.NONE: "Нужно уточнить",
            DueLevel.LATER: "Позднее",
            DueLevel.UPCOMING: "Скоро",
            DueLevel.SOON: "Ближайшие",
            DueLevel.TODAY: "Сегодня",
            DueLevel.OVERDUE: "Просрочено",
        }[self]


@dataclass(frozen=True)
class DueInfo:
    level: DueLevel
    days: int | None
    text: str
    date: date | None

    @property
    def css(self) -> str:
        return self.level.css


def add_months(value: date, months: int) -> date:
    """Календарное прибавление месяцев: 31.01 + 1 мес → последний день февраля."""
    if months == 0:
        return value
    year = value.year + (value.month - 1 + months) // 12
    month = (value.month - 1 + months) % 12 + 1
    day = min(value.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def previous_workday(value: date) -> date:
    """Предыдущий рабочий день (без учёта праздников: календарь не задан)."""
    result = value
    while result.weekday() >= 5:  # 5 = суббота, 6 = воскресенье
        result -= timedelta(days=1)
    return result


def _plural_days(n: int) -> str:
    n = abs(n)
    if n % 10 == 1 and n % 100 != 11:
        return f"{n} день"
    if n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14):
        return f"{n} дня"
    return f"{n} дней"


def due_info(
    plan_date: date | None,
    *,
    status: str,
    today: date | None = None,
    soon_days: int = config.DEFAULT_SOON_DAYS,
    upcoming_days: int = config.DEFAULT_UPCOMING_DAYS,
) -> DueInfo:
    """Определить срочность записи инструктажа."""
    today = today or config.today()
    if status == "Проведён":
        return DueInfo(DueLevel.DONE, None, "Проведён", plan_date)
    if status == "Отменён":
        return DueInfo(DueLevel.DONE, None, "Отменён", plan_date)
    if plan_date is None:
        return DueInfo(DueLevel.NONE, None, "Дата не задана", None)

    delta = (plan_date - today).days
    if delta < 0:
        return DueInfo(DueLevel.OVERDUE, delta, f"Просрочено на {_plural_days(delta)}", plan_date)
    if delta == 0:
        return DueInfo(DueLevel.TODAY, 0, "Сегодня", plan_date)
    if delta <= soon_days:
        return DueInfo(DueLevel.SOON, delta, f"Через {_plural_days(delta)}", plan_date)
    if delta <= upcoming_days:
        return DueInfo(DueLevel.UPCOMING, delta, f"Через {_plural_days(delta)}", plan_date)
    return DueInfo(DueLevel.LATER, delta, f"Через {_plural_days(delta)}", plan_date)


def next_repeat_date(fact_date: date, period_months: int) -> date:
    """Предельная дата следующего повторного инструктажа.

    Считается от даты фактического проведения. Если срок попадает на выходной,
    предлагается предыдущий рабочий день, чтобы не выйти за предел.
    """
    return previous_workday(add_months(fact_date, period_months))


def instructor_valid(valid_to: date | None, on: date) -> bool:
    if valid_to is None:
        return True
    return valid_to >= on


def sort_key(level: DueLevel, days: int | None) -> tuple[int, int]:
    return (-int(level), days if days is not None else 10**6)
