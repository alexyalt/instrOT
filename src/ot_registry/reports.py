"""Отчёты: реестр пакета, план на период, просроченные, сводка по видам."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from typing import Any, Iterable

from . import config, deadlines, rules, services
from .excel_store import Book


@dataclass
class ReportRow:
    cells: list[str]


def _fmt(value: Any) -> str:
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    return "" if value is None else str(value)


def plan_for_period(book: Book, *, date_from: date, date_to: date) -> list[dict[str, Any]]:
    """Инструктажи, запланированные на период."""
    result = []
    for row in book.table("Инструктажи").rows:
        plan_date = row.get("plan_date")
        if not isinstance(plan_date, date):
            continue
        if date_from <= plan_date <= date_to and row.get("status") == "Запланирован":
            result.append(row)
    return sorted(result, key=lambda r: r.get("plan_date"))


def overdue(book: Book) -> list[dict[str, Any]]:
    today = config.today()
    rows = [
        row for row in book.table("Инструктажи").rows
        if row.get("status") == "Запланирован"
        and isinstance(row.get("plan_date"), date)
        and row["plan_date"] < today
    ]
    return sorted(rows, key=lambda r: r["plan_date"])


def incomplete_registrations(book: Book) -> list[dict[str, Any]]:
    """Проведённые записи с неполным набором обязательных реквизитов."""
    result = []
    for row in book.table("Инструктажи").rows:
        if row.get("status") != "Проведён":
            continue
        missing = rules.check_required_fields(row)
        if missing:
            result.append({"row": row, "missing": missing})
    return result


def pack_registry(book: Book, *, case_id: str, pack_id: str) -> list[dict[str, Any]]:
    """Реестр документов пакета: что подшито в эту папку."""
    rows = [
        row for row in book.table("Инструктажи").rows
        if str(row.get("case_id") or "") == case_id
        and str(row.get("pack_id") or "") == pack_id
        and row.get("status") == "Проведён"
    ]
    return sorted(rows, key=lambda r: (r.get("fact_date") or date.min, str(r.get("number_ot") or "")))


def by_type_summary(book: Book, *, year: int | None = None) -> list[dict[str, Any]]:
    """Сводка по видам инструктажа за год."""
    year = year or config.today().year
    totals: dict[str, dict[str, int]] = {}
    for row in book.table("Инструктажи").rows:
        if row.get("status") != "Проведён":
            continue
        fact = row.get("fact_date")
        if not isinstance(fact, date) or fact.year != year:
            continue
        training_type = str(row.get("type") or "—")
        bucket = totals.setdefault(training_type, {"planned": 0, "manual": 0})
        bucket["planned"] += 1
        if str(row.get("source") or "") == "Ручной бланк":
            bucket["manual"] += 1
    return [
        {"type": key, "total": value["planned"], "manual": value["manual"]}
        for key, value in sorted(totals.items())
    ]


def repeat_due_soon(book: Book, *, days: int = 30) -> list[dict[str, Any]]:
    """Предстоящие повторные инструктажи."""
    today = config.today()
    horizon = today + timedelta(days=days)
    rows = [
        row for row in book.table("Инструктажи").rows
        if row.get("status") == "Запланирован"
        and row.get("type") == "Повторный"
        and isinstance(row.get("plan_date"), date)
        and today <= row["plan_date"] <= horizon
    ]
    return sorted(rows, key=lambda r: r["plan_date"])


def instructors_expiring(book: Book, *, days: int = 30) -> list[dict[str, Any]]:
    """Допуски инструктирующих, истекающие в ближайшие дни или уже истёкшие."""
    today = config.today()
    horizon = today + timedelta(days=days)
    result = []
    for row in book.table("Инструктирующие").rows:
        valid_to = row.get("valid_to")
        if isinstance(valid_to, date) and valid_to <= horizon:
            result.append(row)
    return sorted(result, key=lambda r: r.get("valid_to"))


def pack_rows(book: Book) -> list[dict[str, Any]]:
    """Все пакеты с числом документов и состоянием."""
    counts: dict[tuple[str, str], int] = {}
    for row in book.table("Инструктажи").rows:
        if row.get("status") != "Проведён":
            continue
        key = (str(row.get("case_id") or ""), str(row.get("pack_id") or ""))
        counts[key] = counts.get(key, 0) + 1
    result = []
    for row in book.table("Архив").rows:
        if not row.get("pack_id"):
            continue
        key = (str(row.get("case_id") or ""), str(row.get("pack_id") or ""))
        result.append({**row, "documents": counts.get(key, 0)})
    return sorted(result, key=lambda r: (str(r.get("case_id") or ""), str(r.get("pack_id") or "")))
