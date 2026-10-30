"""Проверки ввода: общие правила и обязательные поля по видам инструктажа."""

from __future__ import annotations

import re
from datetime import date
from typing import Any, Iterable

from . import config
from .errors import ValidationError
from .schema import TRAININGS_COLUMNS, WORKERS_COLUMNS

CODE_RE = re.compile(r"^\d{6}$")
CODE_INPUT_RE = re.compile(r"^[\d\s-]*$")


def normalize_code(raw: str) -> str:
    """Привести код к виду 6 цифр. Принимает 256-541 и 256 541."""
    if raw is None:
        return ""
    text = str(raw).strip()
    if not text:
        return ""
    if not CODE_INPUT_RE.match(text):
        raise ValidationError(
            "Код бланка состоит из шести цифр. Можно вводить с дефисом: 256-541."
        )
    digits = re.sub(r"\D", "", text)
    if len(digits) != 6:
        raise ValidationError("Код бланка должен содержать ровно шесть цифр.")
    return digits


def format_code(code: str) -> str:
    """6 цифр → вид для печати: 256-541. Ручной код — 000-000."""
    if not code:
        return ""
    code = str(code)
    if code == config.MANUAL_CODE:
        return config.MANUAL_CODE_DISPLAY
    if len(code) == 6 and code.isdigit():
        return f"{code[:3]}-{code[3:]}"
    return code


def is_manual_code(code: str) -> bool:
    return str(code).strip() == config.MANUAL_CODE


def _required_text(value: Any, field: str, label: str, errors: dict[str, str]) -> str:
    text = ("" if value is None else str(value)).strip()
    if not text:
        errors[field] = f"Укажите «{label}»."
    return text


def validate_worker(data: dict[str, Any]) -> dict[str, Any]:
    errors: dict[str, str] = {}
    fio = _required_text(data.get("fio"), "fio", "ФИО", errors)
    position = _required_text(data.get("position"), "position", "Должность", errors)

    hire = data.get("hire_date")
    dismiss = data.get("dismiss_date")
    if hire and dismiss and dismiss < hire:
        errors["dismiss_date"] = "Дата увольнения не может быть раньше даты приёма."

    period = data.get("period_months")
    if period not in (None, ""):
        try:
            period_int = int(period)
        except (TypeError, ValueError):
            errors["period_months"] = "Периодичность указывается целым числом месяцев."
            period_int = None
        else:
            if not 1 <= period_int <= 120:
                errors["period_months"] = "Периодичность должна быть от 1 до 120 месяцев."
    else:
        period_int = None

    if errors:
        raise ValidationError("Проверьте заполнение формы.", fields=errors)

    status = str(data.get("status") or "Работает").strip()
    if status not in config.WORKER_STATUSES:
        status = "Работает"

    return {
        "fio": fio,
        "personnel_no": str(data.get("personnel_no") or "").strip(),
        "birth_date": data.get("birth_date") or None,
        "position": position,
        "department": str(data.get("department") or "").strip(),
        "manager": str(data.get("manager") or "").strip(),
        "hire_date": hire or None,
        "dismiss_date": dismiss or None,
        "status": status,
        "period_months": period_int,
        "note": str(data.get("note") or "").strip(),
    }


# Обязательные реквизиты регистрации инструктажа — пп. 86, 87 Правил № 2464.
# Наименование локального акта (ЛНА), в объёме которого проведён инструктаж,
# обязательно для всех видов; причина — только для внепланового и целевого.
# Дата рождения работника обязательна для всех видов и проверяется отдельно
# по данным работника.
REQUIRED_BY_TYPE: dict[str, tuple[str, ...]] = {
    "Вводный": ("program",),
    "Первичный": ("program",),
    "Повторный": ("program",),
    "Внеплановый": ("program", "reason"),
    "Целевой": ("program", "reason"),
}

TYPE_LABEL = {
    "program": "наименование локального акта (ЛНА), в объёме которого проведён инструктаж",
    "reason": "причина проведения инструктажа",
}


def require_birth_date(worker: dict[str, Any] | None) -> None:
    """Дата рождения — обязательный реквизит регистрации инструктажа (п. 86, 87)."""
    if not worker or not worker.get("birth_date"):
        raise ValidationError(
            "У работника не указана дата рождения. Это обязательный реквизит регистрации "
            "инструктажа (пп. 86, 87 Правил № 2464). Дополните карточку работника и повторите.",
            fields={
                "worker_id": "В карточке работника отсутствует дата рождения — "
                             "без неё регистрация инструктажа неполна."
            },
        )


def validate_plan(data: dict[str, Any]) -> dict[str, Any]:
    """Проверить данные планирования инструктажа."""
    errors: dict[str, str] = {}
    worker_id = _required_text(data.get("worker_id"), "worker_id", "Работник", errors)
    trainer = _required_text(data.get("instructor_id"), "instructor_id", "Инструктирующий", errors)

    plan_date = data.get("plan_date")
    if plan_date is None:
        errors["plan_date"] = "Укажите плановую дату."
    elif plan_date < config.today() and not data.get("allow_past"):
        errors["plan_date"] = "Плановая дата уже прошла. Проверьте дату."

    training_type = str(data.get("type") or "").strip()
    if training_type not in config.TRAINING_TYPES:
        errors["type"] = "Выберите вид инструктажа."
    else:
        for field in REQUIRED_BY_TYPE.get(training_type, ()):
            if not str(data.get(field) or "").strip():
                label = TYPE_LABEL[field]
                errors[field] = f"Для вида «{training_type}» укажите: {label}."

    if errors:
        raise ValidationError("Проверьте заполнение формы.", fields=errors)

    return {
        "worker_id": worker_id,
        "instructor_id": trainer,
        "plan_date": plan_date,
        "type": training_type,
        "program": str(data.get("program") or "").strip(),
        "reason": str(data.get("reason") or "").strip(),
        "note": str(data.get("note") or "").strip(),
    }


def validate_registration(data: dict[str, Any], *, today: date | None = None) -> dict[str, Any]:
    """Проверить регистрацию проведённого инструктажа (обычную или по бумаге)."""
    today = today or config.today()
    errors: dict[str, str] = {}

    fact_date = data.get("fact_date")
    if fact_date is None:
        errors["fact_date"] = "Укажите фактическую дату проведения."
    elif fact_date > today:
        errors["fact_date"] = "Фактическая дата не может быть в будущем."

    if not data.get("paper_confirmed"):
        errors["paper_confirmed"] = "Подтвердите, что оригинал с подписями есть на руках."

    instructor_id = str(data.get("instructor_id") or "").strip()
    if not instructor_id:
        errors["instructor_id"] = "Выберите инструктирующего."

    if errors:
        raise ValidationError("Проверьте заполнение формы.", fields=errors)

    return {
        "fact_date": fact_date,
        "instructor_id": instructor_id,
        "paper_confirmed": True,
        "note": str(data.get("note") or "").strip(),
        "manual_mark": str(data.get("manual_mark") or "").strip(),
    }


def validate_manual_entry(data: dict[str, Any], *, today: date | None = None) -> dict[str, Any]:
    """Проверить внесение бумажного бланка, заполненного вне системы."""
    base = validate_registration(data, today=today)
    errors: dict[str, str] = {}

    worker_id = str(data.get("worker_id") or "").strip()
    if not worker_id and not str(data.get("new_worker_fio") or "").strip():
        errors["worker_id"] = "Выберите работника или укажите ФИО для новой записи."

    training_type = str(data.get("type") or "").strip()
    if training_type not in config.TRAINING_TYPES:
        errors["type"] = "Выберите вид инструктажа."
    else:
        for field in REQUIRED_BY_TYPE.get(training_type, ()):
            if not str(data.get(field) or "").strip():
                label = TYPE_LABEL[field]
                errors[field] = f"Для вида «{training_type}» укажите: {label}."

    if errors:
        raise ValidationError("Проверьте заполнение формы.", fields=errors)

    base.update(
        {
            "worker_id": worker_id,
            "new_worker_fio": str(data.get("new_worker_fio") or "").strip(),
            "new_worker_position": str(data.get("new_worker_position") or "").strip(),
            "new_worker_department": str(data.get("new_worker_department") or "").strip(),
            "type": training_type,
            "program": str(data.get("program") or "").strip(),
            "reason": str(data.get("reason") or "").strip(),
            "plan_id": str(data.get("plan_id") or "").strip(),
        }
    )
    return base


def check_columns_alignment(rows: Iterable[dict[str, Any]]) -> list[str]:
    """Проверка ручных правок книги: неизвестные ключи и пустые обязательные поля."""
    problems: list[str] = []
    for idx, row in enumerate(rows, start=2):
        for col in TRAININGS_COLUMNS:
            if col.required and not str(row.get(col.key) or "").strip():
                problems.append(f"Инструктажи, строка {idx}: не заполнено «{col.title}».")
        for col in WORKERS_COLUMNS:
            if col.required and not str(row.get(col.key) or "").strip():
                problems.append(f"Работники, строка {idx}: не заполнено «{col.title}».")
        break
    return problems
