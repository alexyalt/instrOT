"""Навигация: список закладок, их порядок и определение активного раздела.

Порядок закладок хранится в настройках организации (ключ `nav_order`),
поэтому у разных организаций он может отличаться.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class NavItem:
    key: str
    endpoint: str
    label: str
    note: str = ""


# Базовый порядок закладок. Меняется пользователем в настройках.
NAV_ITEMS: tuple[NavItem, ...] = (
    NavItem("due", "dashboard", "Сроки", "просроченные и ближайшие инструктажи"),
    NavItem("trainings", "trainings_list", "Инструктажи", "все планы и проведённые"),
    NavItem("plan", "training_new", "Планировать", "создание нового плана"),
    NavItem("register", "register", "Регистрация", "отметить проведение по коду"),
    NavItem("manual", "manual_entry", "Внести бумажный", "бланки, заполненные вручную"),
    NavItem("print", "print_forms", "Печать", "заполненные и пустые бланки"),
    NavItem("workers", "workers_list", "Работники", "список сотрудников"),
    NavItem("programs", "programs_list", "Инструкции", "ЛНА, программы и их утверждение"),
    NavItem("instructors", "instructors_list", "Инструктирующие", "допуски по программам"),
    NavItem("reports", "reports_page", "Отчёты", "контроль и сводки"),
    NavItem("archive", "archive", "Архив", "дела и пакеты документов"),
    NavItem("backup", "backup_page", "Резерв", "копии, отправка на почту"),
    NavItem("excel", "excel_page", "Excel и защита", "ручной режим и защита книги"),
    NavItem("settings", "settings_page", "Настройки", "реквизиты и параметры"),
)

ITEM_BY_KEY: dict[str, NavItem] = {item.key: item for item in NAV_ITEMS}
DEFAULT_ORDER: tuple[str, ...] = tuple(item.key for item in NAV_ITEMS)

# Разделы, доступные только из дополнительных мест (в закладки не выносятся).
EXTRA_ENDPOINTS = {"about", "journal_view", "training_edit", "pack_registry"}


def parse_order(raw: str) -> list[str]:
    """Разобрать сохранённый порядок: неизвестные ключи отбрасываются."""
    keys = [part.strip() for part in str(raw or "").split(",") if part.strip()]
    known = [k for k in keys if k in ITEM_BY_KEY]
    seen: set[str] = set()
    result: list[str] = []
    for key in known:
        if key not in seen:
            seen.add(key)
            result.append(key)
    # Добавляем недостающие закладки в исходном порядке.
    for key in DEFAULT_ORDER:
        if key not in seen:
            result.append(key)
    return result


def serialize_order(keys: list[str]) -> str:
    return ",".join(k for k in keys if k in ITEM_BY_KEY)


def ordered_items(raw_order: str) -> list[NavItem]:
    return [ITEM_BY_KEY[key] for key in parse_order(raw_order)]


def move(raw_order: str, key: str, direction: str) -> str:
    """Переместить закладку на шаг вверх/вниз и вернуть новый порядок."""
    keys = parse_order(raw_order)
    if key not in keys:
        return serialize_order(keys)
    index = keys.index(key)
    target = index - 1 if direction == "up" else index + 1
    if 0 <= target < len(keys):
        keys[index], keys[target] = keys[target], keys[index]
    return serialize_order(keys)


def active_key(request) -> str:
    """Определить активную закладку по имени endpoint."""
    endpoint = (request.endpoint or "").split(".")[-1]
    for item in NAV_ITEMS:
        if item.endpoint == endpoint:
            return item.key
    # Разделы без собственной закладки подсвечивают близкий по смыслу.
    related = {
        "training_edit": "trainings",
        "pack_registry": "archive",
        "program_new": "programs",
        "program_edit": "programs",
        "instructor_new": "instructors",
        "instructor_detail": "instructors",
        "journal_view": "excel",
        "org_new": "settings",
        "connect_org": "settings",
        "select_org": "due",
    }
    return related.get(endpoint, "")


# Палитра цветов фирмы: достаточно различимые оттенки.
BRAND_COLORS: tuple[tuple[str, str], ...] = (
    ("#1F4E79", "Тёмно-синий"),
    ("#2E7D32", "Зелёный"),
    ("#6A1B9A", "Фиолетовый"),
    ("#B26A00", "Янтарный"),
    ("#00695C", "Бирюзовый"),
    ("#8E244D", "Бордовый"),
    ("#37474F", "Графитовый"),
    ("#1565C0", "Синий"),
)


def normalize_color(value: str, fallback: str = "#1F4E79") -> str:
    """Проверить цвет вида #rrggbb."""
    text = str(value or "").strip()
    if len(text) == 7 and text.startswith("#"):
        try:
            int(text[1:], 16)
            return text.lower()
        except ValueError:
            return fallback
    return fallback


def contrast_text(hex_color: str) -> str:
    """Цвет текста поверх фона: белый или тёмный, чтобы читалось."""
    color = normalize_color(hex_color)
    r = int(color[1:3], 16)
    g = int(color[3:5], 16)
    b = int(color[5:7], 16)
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return "#1b2733" if luminance > 0.6 else "#ffffff"
