"""Конфигурация приложения: пути, версия схемы, значения по умолчанию."""

from __future__ import annotations

import os
import sys
from datetime import date
from pathlib import Path

APP_NAME = "instrOT"
APP_TITLE = "Учёт инструктажей по охране труда"
VERSION = "1.0.0.dev0"

SCHEMA_VERSION = 1

# Зарезервированный код пустого бумажного бланка.
MANUAL_CODE = "000000"
MANUAL_CODE_DISPLAY = "000-000"

# Пороги подсветки сроков по умолчанию (дни).
DEFAULT_SOON_DAYS = 7
DEFAULT_UPCOMING_DAYS = 14

# Число карточек инструктажа на одном листе A4.
# 1 — полноразмерная карточка: максимум места для подписей и ручного
# заполнения. Больше 1 — экономия бумаги; проверяйте читаемость на принтере.
# Рекомендуемые значения: 1, 2 или 3. Больше 5 не проверялось.
CARDS_PER_PAGE = 1
MAX_CARDS_PER_PAGE = 5

# Сколько резервных копий хранить по умолчанию.
DEFAULT_BACKUP_KEEP = 50
DEFAULT_DAILY_BACKUP_KEEP = 30

# Периодичность повторного инструктажа по охране труда.
# П. 87 Правил № 2464: повторный инструктаж проводится не реже одного раза в 6 месяцев.
DEFAULT_PERIOD_MONTHS = 6
# Срок свыше этого значения требует явного основания от работодателя
# (например, освобождение от первичного инструктажа для отдельных категорий).
MAX_DEFAULT_PERIOD_MONTHS = 6

TRAINING_TYPES = [
    "Вводный",
    "Первичный",
    "Повторный",
    "Внеплановый",
    "Целевой",
]

# Префиксы номеров по видам инструктажа.
TYPE_PREFIX = {
    "Вводный": "ВВ",
    "Первичный": "ПР",
    "Повторный": "ПВ",
    "Внеплановый": "ВН",
    "Целевой": "ЦЛ",
}

WORKER_STATUSES = ["Работает", "Уволен"]

TRAINING_STATUSES = ["Запланирован", "Проведён", "Отменён"]

TRAINING_SOURCES = ["Система", "Ручной бланк", "Excel"]


def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def app_dir() -> Path:
    """Каталог программы (рядом с EXE или корень проекта при запуске из исходников)."""
    if _is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def data_dir() -> Path:
    """Каталог пользовательских данных (книги организаций, резерв, журналы).

    По умолчанию — пользовательский каталог, а не папка программы, чтобы
    обновление приложения не затрагивало данные.
    """
    override = os.environ.get("OTR_DATA_DIR")
    if override:
        return Path(override).expanduser().resolve()
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA") or str(Path.home())
        return Path(base) / "OTRegistry"
    return Path.home() / ".otregistry"


def settings_dir() -> Path:
    return data_dir() / "settings"


def organizations_dir() -> Path:
    return data_dir() / "organizations"


def templates_dir() -> Path:
    """Шаблоны печатных форм: сначала пользовательские, затем поставляемые."""
    user = data_dir() / "templates"
    if user.exists():
        return user
    return app_dir() / "resources" / "forms"


def today() -> date:
    return date.today()
