"""Прикладные исключения с понятными пользователю сообщениями."""

from __future__ import annotations


class AppError(Exception):
    """Базовая ошибка; текст показывается пользователю."""

    def __init__(self, message: str, *, details: list[str] | None = None) -> None:
        super().__init__(message)
        self.message = message
        self.details = details or []


class ValidationError(AppError):
    """Ввод не прошёл проверку."""

    def __init__(self, message: str, *, fields: dict[str, str] | None = None,
                 details: list[str] | None = None) -> None:
        super().__init__(message, details=details)
        self.fields = fields or {}


class BookLockedError(AppError):
    """Книга занята другим процессом или открыта в Excel."""


class BookChangedError(AppError):
    """Файл книги изменился между чтением и записью."""


class NotFoundError(AppError):
    pass


class ConflictError(AppError):
    pass


class ConfigError(AppError):
    pass
