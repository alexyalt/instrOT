"""Резервное копирование: локальные копии по расписанию и выгрузка на почту.

Две независимые части:

1. **Локальные копии.** Создаются перед каждой записью (это уже делает
   `excel_store`) и дополнительно по расписанию — не чаще, чем задано
   в настройках. Хранится ограниченное число, старые удаляются.

2. **Выгрузка на почту.** Раз в N календарных дней приложение отправляет
   архив с книгой на указанный адрес. Нужно, чтобы копия была вне
   компьютера: локальные копии не спасают от поломки диска.

Почта настраивается в организации. Пароль приложения хранится в локальном
файле настроек рядом с книгой — в саму книгу Excel он не попадает.
Это ограничение: файл доступен пользователю компьютера.
"""

from __future__ import annotations

import json
import smtplib
import ssl
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path
from typing import Any

from . import config
from .errors import ConfigError, ValidationError
from .excel_store import Book

SETTINGS_FILE = "backup_settings.json"

# Почтовые сервисы: готовые настройки SMTP. Можно указать свой.
MAIL_PRESETS: dict[str, dict[str, Any]] = {
    "mail.ru": {"host": "smtp.mail.ru", "port": 465, "ssl": True},
    "yandex.ru": {"host": "smtp.yandex.ru", "port": 465, "ssl": True},
    "gmail.com": {"host": "smtp.gmail.com", "port": 465, "ssl": True},
    "outlook.com": {"host": "smtp.office365.com", "port": 587, "ssl": False},
    "custom": {"host": "", "port": 465, "ssl": True},
}


@dataclass
class BackupSettings:
    """Настройки резервного копирования организации."""

    keep_local: int = config.DEFAULT_BACKUP_KEEP
    # Как часто отправлять копию на почту, календарных дней.
    email_every_days: int = config.DEFAULT_EMAIL_BACKUP_EVERY_DAYS
    email_enabled: bool = False
    smtp_preset: str = "mail.ru"
    smtp_host: str = ""
    smtp_port: int = 465
    smtp_ssl: bool = True
    sender: str = ""
    recipient: str = ""
    username: str = ""
    last_email_at: str = ""
    last_email_error: str = ""
    last_local_at: str = ""

    @property
    def resolved_host(self) -> str:
        if self.smtp_host:
            return self.smtp_host
        preset = MAIL_PRESETS.get(self.smtp_preset, {})
        return str(preset.get("host") or "")

    @property
    def resolved_port(self) -> int:
        if self.smtp_host:
            return self.smtp_port
        preset = MAIL_PRESETS.get(self.smtp_preset, {})
        return int(preset.get("port") or 465)

    @property
    def resolved_ssl(self) -> bool:
        if self.smtp_host:
            return self.smtp_ssl
        preset = MAIL_PRESETS.get(self.smtp_preset, {})
        return bool(preset.get("ssl", True))


def settings_path(book: Book) -> Path:
    return book.path.parent / SETTINGS_FILE


def password_path(book: Book) -> Path:
    """Пароль приложения хранится отдельно от настроек."""
    return book.path.parent / "mail_password.txt"


def load_settings(book: Book) -> BackupSettings:
    path = settings_path(book)
    if not path.exists():
        return BackupSettings()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return BackupSettings()
    known = {f for f in BackupSettings.__dataclass_fields__}
    return BackupSettings(**{k: v for k, v in data.items() if k in known})


def save_settings(book: Book, settings: BackupSettings, *,
                  password: str | None = None) -> None:
    path = settings_path(book)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {f: getattr(settings, f) for f in BackupSettings.__dataclass_fields__}
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    if password is not None:
        # Пустая строка означает «не менять пароль».
        if password:
            password_path(book).write_text(password, encoding="utf-8")


def read_password(book: Book) -> str:
    path = password_path(book)
    if not path.exists():
        return ""
    try:
        return path.read_text(encoding="utf-8").strip()
    except OSError:
        return ""


def validate_settings(settings: BackupSettings) -> None:
    """Проверить настройки перед сохранением."""
    if settings.keep_local < 1:
        raise ValidationError("Число локальных копий должно быть не меньше 1.",
                              fields={"keep_local": "Не меньше 1."})
    if settings.keep_local > 500:
        raise ValidationError("Слишком много локальных копий: до 500.",
                              fields={"keep_local": "Не больше 500."})
    if settings.email_every_days < 1:
        raise ValidationError("Периодичность отправки — не меньше 1 дня.",
                              fields={"email_every_days": "Не меньше 1."})
    if settings.email_enabled:
        if not settings.sender or "@" not in settings.sender:
            raise ValidationError("Укажите адрес отправителя.",
                                  fields={"sender": "Нужен адрес электронной почты."})
        if not settings.recipient or "@" not in settings.recipient:
            raise ValidationError("Укажите адрес получателя копий.",
                                  fields={"recipient": "Нужен адрес электронной почты."})
        if not settings.resolved_host:
            raise ValidationError("Не указан сервер отправки (SMTP).",
                                  fields={"smtp_host": "Укажите сервер или выберите сервис."})


# --------------------------------------------------------------------------- #
# Локальные копии
# --------------------------------------------------------------------------- #

def rotate(book: Book, keep: int | None = None) -> int:
    """Удалить старые локальные копии, оставив последние `keep` штук.

    Возвращает количество удалённых файлов. Последняя копия не удаляется
    никогда, даже если лимит выставлен в 1.
    """
    limit = keep if keep is not None else load_settings(book).keep_local
    limit = max(1, int(limit))
    folder = book.backups_dir()
    if not folder.exists():
        return 0
    # Сортируем по времени изменения: имя содержит микросекунды, но
    # скопированные файлы могут иметь одинаковое время.
    files = sorted(folder.glob("registry_*.xlsx"), key=lambda p: (p.name))
    removed = 0
    while len(files) > limit:
        oldest = files.pop(0)
        try:
            oldest.unlink()
            removed += 1
        except OSError:
            break
    return removed


def make_local_copy(book: Book, *, reason: str = "по расписанию") -> Path:
    """Создать локальную копию (кроме тех, что делает сохранение)."""
    import shutil

    folder = book.backups_dir()
    folder.mkdir(parents=True, exist_ok=True)
    # Микросекунды в имени: несколько копий в одну секунду не перезаписывают
    # друг друга (например, при пачке операций подряд).
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    target = folder / f"registry_{stamp}.xlsx"
    shutil.copy2(book.path, target)
    rotate(book)

    settings = load_settings(book)
    settings.last_local_at = datetime.now().isoformat(timespec="seconds")
    save_settings(book, settings)

    from .journal import log_event

    log_event(book.logs_dir(), organization=book.name, action="backup_local",
              entity="Резерв", reason=reason, details={"file": target.name})
    return target


def make_zip(book: Book) -> Path:
    """Архив с книгой для отправки: книга + краткие сведения."""
    folder = book.backups_dir() / "outbox"
    folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = folder / f"backup_{stamp}.zip"
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.write(book.path, "registry.xlsx")
        info = (
            f"Организация: {book.org_value('name')}\n"
            f"Полное наименование: {book.org_value('full_name')}\n"
            f"ИНН: {book.org_value('inn')}\n"
            f"Версия схемы: {book.org_value('schema_version')}\n"
            f"Копия создана: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
        )
        archive.writestr("info.txt", info.encode("utf-8"))
    return target


# --------------------------------------------------------------------------- #
# Почта
# --------------------------------------------------------------------------- #

def should_send_email(settings: BackupSettings, *, today: date | None = None) -> bool:
    if not settings.email_enabled:
        return False
    today = today or config.today()
    if not settings.last_email_at:
        return True
    try:
        last = datetime.fromisoformat(settings.last_email_at).date()
    except ValueError:
        return True
    return (today - last) >= timedelta(days=max(1, settings.email_every_days))


def send_email(book: Book, settings: BackupSettings | None = None, *,
               password: str | None = None,
               reason: str = "по расписанию") -> dict[str, Any]:
    """Отправить архив с книгой на почту.

    Возвращает результат; при ошибке не поднимает исключение, а пишет её
    в настройки и журнал: сбой почты не должен мешать работе с книгой.
    """
    settings = settings or load_settings(book)
    secret = password if password is not None else read_password(book)
    validate_settings(settings)

    archive = make_zip(book)
    message = EmailMessage()
    org_name = book.org_value("full_name") or book.org_value("name") or book.path.stem
    message["Subject"] = (
        f"Резервная копия: {org_name} — {datetime.now().strftime('%d.%m.%Y')}"
    )
    message["From"] = settings.sender
    message["To"] = settings.recipient
    message.set_content(
        "Автоматическая резервная копия книги учёта инструктажей.\n\n"
        f"Организация: {org_name}\n"
        f"Файл: registry.xlsx (внутри архива)\n"
        f"Создана: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
        f"Причина: {reason}\n\n"
        "Чтобы восстановить: распакуйте архив и подключите registry.xlsx "
        "в приложении через «Подключить существующую».\n"
        "Не удаляйте более новые копии, пока не убедитесь, что эта рабочая.\n"
    )
    message.add_attachment(
        archive.read_bytes(),
        maintype="application",
        subtype="zip",
        filename=archive.name,
    )

    result: dict[str, Any] = {"sent": False, "error": "", "file": archive.name}
    try:
        context = ssl.create_default_context()
        if settings.resolved_ssl:
            with smtplib.SMTP_SSL(settings.resolved_host, settings.resolved_port,
                                  context=context, timeout=30) as server:
                if secret:
                    server.login(settings.username or settings.sender, secret)
                server.send_message(message)
        else:
            with smtplib.SMTP(settings.resolved_host, settings.resolved_port,
                              timeout=30) as server:
                server.starttls(context=context)
                if secret:
                    server.login(settings.username or settings.sender, secret)
                server.send_message(message)
        result["sent"] = True
        settings.last_email_at = datetime.now().isoformat(timespec="seconds")
        settings.last_email_error = ""
    except Exception as exc:  # сети, пароли и SMTP отвечают по-разному
        settings.last_email_error = f"{type(exc).__name__}: {exc}"
        result["error"] = settings.last_email_error

    save_settings(book, settings)

    from .journal import log_event

    log_event(
        book.logs_dir(),
        organization=book.name,
        action="backup_email",
        entity="Резерв",
        reason=reason,
        details={"sent": result["sent"], "file": archive.name,
                 "error": result["error"][:300]},
    )
    return result


def run_scheduled(book: Book) -> dict[str, Any]:
    """Плановая проверка: отправить копию, если срок подошёл.

    Вызывается при открытии организации. Молча ничего не делает, если
    почта не настроена или срок не наступил.
    """
    settings = load_settings(book)
    if not should_send_email(settings):
        return {"ran": False}

    from .journal import log_event

    log_event(book.logs_dir(), organization=book.name, action="backup_due",
              entity="Резерв", reason="подошёл срок отправки копии")
    return {"ran": True, **send_email(book, settings, reason="по расписанию")}


def email_status(settings: BackupSettings) -> str:
    """Текстовая подсказка для интерфейса."""
    if not settings.email_enabled:
        return "отправка на почту выключена"
    if not settings.last_email_at:
        return "копия ещё не отправлялась"
    when = settings.last_email_at.replace("T", " ")
    if settings.last_email_error:
        return f"последняя отправка не удалась ({when}): {settings.last_email_error[:120]}"
    return f"последняя отправка: {when}"
