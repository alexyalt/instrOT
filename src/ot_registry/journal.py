"""Журнал действий: JSONL по месяцам плюс читаемая строка.

Журнал нужен, чтобы понимать, кто и когда изменил данные. Он не является
доверенным аудитом: пользователь с доступом к компьютеру может его изменить.
"""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any


def _journal_path(base: Path, when: datetime) -> Path:
    return base / f"journal_{when:%Y_%m}.jsonl"


def log_event(
    base_dir: Path,
    *,
    organization: str,
    action: str,
    entity: str = "",
    entity_id: str = "",
    details: dict[str, Any] | None = None,
    reason: str = "",
) -> None:
    """Дописать событие в журнал. Ошибки журнала не должны ломать операцию."""
    try:
        when = datetime.now().astimezone()
        base_dir.mkdir(parents=True, exist_ok=True)
        record = {
            "ts": when.isoformat(timespec="seconds"),
            "organization": organization,
            "action": action,
            "entity": entity,
            "entity_id": entity_id,
            "reason": reason,
            "user": os.environ.get("USERNAME") or os.environ.get("USER") or "",
            "details": details or {},
        }
        line = json.dumps(record, ensure_ascii=False, default=str)
        with open(_journal_path(base_dir, when), "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except OSError:
        # Журнал вспомогательный: не прерываем из-за него сохранение данных.
        pass


def read_events(base_dir: Path, *, limit: int = 200) -> list[dict[str, Any]]:
    """Последние события журнала (новые сверху)."""
    events: list[dict[str, Any]] = []
    if not base_dir.exists():
        return events
    for path in sorted(base_dir.glob("journal_*.jsonl"), reverse=True):
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except OSError:
            continue
        for line in reversed(lines):
            line = line.strip()
            if not line:
                continue
            try:
                events.append(json.loads(line))
            except json.JSONDecodeError:
                continue
            if len(events) >= limit:
                return events
    return events
