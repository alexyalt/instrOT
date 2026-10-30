"""Межпроцессная блокировка записи книги.

Один процесс приложения не должен писать в книгу одновременно с другим.
Блокировка — advisory, файловая: работает на Windows (msvcrt) и POSIX (fcntl).
Она не защищает от ручного редактирования в Excel — для этого применяется
проверка отпечатка файла перед заменой.
"""

from __future__ import annotations

import os
import time
from pathlib import Path

from .errors import BookLockedError

_WINDOWS = os.name == "nt"

if _WINDOWS:  # pragma: no cover - платформенная ветка
    import msvcrt
else:
    import fcntl


class FileLock:
    """Контекстный менеджер блокировки файла.

    Пример:
        with FileLock(path):
            ...
    """

    def __init__(self, path: Path, *, timeout: float = 5.0, poll: float = 0.1) -> None:
        self.path = Path(path)
        self.timeout = timeout
        self.poll = poll
        self._fh = None

    def __enter__(self) -> "FileLock":
        self.acquire()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.release()

    def acquire(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        fh = open(self.path, "a+b")
        deadline = time.monotonic() + self.timeout
        while True:
            try:
                if _WINDOWS:  # pragma: no cover
                    fh.seek(0)
                    msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    fcntl.flock(fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except OSError:
                if time.monotonic() >= deadline:
                    fh.close()
                    raise BookLockedError(
                        "Книга занята другим процессом приложения. "
                        "Закройте второе окно или подождите завершения операции."
                    ) from None
                time.sleep(self.poll)
        fh.seek(0)
        fh.truncate()
        fh.write(str(os.getpid()).encode("ascii"))
        fh.flush()
        self._fh = fh

    def release(self) -> None:
        fh = self._fh
        self._fh = None
        if fh is None:
            return
        try:
            if _WINDOWS:  # pragma: no cover
                fh.seek(0)
                msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(fh.fileno(), fcntl.LOCK_UN)
        except OSError:
            pass
        finally:
            fh.close()
