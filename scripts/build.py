"""Сборка приложения в папку EXE (PyInstaller --onedir).

Запуск из корня проекта:
    python scripts/build.py

Требуется установленный PyInstaller. Результат — dist/instrOT/ с исполняемым
файлом и ресурсами. Данные пользователя создаются отдельно и не входят в сборку.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
NAME = "instrOT"
ENTRY = ROOT / "src" / "ot_registry" / "launcher.py"
RESOURCES = ROOT / "resources"


def clean() -> None:
    for folder in ("build", "dist"):
        path = ROOT / folder
        if path.exists():
            shutil.rmtree(path, ignore_errors=True)


def build() -> int:
    if not ENTRY.exists():
        print(f"Не найден файл запуска: {ENTRY}")
        return 1

    add_data = []
    if RESOURCES.exists():
        add_data.append(f"{RESOURCES}{';' if sys.platform == 'win32' else ':'}resources")

    command = [
        sys.executable or "python",
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onedir",
        "--noconsole",
        "--name",
        NAME,
        "--paths",
        str(ROOT / "src"),
    ]
    for item in add_data:
        command += ["--add-data", item]
    command.append(str(ENTRY))

    print("Сборка:", " ".join(command))
    result = subprocess.run(command, cwd=ROOT)
    if result.returncode != 0:
        return result.returncode

    out = ROOT / "dist" / NAME
    if out.exists():
        readme = out / "ЧИТАТЬ.txt"
        readme.write_text(
            "instrOT — учёт инструктажей по охране труда\n\n"
            "Запуск: instrOT.exe\n\n"
            "При первом запуске создаётся каталог данных:\n"
            "  %LOCALAPPDATA%\\OTRegistry\n\n"
            "Там находятся книги организаций, резервные копии и журналы.\n"
            "Каталог можно перенести на другой диск, задав переменную окружения\n"
            "OTR_DATA_DIR.\n\n"
            "Папку data/ удалять нельзя: это рабочие данные.\n"
            "Обновление программы не затрагивает данные.\n",
            encoding="utf-8",
        )
        print(f"Готово: {out}")
    return 0


if __name__ == "__main__":
    clean()
    sys.exit(build())
