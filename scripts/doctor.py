"""Проверка установки: окружение, данные, шаблоны, конвертер PDF."""

from __future__ import annotations

import shutil
import sys
from pathlib import Path


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    sys.path.insert(0, str(root / "src"))

    print("instrOT — проверка окружения\n")

    version = sys.version_info
    ok = version >= (3, 10)
    print(f"{'OK ' if ok else 'ОШИБКА'} Python {version.major}.{version.minor}.{version.micro}")

    for module in ("flask", "openpyxl", "docx"):
        try:
            __import__(module)
            print(f"OK  модуль {module}")
        except ImportError:
            print(f"ОШИБКА нет модуля {module} — установите: pip install -r requirements.txt")
            ok = False

    from ot_registry import config

    data = config.data_dir()
    print(f"\nКаталог данных: {data}")
    try:
        data.mkdir(parents=True, exist_ok=True)
        probe = data / ".probe"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        print("OK  каталог данных доступен для записи")
    except OSError as exc:
        print(f"ОШИБКА нет доступа к каталогу данных: {exc}")
        ok = False

    from ot_registry.organizations import Catalog

    try:
        catalog = Catalog.load(data)
        print(f"OK  организаций подключено: {len(catalog)}")
        for entry in catalog:
            path = entry.resolve(data)
            state = "найден" if path.exists() else "НЕ НАЙДЕН"
            print(f"    - {entry.name}: {state}")
    except Exception as exc:
        print(f"ОШИБКА список организаций: {exc}")
        ok = False

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        print(f"OK  LibreOffice: {soffice}")
    else:
        print("ВНИМАНИЕ LibreOffice не найден — PDF не будет формироваться, останется DOCX")

    print("\nИтог:", "всё готово" if ok else "есть проблемы, см. выше")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
