"""Запуск приложения без установки: python -m ot_registry.launcher"""

from __future__ import annotations

import sys
from pathlib import Path

# Позволяет запускать из корня проекта без установки пакета.
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

if __name__ == "__main__":
    from ot_registry.launcher import main

    sys.exit(main(sys.argv[1:]))
