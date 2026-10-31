"""Проверка DOCX-шаблонов печати и создание примера.

Использование:
    python scripts/check_templates.py
    python scripts/check_templates.py --make-example

Проверка показывает, какие шаблоны найдены, какие плейсхолдеры в них есть,
каких не хватает и какие указаны с ошибкой.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from docx import Document  # noqa: E402
from docx.shared import Cm, Pt  # noqa: E402

from ot_registry import config, docx_templates  # noqa: E402


def make_example(target: Path, cards: int = 1) -> Path:
    """Создать пример шаблона с плейсхолдерами.

    Служит отправной точкой: специалист открывает файл в Word, меняет
    оформление и порядок полей, оставляя `{{...}}` на месте.
    """
    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.2)
    document.styles["Normal"].font.name = "Times New Roman"
    document.styles["Normal"].font.size = Pt(10)

    header = document.add_paragraph()
    header.alignment = 1
    run = header.add_run("{{org_name}}")
    run.bold = True
    run.font.size = Pt(12)

    sub = document.add_paragraph()
    sub.alignment = 1
    sub_run = sub.add_run("ИНН {{org_inn}} · {{org_address}}")
    sub_run.font.size = Pt(8)

    title = document.add_paragraph()
    title.alignment = 1
    title_run = title.add_run("Инструктаж по охране труда: {{type}}")
    title_run.bold = True
    title_run.font.size = Pt(11)

    rows = [
        ("ФИО работника", "{{fio}}"),
        ("Профессия (должность)", "{{position}}"),
        ("Подразделение", "{{department}}"),
        ("Дата рождения", "{{birth_date}}"),
        ("Вид инструктажа", "{{type}}"),
        ("Причина", "{{reason}}"),
        ("ЛНА (программа/инструкции)", "{{program}}"),
        ("Дата проведения", "{{fact_date}}"),
        ("ФИО проводившего", "{{instructor_fio}}"),
        ("Должность проводившего", "{{instructor_position}}"),
        ("Код бланка", "{{code}}"),
        ("Номера после регистрации", "ОТ {{number_ot}} · {{number_type}} · дело {{case_id}} · пакет {{pack_id}}"),
    ]
    table = document.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    for label, value in rows:
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
        cells[0].width = Cm(5.2)
        cells[1].width = Cm(10.6)

    document.add_paragraph()
    sign = document.add_table(rows=2, cols=2)
    sign.style = "Table Grid"
    sign.cell(0, 0).text = "Инструктаж провёл"
    sign.cell(0, 1).text = "подпись {{signature_line}}"
    sign.cell(1, 0).text = "Инструктаж получил"
    sign.cell(1, 1).text = "подпись {{signature_line}}"

    note = document.add_paragraph()
    note_run = note.add_run(
        "Код бланка служит для поиска записи в приложении и не является подписью. "
        "После проведения инструктажа впишите номера в отведённые поля."
    )
    note_run.font.size = Pt(7)
    note_run.italic = True

    target.parent.mkdir(parents=True, exist_ok=True)
    document.save(target)
    return target


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Проверка шаблонов печати")
    parser.add_argument("--make-example", action="store_true",
                        help="создать пример шаблона в каталоге шаблонов")
    parser.add_argument("--cards", type=int, default=1,
                        help="сколько карточек в примере (по умолчанию 1)")
    args = parser.parse_args(argv)

    folder = config.templates_dir()
    print(f"Каталог шаблонов: {folder}\n")

    if args.make_example:
        target = folder / f"card_{args.cards}.docx"
        if target.exists():
            print(f"Файл уже есть: {target}")
        else:
            make_example(target, cards=args.cards)
            print(f"Создан пример: {target}\n")

    infos = docx_templates.discover()
    if not infos:
        print("Шаблоны не найдены. Печать будет использовать встроенный бланк.")
        print("Создать пример: python scripts/check_templates.py --make-example")
        return 0

    ok = True
    for info in infos:
        print(f"Файл: {info.path.name} — карточек на листе: {info.cards}")
        print(f"  найдено плейсхолдеров: {len(info.found)}")
        if info.found:
            print(f"    {', '.join(info.found)}")
        if info.missing:
            ok = False
            print(f"  ОТСУТСТВУЮТ важные поля: {', '.join(info.missing)}")
        if info.unknown:
            ok = False
            print(f"  НЕИЗВЕСТНЫЕ плейсхолдеры (не будут заменены): "
                  f"{', '.join(info.unknown)}")
        if not info.found:
            ok = False
            print("  В шаблоне нет ни одного известного плейсхолдера.")
        print()

    print("Доступные плейсхолдеры:")
    for name, description in docx_templates.placeholder_help():
        print(f"  {{{{{name}}}}} — {description}")

    print()
    print("Итог:", "все шаблоны корректны" if ok else "есть замечания (см. выше)")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
