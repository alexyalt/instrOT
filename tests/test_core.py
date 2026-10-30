"""Автотесты: сроки, нумерация, реквизиты, бумажный режим, ручной Excel."""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from ot_registry import config, deadlines, excel_store, registration, rules, services  # noqa: E402
from ot_registry.errors import ConflictError, ValidationError  # noqa: E402


@pytest.fixture()
def book(tmp_path):
    return excel_store.create_book(
        tmp_path / "org" / "registry.xlsx",
        data_dir=tmp_path / "org",
        name="ООО Тест",
        inn="7700000000",
    )


@pytest.fixture()
def worker(book):
    return services.add_worker(
        book,
        {
            "fio": "Иванов Иван Иванович",
            "position": "Дворник",
            "department": "Участок 1",
            "birth_date": date(1990, 5, 17),
            "hire_date": date(2024, 3, 1),
        },
    )


@pytest.fixture()
def instructor(book):
    return services.add_instructor(
        book,
        {"fio": "Петров Пётр Петрович", "types": "Вводный,Повторный,Внеплановый", "valid_to": date(2030, 1, 1)},
    )


# --------------------------------------------------------------------------- #
# Сроки
# --------------------------------------------------------------------------- #

class TestDeadlines:
    def test_overdue(self):
        info = deadlines.due_info(date(2026, 10, 1), status="Запланирован", today=date(2026, 10, 30))
        assert info.level == deadlines.DueLevel.OVERDUE
        assert "Просрочено" in info.text

    def test_today(self):
        info = deadlines.due_info(date(2026, 10, 30), status="Запланирован", today=date(2026, 10, 30))
        assert info.level == deadlines.DueLevel.TODAY
        assert info.text == "Сегодня"

    def test_thresholds(self):
        today = date(2026, 10, 30)
        assert deadlines.due_info(today.replace(day=29), status="Запланирован", today=today).level == deadlines.DueLevel.OVERDUE
        soon = deadlines.due_info(date(2026, 11, 3), status="Запланирован", today=today)
        assert soon.level == deadlines.DueLevel.SOON
        upcoming = deadlines.due_info(date(2026, 11, 10), status="Запланирован", today=today)
        assert upcoming.level == deadlines.DueLevel.UPCOMING
        later = deadlines.due_info(date(2026, 12, 20), status="Запланирован", today=today)
        assert later.level == deadlines.DueLevel.LATER

    def test_done_ignores_date(self):
        info = deadlines.due_info(date(2026, 1, 1), status="Проведён", today=date(2026, 10, 30))
        assert info.level == deadlines.DueLevel.DONE

    def test_missing_date_is_not_green(self):
        info = deadlines.due_info(None, status="Запланирован")
        assert info.level == deadlines.DueLevel.NONE
        assert "уточнить" in info.text

    def test_add_months_end_of_month(self):
        assert deadlines.add_months(date(2026, 1, 31), 1) == date(2026, 2, 28)
        assert deadlines.add_months(date(2024, 1, 31), 1) == date(2024, 2, 29)

    def test_repeat_six_months(self):
        assert rules.next_repeat(date(2026, 10, 30)) == date(2027, 4, 30)

    def test_repeat_on_weekend_moves_back(self):
        # 1 мая 2027 — суббота, поэтому предлагается пятница 30 апреля.
        assert rules.next_repeat(date(2026, 10, 31)) == date(2027, 4, 30)

    def test_period_longer_than_six_months_is_flagged(self):
        rule = rules.period_rule("Повторный", organization_period=12)
        assert rule.note
        assert "6" in rule.note


# --------------------------------------------------------------------------- #
# Работники и реквизиты
# --------------------------------------------------------------------------- #

class TestWorkers:
    def test_personnel_number_keeps_leading_zeros(self, book):
        row = services.add_worker(book, {"fio": "Петров П", "position": "Дворник", "personnel_no": "0012"})
        fresh = excel_store.open_book(book.path, data_dir=book.data_dir)
        saved = services.worker_by_id(fresh, row["id"])
        assert saved["personnel_no"] == "0012"

    def test_duplicate_fio_requires_confirmation(self, book, worker):
        with pytest.raises(ConflictError):
            services.add_worker(book, {"fio": worker["fio"], "position": "Дворник"})
        second = services.add_worker(
            book, {"fio": worker["fio"], "position": "Дворник"}, allow_duplicate=True
        )
        assert second["id"] != worker["id"]

    def test_transfer_keeps_identity_and_history(self, book, worker):
        services.transfer_worker(book, worker["id"], position="Старший дворник", department="Участок 2")
        fresh = excel_store.open_book(book.path, data_dir=book.data_dir)
        updated = services.worker_by_id(fresh, worker["id"])
        assert updated["position"] == "Старший дворник"
        history = services.worker_history(fresh, worker["id"])
        assert any(h["event_type"] == "Перевод" for h in history)

    def test_dismiss_cancels_plans(self, book, worker, instructor):
        services.plan_training(
            book,
            {
                "worker_id": worker["id"],
                "instructor_id": instructor["id"],
                "plan_date": date.today(),
                "type": "Вводный",
                "program": "ЛНА-1",
            },
        )
        services.dismiss_worker(book, worker["id"], dismiss_date=date.today())
        fresh = excel_store.open_book(book.path, data_dir=book.data_dir)
        assert all(t["status"] != "Запланирован" for t in fresh.table("Инструктажи").rows)


# --------------------------------------------------------------------------- #
# Планирование
# --------------------------------------------------------------------------- #

class TestPlanning:
    def test_plan_gets_code_and_no_numbers(self, book, worker, instructor):
        row = services.plan_training(
            book,
            {
                "worker_id": worker["id"],
                "instructor_id": instructor["id"],
                "plan_date": date.today(),
                "type": "Вводный",
                "program": "ЛНА-1",
            },
        )
        assert row["code"].isdigit() and len(row["code"]) == 6
        assert not row["number_ot"]
        assert row["status"] == "Запланирован"

    def test_missing_birth_date_blocks_planning(self, book, instructor):
        row = services.add_worker(book, {"fio": "Без Даты", "position": "Дворник"})
        with pytest.raises(ValidationError) as exc:
            services.plan_training(
                book,
                {
                    "worker_id": row["id"],
                    "instructor_id": instructor["id"],
                    "plan_date": date.today(),
                    "type": "Вводный",
                    "program": "ЛНА-1",
                },
            )
        assert "дата рождения" in exc.value.message.lower()

    def test_unplanned_reason_required(self, book, worker, instructor):
        with pytest.raises(ValidationError):
            services.plan_training(
                book,
                {
                    "worker_id": worker["id"],
                    "instructor_id": instructor["id"],
                    "plan_date": date.today(),
                    "type": "Внеплановый",
                    "program": "ЛНА-1",
                },
            )

    def test_instructor_scope_is_checked(self, book, worker):
        narrow = services.add_instructor(book, {"fio": "Только вводный", "types": "Вводный"})
        with pytest.raises(ValidationError):
            services.plan_training(
                book,
                {
                    "worker_id": worker["id"],
                    "instructor_id": narrow["id"],
                    "plan_date": date.today(),
                    "type": "Повторный",
                    "program": "ЛНА-1",
                },
            )

    def test_expired_instructor_blocked(self, book, worker):
        expired = services.add_instructor(
            book, {"fio": "Просроченный", "types": "Вводный", "valid_to": date(2020, 1, 1)}
        )
        with pytest.raises(ValidationError):
            services.plan_training(
                book,
                {
                    "worker_id": worker["id"],
                    "instructor_id": expired["id"],
                    "plan_date": date.today(),
                    "type": "Вводный",
                    "program": "ЛНА-1",
                },
            )

    def test_codes_are_unique_across_plans(self, book, worker, instructor):
        from datetime import timedelta

        codes = set()
        for offset in range(5):
            row = services.plan_training(
                book,
                {
                    "worker_id": worker["id"],
                    "instructor_id": instructor["id"],
                    "plan_date": date.today() + timedelta(days=offset),
                    "type": "Вводный",
                    "program": "ЛНА-1",
                },
            )
            codes.add(row["code"])
        assert len(codes) == 5

    def test_identical_plan_requires_confirmation(self, book, worker, instructor):
        payload = {
            "worker_id": worker["id"],
            "instructor_id": instructor["id"],
            "plan_date": date.today(),
            "type": "Вводный",
            "program": "ЛНА-1",
        }
        services.plan_training(book, payload)
        with pytest.raises(ConflictError):
            services.plan_training(book, payload)


# --------------------------------------------------------------------------- #
# Регистрация
# --------------------------------------------------------------------------- #

class TestRegistration:
    def _plan(self, book, worker, instructor, training_type="Повторный"):
        return services.plan_training(
            book,
            {
                "worker_id": worker["id"],
                "instructor_id": instructor["id"],
                "plan_date": date.today(),
                "type": training_type,
                "program": "ЛНА-1",
            },
        )

    def test_registration_assigns_numbers(self, book, worker, instructor):
        plan = self._plan(book, worker, instructor)
        result = registration.register_training(
            book,
            plan["id"],
            {
                "fact_date": date.today(),
                "instructor_id": instructor["id"],
                "paper_confirmed": True,
            },
        )
        row = result["training"]
        assert row["status"] == "Проведён"
        assert row["number_ot"] == "ОТ000001"
        assert row["number_type"].startswith("ПВ")
        assert row["case_id"].startswith("Д")
        assert row["pack_id"].startswith("П")
        assert row["reg_date"] == date.today()

    def test_numbers_are_sequential_and_never_reused(self, book, worker, instructor):
        for _ in range(3):
            plan = self._plan(book, worker, instructor)
            registration.register_training(
                book, plan["id"],
                {"fact_date": date.today(), "instructor_id": instructor["id"], "paper_confirmed": True},
            )
        numbers = [t["number_ot"] for t in book.table("Инструктажи").rows if t["number_ot"]]
        assert numbers == ["ОТ000001", "ОТ000002", "ОТ000003"]

    def test_future_fact_date_blocked(self, book, worker, instructor):
        plan = self._plan(book, worker, instructor)
        with pytest.raises(ValidationError):
            registration.register_training(
                book, plan["id"],
                {"fact_date": date(2099, 1, 1), "instructor_id": instructor["id"], "paper_confirmed": True},
            )

    def test_without_paper_confirmation_blocked(self, book, worker, instructor):
        plan = self._plan(book, worker, instructor)
        with pytest.raises(ValidationError):
            registration.register_training(
                book, plan["id"],
                {"fact_date": date.today(), "instructor_id": instructor["id"], "paper_confirmed": False},
            )

    def test_second_registration_blocked(self, book, worker, instructor):
        plan = self._plan(book, worker, instructor)
        payload = {"fact_date": date.today(), "instructor_id": instructor["id"], "paper_confirmed": True}
        registration.register_training(book, plan["id"], payload)
        with pytest.raises(ConflictError):
            registration.register_training(book, plan["id"], payload)

    def test_repeat_creates_next_plan(self, book, worker, instructor):
        plan = self._plan(book, worker, instructor, "Повторный")
        result = registration.register_training(
            book, plan["id"],
            {"fact_date": date(2026, 10, 30), "instructor_id": instructor["id"], "paper_confirmed": True},
        )
        assert result["next_plan"] is not None
        assert result["next_plan"]["plan_date"] == date(2027, 4, 30)

    def test_non_repeat_does_not_create_plan(self, book, worker, instructor):
        plan = self._plan(book, worker, instructor, "Вводный")
        result = registration.register_training(
            book, plan["id"],
            {"fact_date": date.today(), "instructor_id": instructor["id"], "paper_confirmed": True},
        )
        assert result["next_plan"] is None


# --------------------------------------------------------------------------- #
# Бумажный режим
# --------------------------------------------------------------------------- #

class TestManualPaper:
    def test_manual_code_is_rejected_in_normal_search(self, book, worker, instructor):
        with pytest.raises(ValidationError) as exc:
            registration.find_active_by_code(book, "000-000")
        assert "бумажный бланк" in exc.value.message.lower()

    def test_manual_entry_creates_completed_record_without_plan(self, book, worker, instructor):
        result = registration.register_manual(
            book,
            {
                "worker_id": worker["id"],
                "type": "Повторный",
                "program": "ЛНА-1",
                "fact_date": date(2026, 9, 15),
                "instructor_id": instructor["id"],
                "paper_confirmed": True,
                "manual_mark": "15.09.2026/2",
            },
        )
        row = result["training"]
        assert row["status"] == "Проведён"
        assert row["source"] == "Ручной бланк"
        assert row["code"] == config.MANUAL_CODE
        assert row["plan_date"] is None
        assert row["fact_date"] == date(2026, 9, 15)
        assert row["number_ot"].startswith("ОТ")

    def test_manual_entry_links_to_existing_plan(self, book, worker, instructor):
        from datetime import timedelta

        plan_date = date.today() + timedelta(days=3)
        plan = services.plan_training(
            book,
            {
                "worker_id": worker["id"],
                "instructor_id": instructor["id"],
                "plan_date": plan_date,
                "type": "Повторный",
                "program": "ЛНА-1",
            },
        )
        plan_code = plan["code"]
        registration.register_manual(
            book,
            {
                "worker_id": worker["id"],
                "type": "Повторный",
                "program": "ЛНА-1",
                "fact_date": date.today(),
                "instructor_id": instructor["id"],
                "paper_confirmed": True,
                "plan_id": plan["id"],
            },
        )
        fresh = excel_store.open_book(book.path, data_dir=book.data_dir)
        rows = fresh.table("Инструктажи").rows
        completed = [r for r in rows if r["status"] == "Проведён"]
        # Одна проведённая запись (план завершён), плюс автоматически
        # созданный следующий повторный план.
        assert len(completed) == 1
        assert len(rows) == 2
        assert plan_code in str(completed[0]["note"])

    def test_similar_document_requires_confirmation(self, book, worker, instructor):
        payload = {
            "worker_id": worker["id"],
            "type": "Повторный",
            "program": "ЛНА-1",
            "fact_date": date(2026, 9, 15),
            "instructor_id": instructor["id"],
            "paper_confirmed": True,
        }
        registration.register_manual(book, payload)
        with pytest.raises(ConflictError):
            registration.register_manual(book, payload)
        second = registration.register_manual(book, payload, allow_similar=True)
        assert second["training"]["id"]

    def test_two_different_manual_records_allowed_same_code(self, book, worker, instructor):
        for day in (1, 2, 3):
            registration.register_manual(
                book,
                {
                    "worker_id": worker["id"],
                    "type": "Повторный",
                    "program": "ЛНА-1",
                    "fact_date": date(2026, 9, day),
                    "instructor_id": instructor["id"],
                    "paper_confirmed": True,
                },
            )
        manual_rows = [r for r in book.table("Инструктажи").rows if r["source"] == "Ручной бланк"]
        assert len(manual_rows) == 3
        assert {r["code"] for r in manual_rows} == {config.MANUAL_CODE}
        # Автоматически созданный следующий повторный план имеет обычный код.
        planned_rows = [r for r in book.table("Инструктажи").rows if r["status"] == "Запланирован"]
        assert all(r["code"] != config.MANUAL_CODE for r in planned_rows)

    def test_stale_repeat_plans_detected(self, book, worker, instructor):
        from datetime import timedelta

        plan_date = date.today() + timedelta(days=10)
        plan = services.plan_training(
            book,
            {
                "worker_id": worker["id"],
                "instructor_id": instructor["id"],
                "plan_date": plan_date,
                "type": "Повторный",
                "program": "ЛНА-1",
            },
        )
        # Бумажный документ вносится за более позднюю дату, чем стоял план.
        registration.register_manual(
            book,
            {
                "worker_id": worker["id"],
                "type": "Повторный",
                "program": "ЛНА-1",
                "fact_date": date.today(),
                "instructor_id": instructor["id"],
                "paper_confirmed": True,
                "plan_id": plan["id"],
            },
        )
        report = registration.adjust_repeat_plans(book, dry_run=True)
        assert isinstance(report["stale"], list)


# --------------------------------------------------------------------------- #
# Хранилище
# --------------------------------------------------------------------------- #

class TestStorage:
    def test_backup_created_before_write(self, book, worker):
        backups = list((book.data_dir / "backups").glob("*.xlsx"))
        assert len(backups) >= 1

    def test_conflict_when_file_changed_externally(self, book, worker):
        stale = book.fingerprint
        services.add_worker(book, {"fio": "Второй", "position": "Дворник"})
        book.fingerprint = stale
        with pytest.raises(Exception):
            book.table("Работники").add({"id": "x", "fio": "Третий", "position": "Д", "status": "Работает"})
            excel_store.save_book(book, action="test", expected_fingerprint=stale)

    def test_unknown_column_rejected(self, book, tmp_path):
        from openpyxl import load_workbook

        wb = load_workbook(book.path)
        wb["Работники"].cell(row=1, column=1, value="Неизвестная колонка")
        wb.save(book.path)
        from ot_registry.errors import ConfigError

        with pytest.raises(ConfigError):
            excel_store.open_book(book.path, data_dir=book.data_dir)

    def test_dates_survive_roundtrip(self, book, worker):
        fresh = excel_store.open_book(book.path, data_dir=book.data_dir)
        saved = services.worker_by_id(fresh, worker["id"])
        assert saved["birth_date"] == date(1990, 5, 17)
        assert saved["hire_date"] == date(2024, 3, 1)

    def test_schema_version_written(self, book):
        fresh = excel_store.open_book(book.path, data_dir=book.data_dir)
        assert int(fresh.org_value("schema_version")) == config.SCHEMA_VERSION
