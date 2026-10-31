"""Локальный веб-интерфейс приложения (Flask).

Сервер слушает только 127.0.0.1, требует одноразовый стартовый токен и
использует сессию для защиты от подделки запросов. Все изменения данных
проходят через модули services/registration/excel_store.
"""

from __future__ import annotations

import hmac
import os
import secrets
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Callable

from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    request,
    session,
    send_file,
    url_for,
)

from . import (backups, config, deadlines, docx_templates, manual_mode, migrations,
               navigation, printing, printing_template, registration,
               reports, services, validation)
from .errors import (
    AppError,
    BookChangedError,
    BookLockedError,
    ConflictError,
    NotFoundError,
    ValidationError,
)
from .excel_store import Book, open_book
from .organizations import Catalog

START_TOKEN_PARAM = "t"


def _parse_date(raw: str | None) -> date | None:
    if not raw:
        return None
    raw = raw.strip()
    if not raw:
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    raise ValidationError(f"Не удалось разобрать дату «{raw}». Формат: ДД.ММ.ГГГГ.")


def _book_summary(book: Book) -> dict[str, Any]:
    soon_days = int(book.org_value("soon_days", str(config.DEFAULT_SOON_DAYS)) or config.DEFAULT_SOON_DAYS)
    upcoming = int(book.org_value("upcoming_days", str(config.DEFAULT_UPCOMING_DAYS)) or config.DEFAULT_UPCOMING_DAYS)
    rows = []
    counts = {"overdue": 0, "soon": 0, "today": 0}
    today = config.today()
    for row in book.table("Инструктажи").rows:
        info = deadlines.due_info(
            row.get("plan_date"),
            status=str(row.get("status") or ""),
            today=today,
            soon_days=soon_days,
            upcoming_days=upcoming,
        )
        if info.level == deadlines.DueLevel.OVERDUE:
            counts["overdue"] += 1
        elif info.level == deadlines.DueLevel.TODAY:
            counts["today"] += 1
        elif info.level in (deadlines.DueLevel.SOON, deadlines.DueLevel.UPCOMING):
            counts["soon"] += 1
        rows.append((row, info))
    return {"counts": counts, "rows": rows, "soon_days": soon_days, "upcoming_days": upcoming}


def _days_until_next(settings) -> str:
    """Сколько дней до следующей плановой отправки копии на почту."""
    if not settings.email_enabled:
        return "выключено"
    if not settings.last_email_at:
        return "при следующем открытии"
    try:
        last = datetime.fromisoformat(settings.last_email_at).date()
    except ValueError:
        return "при следующем открытии"
    due = last + timedelta(days=max(1, settings.email_every_days))
    days = (due - config.today()).days
    if days <= 0:
        return "при следующем открытии"
    return f"через {days} дн."


def create_app(*, data_dir: Path | None = None, start_token: str | None = None,
               testing: bool = False) -> Flask:
    base = data_dir or config.data_dir()
    app = Flask(__name__, template_folder="templates", static_folder="static")
    app.config.update(
        SECRET_KEY=os.environ.get("OTR_SECRET") or secrets.token_hex(32),
        OTR_DATA_DIR=base,
        OTR_START_TOKEN=start_token or secrets.token_urlsafe(24),
        TESTING=testing,
        JSON_AS_ASCII=False,
        MAX_CONTENT_LENGTH=4 * 1024 * 1024,
    )

    # -- служебные помощники ---------------------------------------------
    def catalog() -> Catalog:
        return Catalog.load(base)

    def current_org_uuid() -> str:
        return str(session.get("org_uuid") or "")

    def require_org() -> tuple[Catalog, Any]:
        cat = catalog()
        org_uuid = current_org_uuid()
        if not org_uuid:
            if len(cat) == 1:
                entry = cat.entries[0]
                session["org_uuid"] = entry.uuid
                return cat, entry
            raise ValidationError("Выберите организацию на главной странице.")
        return cat, cat.require(org_uuid)

    def load_book(entry) -> Book:
        path = entry.resolve(base)
        book = open_book(path, data_dir=path.parent)
        if book.org_uuid and entry.org_uuid and book.org_uuid != entry.org_uuid:
            raise ConflictError(
                "Файл книги не соответствует выбранной организации. "
                "Проверьте, не заменён ли файл."
            )
        return book

    def safe(fn: Callable, *args, **kwargs):
        """Выполнить действие, превратив прикладные ошибки в сообщения."""
        try:
            return fn(*args, **kwargs), None
        except ValidationError as exc:
            return None, {"message": exc.message, "fields": exc.fields, "details": exc.details}
        except AppError as exc:
            return None, {"message": exc.message, "fields": {}, "details": exc.details}
        except OSError as exc:
            return None, {
                "message": "Не удалось обратиться к файлу книги. Возможно, он открыт в Excel.",
                "fields": {},
                "details": [str(exc)],
            }

    app.jinja_env.globals.update(
        app_title=config.APP_TITLE,
        app_version=config.VERSION,
        manual_code=config.MANUAL_CODE_DISPLAY,
        training_types=config.TRAINING_TYPES,
        format_code=validation.format_code,
        today=config.today,
    )
    app.jinja_env.filters["format_code"] = validation.format_code

    def fmt_date(value: Any) -> str:
        return value.strftime("%d.%m.%Y") if isinstance(value, date) else ""

    app.jinja_env.filters["ru_date"] = fmt_date

    def number_display(value: Any, *, prefix: bool = True) -> str:
        return validation.display_number(value) if prefix else validation.display_number_short(value)

    app.jinja_env.filters["num"] = number_display
    app.jinja_env.filters["num_short"] = lambda v: validation.display_number_short(v)

    def instructor_cell(row: dict) -> Any:
        """Данные для выделенного блока инструктирующего в карточке."""
        return row.get("instructor_fio") or ""

    app.jinja_env.globals["instructor_cell"] = instructor_cell

    def worker_options(rows) -> list[dict[str, str]]:
        """Данные для поля выбора работника с поиском по любому слову."""
        result = []
        for row in rows:
            parts = [str(row.get("fio") or ""), str(row.get("position") or ""),
                     str(row.get("department") or "")]
            label = str(row.get("fio") or "")
            if row.get("position"):
                label += f" — {row['position']}"
            if row.get("department"):
                label += f" ({row['department']})"
            if row.get("personnel_no"):
                parts.append(str(row["personnel_no"]))
            result.append({
                "id": str(row.get("id") or ""),
                "label": label,
                "search": " ".join(parts).lower(),
            })
        return result

    def instructor_options(rows) -> list[dict[str, str]]:
        """Данные для поля выбора инструктирующего: ФИО + программа."""
        result = []
        for row in rows:
            label = str(row.get("fio") or "")
            if row.get("program"):
                label += f" — {row['program']}"
            result.append({
                "id": str(row.get("id") or ""),
                "label": label,
                "search": " ".join(
                    str(row.get(k) or "") for k in ("fio", "program", "position")
                ).lower(),
            })
        return result

    app.jinja_env.globals["worker_options"] = worker_options
    app.jinja_env.globals["instructor_options"] = instructor_options

    # -- доступ -----------------------------------------------------------
    @app.before_request
    def _check_local_access():
        if app.config.get("TESTING"):
            return None
        host = (request.host or "").split(":")[0]
        if host not in {"127.0.0.1", "localhost", "::1"}:
            abort(403)
        if request.method in {"POST", "PUT", "PATCH", "DELETE"}:
            token = session.get("csrf")
            sent = request.form.get("_csrf") or request.headers.get("X-CSRF-Token")
            if not token or not sent or not hmac.compare_digest(str(token), str(sent)):
                abort(400, "Форма устарела. Обновите страницу и повторите действие.")
        return None

    @app.before_request
    def _require_start_token():
        if app.config.get("TESTING"):
            return None
        # Проверка доступности нужна ярлыку до авторизации: она не раскрывает
        # данных и используется, чтобы не запускать второй экземпляр.
        if request.path in ("/healthz", "/favicon.ico"):
            return None
        if session.get("authorized"):
            return None
        supplied = request.args.get(START_TOKEN_PARAM, "")
        expected = app.config["OTR_START_TOKEN"]
        if supplied and hmac.compare_digest(supplied, expected):
            session["authorized"] = True
            session["csrf"] = secrets.token_urlsafe(24)
            return redirect(request.path)
        return (
            "Приложение запускается через ярлык. Откройте приложение заново.",
            403,
        )

    @app.context_processor
    def _inject():
        cat = catalog()
        org_uuid = current_org_uuid()
        entry = cat.get(org_uuid) if org_uuid else None

        brand_color = ""
        brand_text = "#ffffff"
        nav_items: list[navigation.NavItem] = []
        org_heading = ""

        if entry is not None:
            # Цвет и порядок закладок хранятся в книге, но шапка должна
            # работать даже при недоступном файле — читаем осторожно.
            try:
                book = load_book(entry)
                brand_color = navigation.normalize_color(
                    book.org_value("brand_color", entry.color or "#1F4E79")
                )
                nav_items = navigation.ordered_items(book.org_value("nav_order"))
                org_heading = book.org_value("full_name") or book.org_value("name") or entry.name
            except Exception:
                brand_color = navigation.normalize_color(entry.color or "#1F4E79")
                nav_items = navigation.ordered_items("")
                org_heading = entry.name
            brand_text = navigation.contrast_text(brand_color)

        return {
            "csrf_token": session.get("csrf", ""),
            "organizations": cat.entries,
            "current_org": entry,
            "brand_color": brand_color,
            "brand_text": brand_text,
            "nav_items": nav_items,
            "org_heading": org_heading,
            "active_nav": navigation.active_key(request),
            "flash_errors": [],
        }

    # -- главная ----------------------------------------------------------
    @app.get("/")
    def index():
        cat = catalog()
        summaries = []
        for entry in cat.entries:
            item = {"entry": entry, "error": "", "counts": None}
            try:
                book = load_book(entry)
                summary = _book_summary(book)
                item["counts"] = summary["counts"]
                item["total_planned"] = sum(
                    1 for r in book.table("Инструктажи").rows if r.get("status") == "Запланирован"
                )
            except AppError as exc:
                item["error"] = exc.message
            except OSError:
                item["error"] = "Файл книги недоступен."
            summaries.append(item)
        return render_template("index.html", summaries=summaries)

    @app.post("/organizations/select")
    def select_org():
        org_uuid = request.form.get("org_uuid", "")
        cat = catalog()
        entry = cat.get(org_uuid)
        if entry is None:
            flash("Организация не найдена.", "error")
            return redirect(url_for("index"))
        session["org_uuid"] = org_uuid
        cat.touch(org_uuid)
        return redirect(url_for("dashboard"))

    @app.route("/organizations/new", methods=["GET", "POST"])
    def new_org():
        if request.method == "GET":
            return render_template("org_new.html", form={}, error=None)

        form = request.form.to_dict()
        name = (form.get("name") or "").strip()
        if not name:
            return render_template(
                "org_new.html", form=form, error="Укажите краткое название организации."
            )
        from .organizations import prepare_org_dir

        cat = catalog()
        folder, book_path = prepare_org_dir(base_dir=base / "organizations")
        try:
            book = __import__("ot_registry.excel_store", fromlist=["create_book"]).create_book(
                book_path,
                data_dir=folder,
                name=name,
                full_name=(form.get("full_name") or name).strip(),
                inn=(form.get("inn") or "").strip(),
                kpp=(form.get("kpp") or "").strip(),
                ogrn=(form.get("ogrn") or "").strip(),
                address=(form.get("address") or "").strip(),
                director=(form.get("director") or "").strip(),
                ot_specialist=(form.get("ot_specialist") or "").strip(),
                with_demo_data=bool(form.get("with_demo_data")),
            )
        except AppError as exc:
            return render_template("org_new.html", form=form, error=exc.message)

        entry = cat.add(name, book_path, org_uuid=book.org_uuid)
        session["org_uuid"] = entry.uuid
        flash("Организация создана. Книга Excel готова к работе.", "ok")
        return redirect(url_for("dashboard"))

    @app.route("/organizations/connect", methods=["GET", "POST"])
    def connect_org():
        if request.method == "GET":
            return render_template("org_connect.html", form={}, error=None)
        form = request.form.to_dict()
        raw_path = Path((form.get("path") or "").strip())
        if not raw_path.exists() or not raw_path.is_file():
            return render_template(
                "org_connect.html", form=form,
                error="Файл не найден. Укажите полный путь к книге .xlsx.",
            )
        cat = catalog()
        resolved = raw_path.resolve()
        if str(resolved) in cat.existing_paths():
            return render_template(
                "org_connect.html", form=form,
                error="Эта книга уже подключена.",
            )
        book, error = safe(open_book, resolved, data_dir=resolved.parent)
        if error:
            return render_template("org_connect.html", form=form, error=error["message"])
        existing = cat.by_book_uuid(book.org_uuid)
        if existing:
            return render_template(
                "org_connect.html", form=form,
                error=f"Эта книга уже подключена как «{existing.name}». "
                      "Копия книги считается той же организацией.",
            )
        entry = cat.add(book.name or resolved.stem, resolved, org_uuid=book.org_uuid)
        session["org_uuid"] = entry.uuid
        flash("Книга подключена.", "ok")
        return redirect(url_for("dashboard"))

    # -- рабочие экраны ---------------------------------------------------
    @app.get("/org")
    def dashboard():
        cat, entry = require_org() if len(catalog()) else (None, None)
        if entry is None:
            return redirect(url_for("index"))
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)

        # Плановая отправка резервной копии: раз в N дней при открытии.
        # Ошибка почты не мешает работе — только предупреждение.
        email_note = ""
        try:
            mail_result = backups.run_scheduled(book)
            if mail_result.get("ran"):
                email_note = (
                    "Резервная копия отправлена на почту."
                    if mail_result.get("sent")
                    else f"Не удалось отправить копию на почту: {mail_result.get('error', '')}"
                )
        except Exception:
            email_note = ""

        summary = _book_summary(book)
        visible = [
            (row, info) for row, info in summary["rows"]
            if info.level in (
                deadlines.DueLevel.OVERDUE, deadlines.DueLevel.TODAY,
                deadlines.DueLevel.SOON, deadlines.DueLevel.UPCOMING,
            )
        ]
        visible.sort(key=lambda pair: deadlines.sort_key(pair[1].level, pair[1].days))
        workers = services.workers(book)
        instructor_rows = services.instructors(book)
        expiring = [
            r for r in instructor_rows
            if r.get("valid_to") and isinstance(r.get("valid_to"), date)
            and (r["valid_to"] - config.today()).days <= 30
        ]
        return render_template(
            "dashboard.html",
            entry=entry,
            book=book,
            summary=summary,
            due_rows=visible[:40],
            workers=workers,
            instructors=instructor_rows,
            expiring=expiring,
            counts=summary["counts"],
            email_note=email_note,
        )

    @app.get("/trainings")
    def trainings_list():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        summary = _book_summary(book)

        # По умолчанию показываются запланированные: с ними работают чаще всего.
        status = request.args.get("status", "Запланирован")
        search = (request.args.get("q") or "").strip().lower()
        sort_key = request.args.get("sort", "plan")

        rows = []
        for row, info in summary["rows"]:
            if status and str(row.get("status") or "") != status:
                continue
            if search and search not in " ".join(
                str(row.get(k) or "") for k in ("fio", "type", "code", "number_ot", "program")
            ).lower():
                continue
            rows.append((row, info))

        def sort_default(pair):
            row = pair[0]
            if sort_key == "plan":
                return (row.get("plan_date") or date.max, str(row.get("fio") or ""))
            if sort_key == "fact":
                return (row.get("fact_date") or date.max, str(row.get("fio") or ""))
            if sort_key == "fio":
                return (str(row.get("fio") or "").lower(), row.get("plan_date") or date.max)
            return (row.get("plan_date") or date.max,)

        rows.sort(key=sort_default)
        return render_template(
            "trainings.html", entry=entry, book=book, rows=rows, summary=summary,
            status=status, search=search, sort=sort_key,
        )

    @app.route("/trainings/new", methods=["GET", "POST"])
    def training_new():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)

        workers_rows = services.workers(book)
        instructor_rows = services.instructors(book, on=config.today())
        dicts = services.dictionaries(book)
        program_rows = services.programs(book)

        def render(form, errors=None, message="", details=None):
            # Подписи для полей с поиском: если ID уже выбран, показываем текст.
            worker_label = ""
            selected_worker = services.worker_by_id(book, form.get("worker_id", "") or "")
            if selected_worker:
                worker_label = str(selected_worker.get("fio") or "")
                if selected_worker.get("position"):
                    worker_label += f" — {selected_worker['position']}"
            instructor_label = ""
            chosen = services.instructor_by_id(book, form.get("instructor_id", "") or "")
            if chosen:
                instructor_label = str(chosen.get("fio") or "")
                if chosen.get("program"):
                    instructor_label += f" — {chosen['program']}"
            return render_template(
                "training_new.html", entry=entry, book=book,
                workers=workers_rows, instructors=instructor_rows, dicts=dicts,
                program_titles=[services.program_title(r) for r in program_rows],
                form=form, errors=errors or {}, message=message, details=details,
                selected_worker_label=worker_label,
                selected_instructor_label=instructor_label,
            )

        if request.method == "GET":
            worker_id = request.args.get("worker_id", "")
            return render(
                {"worker_id": worker_id, "plan_date": config.today().isoformat()}
            )

        form = request.form.to_dict()
        # Поле поиска присылает скрытый ID; если его нет, пробуем распознать
        # введённый текст как ФИО (удобно при быстром вводе).
        if not form.get("worker_id"):
            typed = (form.get("worker_search") or "").strip().lower()
            if typed:
                found = [
                    w for w in workers_rows
                    if typed in " ".join(
                        str(w.get(k) or "") for k in ("fio", "position", "department", "personnel_no")
                    ).lower()
                ]
                if len(found) == 1:
                    form["worker_id"] = str(found[0]["id"])
                elif len(found) > 1:
                    # Несколько совпадений: выбрать за оператора нельзя.
                    return render(
                        form,
                        {"worker_id": f"По запросу «{typed}» найдено несколько работников. "
                                      "Уточните фамилию или выберите строку из списка."},
                    )
        if not form.get("instructor_id"):
            typed = (form.get("instructor_search") or "").strip().lower()
            if typed:
                found = [
                    i for i in instructor_rows
                    if typed in " ".join(
                        str(i.get(k) or "") for k in ("fio", "program", "position")
                    ).lower()
                ]
                if len(found) == 1:
                    form["instructor_id"] = str(found[0]["id"])
                elif len(found) > 1:
                    names = ", ".join(
                        f"{i.get('fio')} ({i.get('program') or 'без программы'})"
                        for i in found[:4]
                    )
                    return render(
                        form,
                        {"instructor_id":
                            f"По запросу «{typed}» найдено несколько допусков: {names}. "
                            "Уточните программу или выберите строку из списка."},
                    )

        worker_ids = request.form.getlist("worker_ids") or (
            [form["worker_id"]] if form.get("worker_id") else []
        )
        payload = {
            "worker_id": worker_ids[0] if len(worker_ids) == 1 else "",
            "instructor_id": form.get("instructor_id", ""),
            "plan_date": None,
            "type": form.get("type", ""),
            "program": form.get("program", ""),
            "reason": form.get("reason", ""),
            "note": form.get("note", ""),
        }
        try:
            payload["plan_date"] = _parse_date(form.get("plan_date"))
        except ValidationError as exc:
            return render(form, {"plan_date": exc.message})

        try:
            if len(worker_ids) > 1:
                created = services.plan_batch(
                    book,
                    worker_ids=worker_ids,
                    instructor_id=payload["instructor_id"],
                    plan_date=payload["plan_date"],
                    training_type=payload["type"],
                    program=payload["program"],
                    reason=payload["reason"],
                )
                flash(f"Создано планов: {len(created)}.", "ok")
            else:
                payload["worker_id"] = worker_ids[0] if worker_ids else ""
                created_row = services.plan_training(book, payload)
                flash(
                    f"План создан. Код бланка: {validation.format_code(created_row['code'])}.",
                    "ok",
                )
            return redirect(url_for("trainings_list"))
        except ValidationError as exc:
            return render(form, exc.fields, exc.message, exc.details)
        except AppError as exc:
            return render(form, {}, exc.message, exc.details)

    @app.route("/register", methods=["GET", "POST"])
    def register():
        """Регистрация факта по коду бланка."""
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)

        if request.method == "GET":
            code = request.args.get("code", "")
            return render_template(
                "register.html", entry=entry, book=book, code=code,
                training=None, errors={}, message="",
            )

        form = request.form.to_dict()
        code = form.get("code", "")
        action = form.get("action", "find")
        if action == "find":
            try:
                normalized = validation.normalize_code(code)
                if validation.is_manual_code(normalized):
                    return redirect(url_for("manual_entry", code=normalized))
                row = registration.find_active_by_code(book, code)
            except AppError as exc:
                return render_template(
                    "register.html", entry=entry, book=book, code=code,
                    training=None, errors={}, message=exc.message, details=exc.details,
                )
            return render_template(
                "register.html", entry=entry, book=book, code=code,
                training=row, errors={}, message="",
                instructors=services.instructors(book, on=config.today()),
                fact_date=config.today(),
            )

        training_id = form.get("training_id", "")
        payload = {
            "fact_date": None,
            "instructor_id": form.get("instructor_id", ""),
            "paper_confirmed": bool(form.get("paper_confirmed")),
            "note": form.get("note", ""),
        }
        try:
            payload["fact_date"] = _parse_date(form.get("fact_date"))
        except ValidationError as exc:
            payload["fact_date"] = None
            row = book.table("Инструктажи").by_key("id", training_id) or {}
            return render_template(
                "register.html", entry=entry, book=book, code=code, training=row,
                errors={"fact_date": exc.message}, message="",
                instructors=services.instructors(book, on=config.today()),
                fact_date=None,
            )

        result, error = safe(registration.register_training, book, training_id, payload)
        if error:
            row = book.table("Инструктажи").by_key("id", training_id) or {}
            return render_template(
                "register.html", entry=entry, book=book, code=code, training=row,
                errors=error["fields"], message=error["message"],
                details=error.get("details"),
                instructors=services.instructors(book, on=config.today()),
                fact_date=payload["fact_date"],
            )
        trainer = services.instructor_by_id(
            book, str(result["training"].get("instructor_id") or "")
        )
        return render_template(
            "register_done.html", entry=entry, book=book,
            training=result["training"], next_plan=result.get("next_plan"),
            instructor_position=str((trainer or {}).get("program") or ""),
        )

    @app.route("/manual", methods=["GET", "POST"])
    def manual_entry():
        """Внесение бумажного бланка, заполненного вне системы."""
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)

        workers = services.workers(book)
        instructor_rows = services.instructors(book, on=config.today())
        dicts = services.dictionaries(book)

        def render_manual(form, errors=None, message="", details=None, plans_list=None):
            selected_worker_label = ""
            chosen_worker = services.worker_by_id(book, form.get("worker_id", "") or "")
            if chosen_worker:
                selected_worker_label = str(chosen_worker.get("fio") or "")
                if chosen_worker.get("position"):
                    selected_worker_label += f" — {chosen_worker['position']}"
            return render_template(
                "manual.html", entry=entry, book=book, workers=workers,
                instructors=instructor_rows, dicts=dicts,
                form=form, errors=errors or {}, message=message, details=details,
                plans=plans_list or [],
                selected_worker_label=selected_worker_label,
            )

        if request.method == "GET":
            return render_manual({"fact_date": config.today().isoformat()})

        form = request.form.to_dict()
        try:
            fact = _parse_date(form.get("fact_date"))
        except ValidationError as exc:
            form["fact_date"] = ""
            return render_manual(form, {"fact_date": exc.message})
        payload = {
            "worker_id": form.get("worker_id", ""),
            "type": form.get("type", ""),
            "program": form.get("program", ""),
            "reason": form.get("reason", ""),
            "fact_date": fact,
            "instructor_id": form.get("instructor_id", ""),
            "paper_confirmed": bool(form.get("paper_confirmed")),
            "manual_mark": form.get("manual_mark", ""),
            "note": form.get("note", ""),
            "plan_id": form.get("plan_id", ""),
        }
        plans = []
        if payload["worker_id"] and payload["type"]:
            plans = registration.suggest_link_candidates(
                book, worker_id=payload["worker_id"], training_type=payload["type"]
            )

        result, error = safe(
            registration.register_manual, book, payload,
            allow_similar=bool(form.get("allow_similar")),
        )
        if error:
            return render_manual(form, error["fields"], error["message"],
                                 error.get("details") or [], plans)
        return render_template(
            "manual_done.html", entry=entry, book=book,
            training=result["training"], next_plan=result.get("next_plan"),
        )

    @app.get("/workers")
    def workers_list():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        search = (request.args.get("q") or "").strip().lower()
        rows = services.workers(book, only_active=False)
        if search:
            rows = [
                r for r in rows
                if search in " ".join(
                    str(r.get(k) or "") for k in ("fio", "position", "department", "personnel_no")
                ).lower()
            ]
        return render_template("workers.html", entry=entry, book=book, rows=rows, search=search)

    @app.route("/workers/new", methods=["GET", "POST"])
    def worker_new():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        dicts = services.dictionaries(book)
        if request.method == "GET":
            return render_template("worker_new.html", entry=entry, book=book, dicts=dicts,
                                   form={}, errors={}, message="")

        form = request.form.to_dict()
        data = {
            "fio": form.get("fio", ""),
            "personnel_no": form.get("personnel_no", ""),
            "position": form.get("position", ""),
            "department": form.get("department", ""),
            "manager": form.get("manager", ""),
            "status": "Работает",
            "period_months": form.get("period_months") or None,
            "note": form.get("note", ""),
            "hire_date": None,
            "birth_date": None,
        }
        try:
            data["hire_date"] = _parse_date(form.get("hire_date"))
            data["birth_date"] = _parse_date(form.get("birth_date"))
        except ValidationError as exc:
            return render_template("worker_new.html", entry=entry, book=book, dicts=dicts,
                                   form=form, errors={"hire_date": exc.message}, message="")

        row, error = safe(
            services.add_worker, book, data,
            allow_duplicate=bool(form.get("allow_duplicate")),
        )
        if error:
            return render_template("worker_new.html", entry=entry, book=book, dicts=dicts,
                                   form=form, errors=error["fields"],
                                   message=error["message"],
                                   details=error.get("details"))
        flash(f"Работник добавлен: {row['fio']}.", "ok")
        return redirect(url_for("worker_card", worker_id=row["id"]))

    @app.get("/workers/<worker_id>")
    def worker_card(worker_id: str):
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        worker = services.worker_by_id(book, worker_id)
        if worker is None:
            flash("Работник не найден.", "error")
            return redirect(url_for("workers_list"))
        summary = _book_summary(book)
        rows = [(r, i) for r, i in summary["rows"] if r.get("worker_id") == worker_id]
        rows.sort(key=lambda p: p[0].get("plan_date") or date.min, reverse=True)
        history = services.worker_history(book, worker_id)
        return render_template("worker_card.html", entry=entry, book=book, worker=worker,
                               rows=rows, history=history)

    @app.post("/workers/<worker_id>/transfer")
    def worker_transfer(worker_id: str):
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        form = request.form.to_dict()
        _, error = safe(
            services.transfer_worker, book, worker_id,
            position=form.get("position", ""),
            department=form.get("department", ""),
            manager=form.get("manager", ""),
            reason=form.get("reason", ""),
        )
        if error:
            flash(error["message"], "error")
        else:
            flash("Перевод сохранён.", "ok")
        return redirect(url_for("worker_card", worker_id=worker_id))

    @app.post("/workers/<worker_id>/dismiss")
    def worker_dismiss(worker_id: str):
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        form = request.form.to_dict()
        try:
            dismiss_date = _parse_date(form.get("dismiss_date")) or config.today()
        except ValidationError as exc:
            flash(exc.message, "error")
            return redirect(url_for("worker_card", worker_id=worker_id))
        _, error = safe(services.dismiss_worker, book, worker_id,
                        dismiss_date=dismiss_date, reason=form.get("reason", ""))
        flash("Работник уволен." if not error else error["message"],
              "ok" if not error else "error")
        return redirect(url_for("worker_card", worker_id=worker_id))

    @app.get("/instructors")
    def instructors_list():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)

        state = request.args.get("state", "active")
        search = (request.args.get("q") or "").strip().lower()
        include_revoked = state in ("all", "revoked")
        rows = services.instructors(book, include_revoked=include_revoked)
        if state == "revoked":
            rows = [r for r in rows if str(r.get("status") or "") == "Аннулирован"]
        if search:
            rows = [
                r for r in rows
                if search in " ".join(
                    str(r.get(k) or "") for k in ("fio", "position", "program", "types")
                ).lower()
            ]

        deputy_names = {}
        for row in book.table("Инструктирующие").rows:
            deputy_id = str(row.get("deputy_id") or "")
            if deputy_id:
                deputy = book.table("Инструктирующие").by_key("id", deputy_id)
                if deputy:
                    deputy_names[str(row.get("id"))] = f"{deputy.get('fio')}"

        return render_template(
            "instructors.html", entry=entry, book=book, rows=rows,
            state=state, search=search,
            deputy_names=deputy_names,
            revoke_reasons=services.instructor_revoke_reasons(book),
            without_deputy=services.instructors_without_deputy(book) if state == "active" else [],
            dicts=services.dictionaries(book),
            warning=request.args.get("warning", ""),
        )

    @app.route("/instructors/new", methods=["GET", "POST"])
    def instructor_new():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)

        dicts = services.dictionaries(book)
        program_rows = services.programs(book)
        people = services.instructors(book, include_revoked=True)

        if request.method == "GET":
            return render_template(
                "instructor_new.html", entry=entry, book=book, dicts=dicts,
                programs=program_rows, people=people, form={}, errors={}, message="",
            )

        form = request.form.to_dict()
        data = {
            "fio": form.get("fio", ""),
            "position": form.get("position", ""),
            "program": form.get("program", ""),
            "types": ",".join(request.form.getlist("types")) or form.get("types", ""),
            "valid_from": None,
            "valid_to": None,
            "note": form.get("note", ""),
        }
        try:
            data["valid_from"] = _parse_date(form.get("valid_from"))
            data["valid_to"] = _parse_date(form.get("valid_to"))
        except ValidationError as exc:
            return render_template(
                "instructor_new.html", entry=entry, book=book, dicts=dicts,
                programs=program_rows, people=people, form=form,
                errors={"valid_to": exc.message}, message="",
            )
        row, error = safe(services.add_instructor, book, data)
        if error:
            return render_template(
                "instructor_new.html", entry=entry, book=book, dicts=dicts,
                programs=program_rows, people=people, form=form,
                errors=error["fields"], message=error["message"],
                details=error.get("details"),
            )
        flash(f"Допуск добавлен: {row['fio']}.", "ok")
        return redirect(url_for("instructors_list"))

    @app.post("/instructors/<row_id>/revoke")
    def instructor_revoke(row_id: str):
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        form = request.form.to_dict()
        try:
            revoke_date = _parse_date(form.get("revoke_date")) or config.today()
        except ValidationError:
            revoke_date = config.today()
        _, error = safe(services.revoke_instructor, book, row_id,
                        reason=form.get("reason", ""), revoke_date=revoke_date)
        flash("Допуск аннулирован." if not error else error["message"],
              "ok" if not error else "error")
        return redirect(url_for("instructors_list"))

    @app.post("/instructors/<row_id>/restore")
    def instructor_restore(row_id: str):
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        _, error = safe(services.restore_instructor, book, row_id)
        flash("Допуск восстановлен." if not error else error["message"],
              "ok" if not error else "error")
        return redirect(url_for("instructors_list", state="all"))

    @app.post("/instructors/<row_id>/deputy")
    def instructor_deputy(row_id: str):
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        form = request.form.to_dict()
        _, error = safe(services.set_deputy, book, row_id, form.get("deputy_id", ""))
        flash("Заместитель назначен." if not error else error["message"],
              "ok" if not error else "error")
        return redirect(url_for("instructors_list"))

    # -- инструкции и программы -------------------------------------------
    @app.get("/programs")
    def programs_list():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        show_all = request.args.get("all") == "1"
        search = (request.args.get("q") or "").strip().lower()
        rows = services.programs(book, only_active=not show_all)
        if search:
            rows = [
                r for r in rows
                if search in " ".join(
                    str(r.get(k) or "") for k in ("number", "title", "description", "approved_by")
                ).lower()
            ]
        return render_template(
            "programs.html", entry=entry, book=book, rows=rows,
            show_all=show_all, search=search,
            active_count=len(services.programs(book)),
        )

    @app.route("/programs/new", methods=["GET", "POST"])
    def program_new():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)

        if request.method == "GET":
            return render_template("program_form.html", entry=entry, book=book,
                                   form={}, errors={}, message="", row=None)

        form = request.form.to_dict()
        data = {
            "number": form.get("number", ""),
            "title": form.get("title", ""),
            "description": form.get("description", ""),
            "approved_by": form.get("approved_by", ""),
            "approved_at": None,
            "status": form.get("status", "Действует"),
            "replaced_by": form.get("replaced_by", ""),
            "note": form.get("note", ""),
        }
        try:
            data["approved_at"] = _parse_date(form.get("approved_at"))
        except ValidationError as exc:
            return render_template("program_form.html", entry=entry, book=book,
                                   form=form, errors={"approved_at": exc.message},
                                   message="", row=None)
        row, error = safe(services.add_program, book, data)
        if error:
            return render_template("program_form.html", entry=entry, book=book,
                                   form=form, errors=error["fields"],
                                   message=error["message"], row=None)
        flash("Инструкция добавлена.", "ok")
        return redirect(url_for("programs_list"))

    @app.route("/programs/<program_id>/edit", methods=["GET", "POST"])
    def program_edit(program_id: str):
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        row = book.table("Инструкции").by_key("id", program_id)
        if row is None:
            flash("Инструкция не найдена.", "error")
            return redirect(url_for("programs_list"))

        if request.method == "GET":
            return render_template("program_form.html", entry=entry, book=book,
                                   form={}, errors={}, message="", row=row)

        form = request.form.to_dict()
        data = {
            "number": form.get("number", ""),
            "title": form.get("title", ""),
            "description": form.get("description", ""),
            "approved_by": form.get("approved_by", ""),
            "approved_at": None,
            "status": form.get("status", "Действует"),
            "replaced_by": form.get("replaced_by", ""),
            "note": form.get("note", ""),
        }
        try:
            data["approved_at"] = _parse_date(form.get("approved_at"))
        except ValidationError as exc:
            return render_template("program_form.html", entry=entry, book=book,
                                   form=form, errors={"approved_at": exc.message},
                                   message="", row=row)
        _, error = safe(services.update_program, book, program_id, data)
        if error:
            return render_template("program_form.html", entry=entry, book=book,
                                   form=form, errors=error["fields"],
                                   message=error["message"], row=row)
        flash("Изменения сохранены.", "ok")
        return redirect(url_for("programs_list"))

    @app.post("/programs/<program_id>/deactivate")
    def program_deactivate(program_id: str):
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        _, error = safe(services.deactivate_program, book, program_id,
                        reason=request.form.get("reason", ""))
        flash("Инструкция переведена в состояние «Не действует»." if not error else error["message"],
              "ok" if not error else "error")
        return redirect(url_for("programs_list"))

    @app.route("/print", methods=["GET", "POST"])
    def print_forms():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)

        if request.method == "GET":
            summary = _book_summary(book)
            planned_rows = [(r, i) for r, i in summary["rows"] if r.get("status") == "Запланирован"]
            planned_rows.sort(key=lambda p: p[0].get("plan_date") or date.min)
            return render_template(
                "print.html", entry=entry, book=book,
                rows=planned_rows, result=None, mode="filled",
                templates=docx_templates.discover(),
                templates_dir=str(docx_templates.templates_dir()),
                using_template=docx_templates.pick_template(
                    config.CARDS_PER_PAGE) is not None,
            )

        mode = request.form.get("mode", "filled")
        try:
            per_page = int(request.form.get("per_page") or config.CARDS_PER_PAGE)
        except ValueError:
            per_page = config.CARDS_PER_PAGE
        per_page = max(1, min(per_page, config.MAX_CARDS_PER_PAGE))

        if mode == "blank":
            training_type = request.form.get("type", "")
            count = request.form.get("count", "1")
            try:
                count_int = int(count)
            except ValueError:
                count_int = 1
            result, error = safe(
                printing_template.print_blank_template, book,
                training_type=training_type, count=count_int,
                cards_per_page=per_page,
            )
            if result is None and not error:
                # Шаблона нет — печатаем встроенный бланк.
                result, error = safe(
                    printing.print_blank_grid, book,
                    training_type=training_type, count=count_int, per_page=per_page,
                )
        else:
            ids = request.form.getlist("training_ids")
            result, error = safe(
                printing_template.print_filled_template, book, ids,
                cards_per_page=per_page,
            )
            if result is None and not error:
                result, error = safe(
                    printing.print_filled_grid, book, ids, per_page=per_page,
                )

        if error:
            summary = _book_summary(book)
            planned_rows = [(r, i) for r, i in summary["rows"] if r.get("status") == "Запланирован"]
            return render_template("print.html", entry=entry, book=book, rows=planned_rows,
                                   result=None, mode=mode, message=error["message"])
        return render_template("print.html", entry=entry, book=book, rows=[], result=result,
                               mode=mode, message="")

    @app.get("/download")
    def download():
        """Выдать сформированный файл по имени из папки форм организации."""
        cat, entry = require_org()
        name = request.args.get("name", "")
        folder = (entry.resolve(base).parent / "forms").resolve()
        target = (folder / name).resolve()
        if folder not in target.parents or not target.exists():
            abort(404)
        return send_file(target, as_attachment=True, download_name=target.name)

    @app.get("/archive")
    def archive():
        """Архив: дела и пакеты с количеством документов внутри."""
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)

        # По умолчанию — только открытые: закрытые пакеты нужны реже.
        show_closed = request.args.get("closed") == "1"
        rows = reports.pack_rows(book)
        if not show_closed:
            rows = [r for r in rows if str(r.get("status") or "") == "Открыт"]

        cases = reports.case_rows(book)
        open_count = len([r for r in reports.pack_rows(book) if r.get("status") == "Открыт"])
        closed_count = len([r for r in reports.pack_rows(book) if r.get("status") == "Закрыт"])

        return render_template(
            "archive.html", entry=entry, book=book, rows=rows, cases=cases,
            show_closed=show_closed, open_count=open_count, closed_count=closed_count,
        )

    @app.post("/archive/close-pack")
    def close_pack():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        form = request.form.to_dict()
        _, error = safe(services.close_pack, book, form.get("case_id", ""),
                        form.get("pack_id", ""), close_date=config.today(),
                        reason=form.get("reason", ""))
        flash("Пакет закрыт." if not error else error["message"],
              "ok" if not error else "error")
        return redirect(url_for("archive"))

    # -- отчёты ------------------------------------------------------------
    @app.get("/trainings/<training_id>")
    def training_card(training_id: str):
        """Карточка инструктажа: данные, действия, предупреждения."""
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        data, error = safe(services.training_card_data, book, training_id)
        if error:
            flash(error["message"], "error")
            return redirect(url_for("trainings_list"))
        return render_template(
            "training_card.html", entry=entry, book=book,
            available_instructors=services.instructors(book),
            program_titles=[services.program_title(r) for r in services.programs(book)],
            worker_id=str(data["training"].get("worker_id") or ""),
            errors={},
            **data,
        )

    @app.post("/trainings/<training_id>/update")
    def training_update(training_id: str):
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        form = request.form.to_dict()
        try:
            plan_date = _parse_date(form.get("plan_date"))
        except ValidationError as exc:
            data, _ = safe(services.training_card_data, book, training_id)
            return render_template(
                "training_card.html", entry=entry, book=book,
                available_instructors=services.instructors(book),
                program_titles=[services.program_title(r) for r in services.programs(book)],
                worker_id=str((data or {}).get("training", {}).get("worker_id") or ""),
                errors={"plan_date": exc.message},
                **(data or {}),
            )
        payload = {
            "plan_date": plan_date,
            "type": form.get("type", ""),
            "instructor_id": form.get("instructor_id", ""),
            "program": form.get("program", ""),
            "reason": form.get("reason", ""),
        }
        _, error = safe(services.update_plan, book, training_id, payload,
                        reason=form.get("change_reason", ""))
        if error:
            flash(error["message"], "error")
        else:
            flash("Изменения сохранены.", "ok")
        return redirect(url_for("training_card", training_id=training_id))

    @app.post("/trainings/<training_id>/cancel")
    def training_cancel(training_id: str):
        """Отмена инструктажа с обязательной причиной."""
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        form = request.form.to_dict()
        try:
            cancel_date = _parse_date(form.get("cancel_date")) or config.today()
        except ValidationError:
            cancel_date = config.today()
        _, error = safe(services.cancel_training, book, training_id,
                        reason=form.get("reason", ""), cancel_date=cancel_date)
        flash("Инструктаж отменён." if not error else error["message"],
              "ok" if not error else "error")
        if error:
            return redirect(url_for("training_card", training_id=training_id))
        return redirect(url_for("trainings_list"))

    @app.get("/reports")
    def reports_page():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        summary = _book_summary(book)
        return render_template(
            "reports.html",
            entry=entry,
            book=book,
            counts=summary["counts"],
            incomplete=reports.incomplete_registrations(book),
            repeat_soon=reports.repeat_due_soon(book, days=30),
            expiring=reports.instructors_expiring(book, days=30),
            summary=reports.by_type_summary(book),
            packs=reports.pack_rows(book),
            year=config.today().year,
        )

    @app.get("/reports/pack/<case_id>/<pack_id>")
    def pack_registry(case_id: str, pack_id: str):
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        rows = reports.pack_registry(book, case_id=case_id, pack_id=pack_id)
        return render_template(
            "pack_registry.html", entry=entry, book=book,
            rows=rows, case_id=case_id, pack_id=pack_id,
        )

    # -- ручной режим Excel ----------------------------------------------
    def _excel_context(book, entry, **extra):
        state = manual_mode.read_state(book)
        backups_dir = book.backups_dir()
        backups = []
        if backups_dir.exists():
            for path in sorted(backups_dir.glob("*.xlsx"), reverse=True)[:15]:
                stat = path.stat()
                backups.append({
                    "name": path.name,
                    "size": f"{stat.st_size / 1024:.0f} КБ",
                    "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%d.%m.%Y %H:%M"),
                })
        context = {
            "entry": entry,
            "book": book,
            "state": state,
            "protection_password": book.protection_password,
            "backups": backups,
            "backup_count": len(list(backups_dir.glob("*.xlsx"))) if backups_dir.exists() else 0,
            "last_backup": backups[0]["name"] if backups else "",
            "keep": config.DEFAULT_BACKUP_KEEP,
            "report": None,
            "message": "",
            "message_kind": "info",
        }
        context.update(extra)
        return context

    @app.get("/excel")
    def excel_page():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        return render_template("excel.html", **_excel_context(book, entry))

    @app.route("/excel/open", methods=["GET", "POST"])
    def excel_open():
        """Включить ручной режим: контрольная копия, снимок, блокировка записи."""
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        result, error = safe(manual_mode.open_for_manual, book)
        if error:
            return render_template(
                "excel.html",
                **_excel_context(book, entry, message=error["message"], message_kind="error"),
            )
        launched = manual_mode.launch_excel(book)
        return render_template(
            "excel.html",
            **_excel_context(
                book, entry,
                message=("Ручной режим включён. Запись через приложение заблокирована. "
                         + launched),
                message_kind="ok",
            ),
        )

    @app.post("/excel/check")
    def excel_check():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        report, error = safe(manual_mode.check_book, book)
        if error:
            return render_template(
                "excel.html",
                **_excel_context(book, entry, message=error["message"], message_kind="error"),
            )
        return render_template("excel.html", **_excel_context(book, entry, report=report))

    @app.post("/excel/accept")
    def excel_accept():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        result, error = safe(manual_mode.accept_changes, book,
                             reason=request.form.get("reason", ""))
        if error:
            report, _ = safe(manual_mode.check_book, book)
            return render_template(
                "excel.html",
                **_excel_context(book, entry, report=report,
                                 message=error["message"], message_kind="error"),
            )
        fresh = open_book(book.path, data_dir=book.path.parent)
        assigned = result.get("assigned", 0)
        text = "Изменения приняты, снимок обновлён, запись через приложение возобновлена."
        if assigned:
            text += f" Новым строкам назначено ID: {assigned}."
        return render_template(
            "excel.html",
            **_excel_context(fresh, entry, message=text, message_kind="ok"),
        )

    @app.post("/excel/cancel")
    def excel_cancel():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        restore = bool(request.form.get("restore"))
        message, error = safe(manual_mode.cancel_manual, book, restore=restore)
        if error:
            return render_template(
                "excel.html",
                **_excel_context(book, entry, message=error["message"], message_kind="error"),
            )
        fresh, open_error = safe(open_book, book.path, data_dir=book.path.parent)
        target = fresh if fresh else book
        return render_template(
            "excel.html",
            **_excel_context(target, entry, message=message, message_kind="ok"),
        )

    @app.post("/excel/migrate")
    def excel_migrate():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        result, error = safe(migrations.migrate, book.path)
        if error:
            return render_template(
                "excel.html",
                **_excel_context(book, entry, message=error["message"], message_kind="error"),
            )
        fresh = open_book(book.path, data_dir=book.path.parent)
        text = (f"Схема обновлена: {result.from_version} → {result.to_version}. "
                f"Создан резерв: {Path(result.backup).name if result.backup else '—'}.")
        return render_template(
            "excel.html",
            **_excel_context(fresh, entry, message=text, message_kind="ok"),
        )

    @app.get("/settings")
    def settings_page():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        return render_template(
            "settings.html",
            entry=entry,
            book=book,
            nav_items=navigation.ordered_items(book.org_value("nav_order")),
            nav_full=[navigation.ITEM_BY_KEY[k] for k in navigation.DEFAULT_ORDER],
            brand_colors=navigation.BRAND_COLORS,
            current_color=navigation.normalize_color(
                book.org_value("brand_color", entry.color or "#1F4E79")
            ),
        )

    @app.post("/settings")
    def settings_save():
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        form = request.form.to_dict()
        for key in ("full_name", "inn", "kpp", "ogrn", "address", "director", "ot_specialist"):
            book.set_org_value(key, form.get(key, ""))
        for key, default in (("soon_days", config.DEFAULT_SOON_DAYS),
                             ("upcoming_days", config.DEFAULT_UPCOMING_DAYS),
                             ("default_period", config.DEFAULT_PERIOD_MONTHS)):
            raw = form.get(key, "")
            try:
                value = int(raw) if raw else default
            except ValueError:
                value = default
            book.set_org_value(key, value)
        from .excel_store import save_book

        _, error = safe(save_book, book, action="settings_save", reason="Изменение настроек организации")
        flash("Настройки сохранены." if not error else error["message"],
              "ok" if not error else "error")
        return redirect(url_for("settings_page"))

    @app.post("/settings/brand")
    def settings_brand():
        """Цвет фирмы: влияет на шапку и активную закладку."""
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        form = request.form.to_dict()
        color = navigation.normalize_color(form.get("brand_color", ""), "#1F4E79")
        book.set_org_value("brand_color", color)
        entry.color = color
        cat.save()
        from .excel_store import save_book

        _, error = safe(save_book, book, action="settings_brand",
                        reason=f"Изменён цвет фирмы: {color}")
        flash("Цвет фирмы сохранён." if not error else error["message"],
              "ok" if not error else "error")
        return redirect(url_for("settings_page"))

    @app.post("/settings/nav")
    def settings_nav():
        """Порядок закладок: вверх/вниз и сброс."""
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        form = request.form.to_dict()
        current = book.org_value("nav_order")
        action = form.get("action", "")
        if action == "reset":
            updated = ""
            note = "Порядок закладок сброшен"
        else:
            key = form.get("key", "")
            direction = "up" if action == "up" else "down"
            updated = navigation.move(current, key, direction)
            item = navigation.ITEM_BY_KEY.get(key)
            note = f"Закладка «{item.label if item else key}» перемещена"
        book.set_org_value("nav_order", updated)
        from .excel_store import save_book

        _, error = safe(save_book, book, action="settings_nav", reason=note)
        flash(note if not error else error["message"], "ok" if not error else "error")
        return redirect(url_for("settings_page") + "#nav")

    # -- резервное копирование -------------------------------------------
    @app.route("/backup", methods=["GET", "POST"])
    def backup_page():
        """Настройки резервирования: локальные копии и отправка на почту."""
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)

        if request.method == "POST":
            form = request.form.to_dict()
            current = backups.load_settings(book)
            try:
                current.keep_local = int(form.get("keep_local") or current.keep_local)
                current.email_every_days = int(
                    form.get("email_every_days") or current.email_every_days
                )
            except ValueError:
                pass
            current.email_enabled = bool(form.get("email_enabled"))
            current.smtp_preset = form.get("smtp_preset", current.smtp_preset)
            current.smtp_host = (form.get("smtp_host") or "").strip()
            try:
                current.smtp_port = int(form.get("smtp_port") or current.resolved_port)
            except ValueError:
                current.smtp_port = current.resolved_port
            current.smtp_ssl = bool(form.get("smtp_ssl"))
            current.sender = (form.get("sender") or "").strip()
            current.recipient = (form.get("recipient") or "").strip()
            current.username = (form.get("username") or "").strip()

            action = form.get("action", "save")
            message, kind = "", "ok"
            try:
                backups.validate_settings(current)
            except ValidationError as exc:
                message, kind = exc.message, "error"
            else:
                password = form.get("mail_password") or ""
                backups.save_settings(book, current, password=password)
                if action == "test":
                    result = backups.send_email(book, current, reason="проверка настроек")
                    if result["sent"]:
                        message = f"Проверочное письмо отправлено на {current.recipient}."
                    else:
                        message = f"Отправка не удалась: {result['error']}"
                        kind = "error"
                elif action == "now":
                    result = backups.send_email(book, current, reason="по команде оператора")
                    message = ("Копия отправлена на " + current.recipient
                               if result["sent"] else f"Отправка не удалась: {result['error']}")
                    kind = "ok" if result["sent"] else "error"
                else:
                    message = "Настройки резервирования сохранены."
            flash(message, kind)
            return redirect(url_for("backup_page"))

        settings_data = backups.load_settings(book)
        folder = book.backups_dir()
        files = sorted(folder.glob("registry_*.xlsx"), reverse=True) if folder.exists() else []
        listing = []
        for path in files[:20]:
            stat = path.stat()
            listing.append({
                "name": path.name,
                "size": f"{stat.st_size / 1024:.0f} КБ",
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%d.%m.%Y %H:%M"),
            })
        return render_template(
            "backup.html", entry=entry, book=book,
            settings=settings_data,
            presets=backups.MAIL_PRESETS,
            has_password=bool(backups.read_password(book)),
            status_text=backups.email_status(settings_data),
            backups=listing,
            total=len(files),
            next_email_in=_days_until_next(settings_data),
        )

    @app.post("/backup/local")
    def backup_local():
        """Создать локальную копию по команде оператора."""
        cat, entry = require_org()
        book, error = safe(load_book, entry)
        if error:
            return render_template("org_error.html", error=error, entry=entry)
        path, error = safe(backups.make_local_copy, book, reason="по команде оператора")
        flash(f"Копия создана: {path.name}" if not error else error["message"],
              "ok" if not error else "error")
        return redirect(url_for("backup_page"))

    @app.get("/journal")
    def journal_view():
        cat, entry = require_org()
        from .journal import read_events

        # Оператору нужны последние действия: 50 достаточно, чтобы понять,
        # что происходило, и не листать историю за месяцы.
        events = read_events(entry.resolve(base).parent / "logs", limit=50)
        return render_template("journal.html", entry=entry, events=events)

    @app.get("/about")
    def about():
        return render_template("about.html")

    @app.errorhandler(404)
    def not_found(_):
        return render_template("error.html", message="Страница не найдена."), 404

    @app.errorhandler(403)
    def forbidden(exc):
        return render_template(
            "error.html",
            message="Доступ к приложению возможен только с этого компьютера.",
        ), 403

    @app.errorhandler(400)
    def bad_request(exc):
        return render_template("error.html",
                               message=getattr(exc, "description", "Некорректный запрос.")), 400

    return app
