/* Помощник интерфейса: автодополнение, сортировка таблиц, защита от двойной отправки. */
(function () {
  "use strict";

  /* ------------------------------------------------------------------ */
  /* Автодополнение полей ввода                                          */
  /* ------------------------------------------------------------------ */

  function attachAutocomplete(input) {
    var source = input.dataset.suggestions;
    var values = [];
    if (source) {
      try { values = JSON.parse(source); } catch (e) { values = []; }
    }
    if (!values.length) return;

    var host = document.createElement("div");
    host.className = "autocomplete-host";
    input.parentNode.insertBefore(host, input);
    host.appendChild(input);
    // Родные подсказки браузера отключены: они конкурируют с нашим списком.
    input.setAttribute("autocomplete", "off");
    input.setAttribute("autocorrect", "off");
    input.setAttribute("autocapitalize", "off");
    input.setAttribute("spellcheck", "false");

    var list = document.createElement("ul");
    list.className = "autocomplete-list";
    list.hidden = true;
    host.appendChild(list);

    var active = -1;
    var current = [];

    function close() {
      list.hidden = true;
      list.innerHTML = "";
      active = -1;
      current = [];
    }

    function render(items) {
      current = items;
      if (!items.length) { close(); return; }
      list.innerHTML = "";
      items.slice(0, 12).forEach(function (value, index) {
        var li = document.createElement("li");
        li.textContent = value;
        li.addEventListener("mousedown", function (event) {
          event.preventDefault();
          input.value = value;
          close();
          input.dispatchEvent(new Event("change", { bubbles: true }));
        });
        if (index === active) li.classList.add("active");
        list.appendChild(li);
      });
      list.hidden = false;
    }

    // Поиск по любому слову, а не только с начала строки:
    // «охран» находит «Инструкция по охране труда».
    function filter() {
      var query = input.value.trim().toLowerCase();
      if (!query) { render(values); return; }
      var matches = values.filter(function (value) {
        return String(value).toLowerCase().indexOf(query) !== -1;
      });
      render(matches);
    }

    input.addEventListener("focus", filter);
    input.addEventListener("input", function () { active = -1; filter(); });
    input.addEventListener("keydown", function (event) {
      var items = list.querySelectorAll("li");
      if (event.key === "Escape") { close(); return; }
      if (!items.length) return;
      if (event.key === "ArrowDown" || event.key === "ArrowUp") {
        event.preventDefault();
        active += event.key === "ArrowDown" ? 1 : -1;
        if (active < 0) active = items.length - 1;
        if (active >= items.length) active = 0;
        items.forEach(function (li, i) { li.classList.toggle("active", i === active); });
      } else if (event.key === "Enter" && active >= 0) {
        event.preventDefault();
        input.value = items[active].textContent;
        close();
      }
    });
    input.addEventListener("blur", function () { setTimeout(close, 120); });
  }

  /* ------------------------------------------------------------------ */
  /* Сортировка таблиц по клику на заголовок                             */
  /* ------------------------------------------------------------------ */

  function cellValue(row, index) {
    var cell = row.cells[index];
    if (!cell) return "";
    // Данные для сортировки можно задать явно через data-sort.
    if (cell.dataset && cell.dataset.sort !== undefined) return cell.dataset.sort;
    return cell.textContent.trim();
  }

  function isDate(value) {
    return /^\d{2}\.\d{2}\.\d{4}$/.test(value);
  }

  function parseDate(value) {
    var parts = value.split(".");
    return new Date(+parts[2], +parts[1] - 1, +parts[0]).getTime();
  }

  function compare(a, b) {
    var emptyA = a === "" || a === "—" || a === "-";
    var emptyB = b === "" || b === "—" || b === "-";
    // Пустые значения всегда в конце, независимо от направления.
    if (emptyA && emptyB) return 0;
    if (emptyA) return 1;
    if (emptyB) return -1;

    if (isDate(a) && isDate(b)) return parseDate(a) - parseDate(b);

    var numA = Number(a.replace(/\s/g, "").replace(",", "."));
    var numB = Number(b.replace(/\s/g, "").replace(",", "."));
    if (!isNaN(numA) && !isNaN(numB) && a !== "" && b !== "") return numA - numB;

    return a.localeCompare(b, "ru", { numeric: true, sensitivity: "base" });
  }

  function makeSortable(table) {
    if (table.dataset.sortable === "off") return;
    var headerRow = table.tHead ? table.tHead.rows[0] : null;
    if (!headerRow) return;

    // Тело таблицы: tbody, а если его нет — все строки кроме заголовка.
    var tbody = table.tBodies[0] || table;
    var rows = Array.prototype.slice.call(tbody.rows);
    // Строки-разделители (с colspan) не сортируются.
    var sortableRows = rows.filter(function (r) { return r.cells.length > 1; });
    if (sortableRows.length < 2) return;

    Array.prototype.forEach.call(headerRow.cells, function (th, index) {
      if (th.dataset.sortable === "off" || th.colSpan > 1) return;
      th.classList.add("sortable");
      th.setAttribute("role", "button");
      th.setAttribute("tabindex", "0");
      th.title = "Нажмите для сортировки";

      var state = 0; // 0 — нет, 1 — по возрастанию, 2 — по убыванию

      function apply() {
        state = state === 2 ? 1 : state + 1;
        Array.prototype.forEach.call(headerRow.cells, function (other) {
          other.classList.remove("sort-asc", "sort-desc");
          if (other !== th) {
            other.removeAttribute("aria-sort");
            if (other._sortState) other._sortState.value = 0;
          }
        });
        th.classList.add(state === 1 ? "sort-asc" : "sort-desc");
        th.setAttribute("aria-sort", state === 1 ? "ascending" : "descending");

        var sorted = rows.slice().sort(function (r1, r2) {
          var result = compare(cellValue(r1, index), cellValue(r2, index));
          if (state === 2) result = -result;
          // Пустые значения оставляем в конце и при обратном порядке.
          var e1 = cellValue(r1, index) === "";
          var e2 = cellValue(r2, index) === "";
          if (e1 !== e2) return e1 ? 1 : -1;
          return result;
        });
        sorted.forEach(function (row) { tbody.appendChild(row); });
      }

      th._sortState = { get value() { return state; }, set value(v) { state = v; } };
      th.addEventListener("click", apply);
      th.addEventListener("keydown", function (event) {
        if (event.key === "Enter" || event.key === " ") { event.preventDefault(); apply(); }
      });
    });
  }

  /* ------------------------------------------------------------------ */
  /* Инициализация                                                       */
  /* ------------------------------------------------------------------ */

  document.addEventListener("DOMContentLoaded", function () {
    // Отключаем подсказки браузера во всех полях: мешают и мешают
    // собственному автодополнению.
    document.querySelectorAll("input:not([type=checkbox]):not([type=radio])").forEach(function (el) {
      if (!el.hasAttribute("autocomplete")) el.setAttribute("autocomplete", "off");
    });

    document.querySelectorAll("[data-suggestions]").forEach(attachAutocomplete);

    document.querySelectorAll("table.data").forEach(makeSortable);

    // Защита от двойной отправки формы.
    document.querySelectorAll("form[data-once]").forEach(function (form) {
      form.addEventListener("submit", function () {
        var buttons = form.querySelectorAll("button[type=submit]");
        window.setTimeout(function () {
          buttons.forEach(function (b) { b.disabled = true; });
        }, 0);
      });
    });

    // Отметить все строки в таблице печати.
    document.querySelectorAll("[data-select-all]").forEach(function (master) {
      var target = document.querySelector(master.dataset.selectAll);
      if (!target) return;
      master.addEventListener("change", function () {
        target.querySelectorAll("input[type=checkbox]").forEach(function (box) {
          box.checked = master.checked;
        });
      });
    });
  });
}());
