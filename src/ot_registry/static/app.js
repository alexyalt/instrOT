/* Небольшой помощник для полей с автодополнением.
   Данные берутся из <datalist> или из data-атрибута suggestions. */
(function () {
  "use strict";

  function attach(input) {
    var source = input.dataset.suggestions;
    var values = [];
    if (source) {
      try { values = JSON.parse(source); } catch (e) { values = []; }
    } else if (input.list) {
      values = Array.prototype.map.call(input.list.options, function (o) { return o.value; });
    }
    if (!values.length) return;

    var host = document.createElement("div");
    host.className = "autocomplete-host";
    input.parentNode.insertBefore(host, input);
    host.appendChild(input);
    input.setAttribute("autocomplete", "off");

    var list = document.createElement("ul");
    list.className = "autocomplete-list";
    list.hidden = true;
    host.appendChild(list);

    var active = -1;

    function close() {
      list.hidden = true;
      list.innerHTML = "";
      active = -1;
    }

    function render(items) {
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

    function filter() {
      var query = input.value.trim().toLowerCase();
      if (!query) { render(values); return; }
      render(values.filter(function (v) {
        return String(v).toLowerCase().indexOf(query) !== -1;
      }));
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

  document.addEventListener("DOMContentLoaded", function () {
    document.querySelectorAll("[data-suggestions], input[list]").forEach(attach);

    // Защита от двойной отправки формы.
    document.querySelectorAll("form[data-once]").forEach(function (form) {
      form.addEventListener("submit", function () {
        var buttons = form.querySelectorAll("button[type=submit]");
        window.setTimeout(function () {
          buttons.forEach(function (b) { b.disabled = true; });
        }, 0);
      });
    });

    // Подсветка выбранных строк в таблице печати.
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
})();
