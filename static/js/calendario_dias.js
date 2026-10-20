(function () {
  var DIAS_SEMANA = [
    "Domingo",
    "Segunda-feira",
    "Terça-feira",
    "Quarta-feira",
    "Quinta-feira",
    "Sexta-feira",
    "Sábado"
  ];

  function pad(n) {
    return String(n).padStart(2, "0");
  }

  function toISO(date) {
    return date.getFullYear() + "-" + pad(date.getMonth() + 1) + "-" + pad(date.getDate());
  }

  function formatarData(iso) {
    var partes = String(iso || "").split("-");
    if (partes.length !== 3) return iso || "";
    var data = new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
    if (isNaN(data.getTime())) return iso;
    var rotulo = pad(data.getDate()) + "/" + pad(data.getMonth() + 1) + "/" + data.getFullYear();
    return rotulo + " — " + DIAS_SEMANA[data.getDay()];
  }

  function horaOk(valor) {
    var match = /^(\d{2}):(\d{2})$/.exec(valor || "");
    if (!match) return false;
    var hora = Number(match[1]);
    var minuto = Number(match[2]);
    return hora >= 0 && hora <= 23 && minuto >= 0 && minuto <= 59;
  }

  function diffHours(inicio, fim) {
    if (!horaOk(inicio) || !horaOk(fim)) return 0;
    var a = inicio.split(":").map(Number);
    var b = fim.split(":").map(Number);
    var start = a[0] * 60 + a[1];
    var end = b[0] * 60 + b[1];
    if (end < start) end += 24 * 60;
    if (end === start) return 0;
    return Math.round(((end - start) / 60) * 100) / 100;
  }

  function formatHours(value) {
    return String(value.toFixed(2)).replace(/\.00$/, "").replace(/(\.\d)0$/, "$1");
  }

  function lerHorario(input, fallback) {
    if (input && horaOk(input.value)) return input.value;
    if (input && horaOk(input.dataset.committed)) return input.dataset.committed;
    return fallback;
  }

  function charsHorario(valor) {
    var texto = valor || "";
    if (/^[0-9_]{2}:[0-9_]{2}$/.test(texto)) return texto.split("");
    var digitos = texto.replace(/\D/g, "").slice(0, 4);
    var slots = ["_", "_", "_", "_"];
    for (var i = 0; i < digitos.length; i++) slots[i] = digitos[i];
    return [slots[0], slots[1], ":", slots[2], slots[3]];
  }

  function aceitaDigito(chars, index, digito) {
    if (!/^\d$/.test(digito) || chars[index] === ":") return false;
    if (index === 0) return digito >= "0" && digito <= "2";
    if (index === 1) return chars[0] !== "2" || (digito >= "0" && digito <= "3");
    if (index === 3) return digito >= "0" && digito <= "5";
    return index === 4;
  }

  function indiceDoDigito(pos) {
    if (pos <= 0) return 0;
    if (pos === 1) return 1;
    if (pos === 2 || pos === 3) return 3;
    return 4;
  }

  function caretDepois(index) {
    if (index === 0) return 1;
    if (index === 1) return 3;
    if (index === 3) return 4;
    return 5;
  }

  function publicarHorario(input) {
    if (!horaOk(input.value)) return;
    input.dataset.committed = input.value;
    input.classList.remove("is-invalid");
    input.dispatchEvent(new Event("input", { bubbles: true }));
  }

  function finalizarHorario(input) {
    if (horaOk(input.value)) {
      input.dataset.committed = input.value;
      input.classList.remove("is-invalid");
      return input.value;
    }
    if (horaOk(input.dataset.committed)) {
      input.value = input.dataset.committed;
      input.classList.remove("is-invalid");
      return input.dataset.committed;
    }
    input.value = "";
    input.classList.add("is-invalid");
    return "";
  }

  function aplicarDigitos(input, texto) {
    var chars = ["_", "_", ":", "_", "_"];
    var digitos = String(texto || "").replace(/\D/g, "").slice(0, 4);
    var mapa = [0, 1, 3, 4];
    for (var i = 0; i < digitos.length; i++) {
      if (!aceitaDigito(chars, mapa[i], digitos[i])) break;
      chars[mapa[i]] = digitos[i];
    }
    input.value = chars.join("");
    publicarHorario(input);
  }

  function tratarTeclaHorario(ev) {
    var input = ev.target;
    if (!input || !input.hasAttribute || !input.hasAttribute("data-horario")) return;
    if (ev.ctrlKey || ev.metaKey || ev.altKey) return;
    var tecla = ev.key;
    if (tecla === "Tab") return;

    if (tecla === "ArrowLeft" || tecla === "ArrowRight") {
      ev.preventDefault();
      var atual = input.selectionStart || 0;
      var passo = tecla === "ArrowLeft" ? -1 : 1;
      var destino = atual + passo;
      if (destino === 2) destino += passo;
      destino = Math.max(0, Math.min(5, destino));
      input.setSelectionRange(destino, destino);
      return;
    }

    if (tecla === ":" || tecla === ";") {
      ev.preventDefault();
      input.setSelectionRange(3, 3);
      return;
    }

    if (tecla === "Backspace") {
      ev.preventDefault();
      var chars = charsHorario(input.value);
      var inicioSel = input.selectionStart || 0;
      var fimSel = input.selectionEnd || 0;
      if (inicioSel !== fimSel) {
        for (var i = inicioSel; i < fimSel && i < chars.length; i++) {
          if (chars[i] !== ":") chars[i] = "_";
        }
        input.value = chars.join("");
        input.setSelectionRange(inicioSel, inicioSel);
        return;
      }
      var idx = inicioSel - 1;
      if (idx === 2) idx = 1;
      if (idx < 0) return;
      chars[idx] = "_";
      input.value = chars.join("");
      input.setSelectionRange(idx, idx);
      return;
    }

    if (tecla === "Delete") {
      ev.preventDefault();
      var base = charsHorario(input.value);
      var pos = input.selectionStart || 0;
      if (pos === 2) pos = 3;
      if (pos > 4) return;
      if (base[pos] !== ":") base[pos] = "_";
      input.value = base.join("");
      input.setSelectionRange(pos, pos);
      return;
    }

    if (/^\d$/.test(tecla)) {
      ev.preventDefault();
      var atuais = charsHorario(input.value);
      var index = indiceDoDigito(input.selectionStart || 0);
      if (!aceitaDigito(atuais, index, tecla)) return;
      atuais[index] = tecla;
      if (index === 0 && tecla === "2" && atuais[1] !== "_" && atuais[1] > "3") {
        atuais[1] = "_";
      }
      input.value = atuais.join("");
      var caret = caretDepois(index);
      input.setSelectionRange(caret, caret);
      publicarHorario(input);
      return;
    }

    if (tecla.length === 1) ev.preventDefault();
  }

  function initCalendario(root) {
    if (!root) return;
    var hidden = document.getElementById(root.dataset.target);
    var listEl = document.getElementById(root.dataset.list);
    var totalEl = document.getElementById(root.dataset.total);
    var monthLabel = root.querySelector("[data-month-label]");
    var grid = root.querySelector("[data-grid]");
    var defaultInicio = root.querySelector("[data-default-inicio]");
    var defaultFim = root.querySelector("[data-default-fim]");
    var prevBtn = root.querySelector("[data-prev]");
    var nextBtn = root.querySelector("[data-next]");
    var applyBtn = root.querySelector("[data-apply-hours]");

    var state = [];
    try {
      state = JSON.parse(hidden.value || "[]");
    } catch (e) {
      state = [];
    }

    var cursor = new Date();
    cursor.setDate(1);

    function atualizarTotal() {
      if (!totalEl) return;
      var total = state.reduce(function (acc, item) {
        return acc + diffHours(item.hora_inicio, item.hora_fim);
      }, 0);
      totalEl.textContent = formatHours(total);
    }

    function renderLista() {
      if (!listEl) return;
      listEl.innerHTML = state
        .slice()
        .sort(function (a, b) { return a.data.localeCompare(b.data); })
        .map(function (item) {
          var horas = diffHours(item.hora_inicio, item.hora_fim);
          var inicio = horaOk(item.hora_inicio) ? item.hora_inicio : "";
          var fim = horaOk(item.hora_fim) ? item.hora_fim : "";
          return (
            '<div class="border-bottom py-2" data-dia="' + item.data + '">' +
            '<div class="d-flex justify-content-between align-items-center mb-1 gap-2">' +
            "<strong>" + formatarData(item.data) + "</strong>" +
            '<button type="button" class="btn btn-sm btn-outline-danger" data-remove="' + item.data + '">Remover</button>' +
            "</div>" +
            '<div class="d-flex flex-wrap gap-2 align-items-center">' +
            '<label class="small mb-0">Início</label>' +
            '<input type="text" class="form-control form-control-sm horario-input" data-horario data-inicio="' + item.data + '" value="' + inicio + '" data-committed="' + inicio + '" inputmode="numeric" maxlength="5" placeholder="HH:MM" autocomplete="off" spellcheck="false" aria-label="Início">' +
            '<label class="small mb-0">Fim</label>' +
            '<input type="text" class="form-control form-control-sm horario-input" data-horario data-fim="' + item.data + '" value="' + fim + '" data-committed="' + fim + '" inputmode="numeric" maxlength="5" placeholder="HH:MM" autocomplete="off" spellcheck="false" aria-label="Fim">' +
            '<span class="small text-muted" data-horas="' + item.data + '">' + formatHours(horas) + "h</span>" +
            "</div></div>"
          );
        })
        .join("");
    }

    function sync() {
      hidden.value = JSON.stringify(state);
      renderLista();
      atualizarTotal();
      render();
    }

    function toggleDay(iso) {
      var idx = state.findIndex(function (item) { return item.data === iso; });
      if (idx >= 0) {
        state.splice(idx, 1);
      } else {
        state.push({
          data: iso,
          hora_inicio: lerHorario(defaultInicio, "08:00"),
          hora_fim: lerHorario(defaultFim, "17:00"),
          observacao: ""
        });
      }
      sync();
    }

    function render() {
      var year = cursor.getFullYear();
      var month = cursor.getMonth();
      monthLabel.textContent = cursor.toLocaleDateString("pt-BR", { month: "long", year: "numeric" });
      grid.innerHTML = "";
      ["D", "S", "T", "Q", "Q", "S", "S"].forEach(function (label) {
        var h = document.createElement("div");
        h.className = "text-center small text-muted";
        h.textContent = label;
        grid.appendChild(h);
      });
      var firstDow = new Date(year, month, 1).getDay();
      var daysInMonth = new Date(year, month + 1, 0).getDate();
      for (var i = 0; i < firstDow; i++) {
        var empty = document.createElement("div");
        empty.className = "calendar-day muted";
        grid.appendChild(empty);
      }
      for (var day = 1; day <= daysInMonth; day++) {
        var iso = toISO(new Date(year, month, day));
        var cell = document.createElement("button");
        cell.type = "button";
        cell.className = "calendar-day";
        if (state.some(function (item) { return item.data === iso; })) {
          cell.classList.add("selected");
        }
        cell.textContent = String(day);
        cell.addEventListener("click", function (isoRef) {
          return function () { toggleDay(isoRef); };
        }(iso));
        grid.appendChild(cell);
      }
    }

    root.addEventListener("keydown", tratarTeclaHorario);
    root.addEventListener("paste", function (ev) {
      var input = ev.target;
      if (!input || !input.hasAttribute || !input.hasAttribute("data-horario")) return;
      ev.preventDefault();
      var texto = (ev.clipboardData || window.clipboardData).getData("text");
      aplicarDigitos(input, texto);
    });
    root.addEventListener("focusin", function (ev) {
      var input = ev.target;
      if (!input || !input.hasAttribute || !input.hasAttribute("data-horario")) return;
      if (horaOk(input.value)) return;
      if (!input.value) {
        input.value = "__:__";
        input.setSelectionRange(0, 0);
      }
    });
    root.addEventListener("focusout", function (ev) {
      var input = ev.target;
      if (!input || !input.hasAttribute || !input.hasAttribute("data-horario")) return;
      finalizarHorario(input);
    });

    if (listEl) {
      listEl.addEventListener("input", function (ev) {
        var target = ev.target;
        if (!target || !horaOk(target.value)) return;
        var item = null;
        if (target.dataset.inicio) {
          item = state.find(function (d) { return d.data === target.dataset.inicio; });
          if (item) item.hora_inicio = target.value;
        }
        if (target.dataset.fim) {
          item = state.find(function (d) { return d.data === target.dataset.fim; });
          if (item) item.hora_fim = target.value;
        }
        if (!item) return;
        target.dataset.committed = target.value;
        hidden.value = JSON.stringify(state);
        var span = listEl.querySelector('[data-horas="' + item.data + '"]');
        if (span) span.textContent = formatHours(diffHours(item.hora_inicio, item.hora_fim)) + "h";
        atualizarTotal();
      });
      listEl.addEventListener("click", function (ev) {
        var btn = ev.target.closest("[data-remove]");
        if (!btn) return;
        state = state.filter(function (item) { return item.data !== btn.dataset.remove; });
        sync();
      });
    }

    if (applyBtn) {
      applyBtn.addEventListener("click", function () {
        var inicio = lerHorario(defaultInicio, "");
        var fim = lerHorario(defaultFim, "");
        if (defaultInicio) defaultInicio.classList.toggle("is-invalid", !inicio);
        if (defaultFim) defaultFim.classList.toggle("is-invalid", !fim);
        if (!inicio || !fim) return;
        state = state.map(function (item) {
          return {
            data: item.data,
            hora_inicio: inicio,
            hora_fim: fim,
            observacao: item.observacao || ""
          };
        });
        sync();
      });
    }

    prevBtn.addEventListener("click", function () {
      cursor.setMonth(cursor.getMonth() - 1);
      render();
    });
    nextBtn.addEventListener("click", function () {
      cursor.setMonth(cursor.getMonth() + 1);
      render();
    });

    sync();
  }

  document.addEventListener("DOMContentLoaded", function () {
    document.querySelectorAll("[data-calendario]").forEach(initCalendario);
  });
})();
