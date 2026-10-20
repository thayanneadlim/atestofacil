(function () {
  function normalizar(texto) {
    return (texto || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function initBusca(select) {
    if (select.dataset.buscaPronta) return;
    select.dataset.buscaPronta = "1";
    var form = select.closest("form");
    var url = form ? form.dataset.buscaUrl : "";
    var tipo = select.dataset.buscaCadastro;
    var wrapper = document.createElement("div");
    wrapper.className = "busca-cadastro";
    var input = document.createElement("input");
    input.type = "text";
    input.className = "form-control";
    input.autocomplete = "off";
    input.spellcheck = false;
    input.setAttribute("data-busca-input", "");
    input.setAttribute("aria-autocomplete", "list");
    input.placeholder = tipo === "empresa" ? "Digite o nome da empresa" : "Digite o nome da pessoa";
    var lista = document.createElement("div");
    lista.className = "busca-cadastro-lista d-none";
    lista.setAttribute("role", "listbox");
    var erro = document.createElement("div");
    erro.className = "text-danger small mt-1 d-none";
    erro.setAttribute("data-busca-erro", "");

    select.classList.add("d-none");
    select.removeAttribute("required");
    select.parentNode.insertBefore(wrapper, select);
    wrapper.appendChild(input);
    wrapper.appendChild(lista);
    wrapper.appendChild(erro);
    wrapper.appendChild(select);

    var selecionado = select.options[select.selectedIndex];
    if (selecionado && selecionado.value) input.value = selecionado.textContent.trim();
    if (select.disabled) {
      input.readOnly = true;
      input.disabled = true;
      return;
    }

    var resultados = [];
    var destaque = -1;
    var sequencia = 0;
    var timer = null;

    function fechar() {
      lista.classList.add("d-none");
      lista.innerHTML = "";
      destaque = -1;
    }

    function nomeSelecionado() {
      var opcao = select.options[select.selectedIndex];
      return opcao && opcao.value ? opcao.textContent.trim() : "";
    }

    function definirRegistro(item) {
      var valor = String(item.id);
      var achou = false;
      for (var i = 0; i < select.options.length; i++) {
        if (select.options[i].value === valor) {
          select.selectedIndex = i;
          achou = true;
          break;
        }
      }
      if (!achou) {
        var opcao = new Option(item.nome, valor, true, true);
        select.add(opcao);
      }
      select.value = valor;
      input.value = item.nome;
      erro.classList.add("d-none");
      fechar();
    }

    function desenhar() {
      lista.innerHTML = "";
      if (!resultados.length) {
        var vazio = document.createElement("div");
        vazio.className = "busca-cadastro-vazio";
        vazio.textContent = "Nenhum resultado encontrado.";
        lista.appendChild(vazio);
        lista.classList.remove("d-none");
        return;
      }
      resultados.forEach(function (item, indice) {
        var botao = document.createElement("button");
        botao.type = "button";
        botao.className = "busca-cadastro-item" + (indice === destaque ? " is-active" : "");
        botao.textContent = item.nome;
        botao.setAttribute("role", "option");
        botao.addEventListener("mousedown", function (ev) { ev.preventDefault(); });
        botao.addEventListener("click", function () { definirRegistro(item); });
        lista.appendChild(botao);
      });
      lista.classList.remove("d-none");
      var ativo = lista.querySelector(".is-active");
      if (ativo && ativo.scrollIntoView) ativo.scrollIntoView({ block: "nearest" });
    }

    function exatoDigitado() {
      var digitado = normalizar(input.value);
      if (!digitado) return null;
      var exatos = resultados.filter(function (item) {
        return normalizar(item.nome) === digitado;
      });
      return exatos.length === 1 ? exatos[0] : null;
    }

    function confirmar() {
      var exato = exatoDigitado();
      if (exato) {
        definirRegistro(exato);
        return true;
      }
      if (destaque >= 0 && resultados[destaque]) {
        definirRegistro(resultados[destaque]);
        return true;
      }
      return false;
    }

    function buscar() {
      var termo = input.value.trim();
      if (normalizar(termo) !== normalizar(nomeSelecionado())) select.value = "";
      if (!url || termo.length < 1) {
        resultados = [];
        fechar();
        return;
      }
      var pedido = ++sequencia;
      var destino = url + "?tipo=" + encodeURIComponent(tipo) + "&q=" + encodeURIComponent(termo);
      fetch(destino, { headers: { Accept: "application/json" }, credentials: "same-origin" })
        .then(function (resp) { return resp.ok ? resp.json() : { results: [] }; })
        .then(function (dados) {
          if (pedido !== sequencia) return;
          resultados = (dados && dados.results) || [];
          destaque = resultados.length ? 0 : -1;
          desenhar();
        })
        .catch(function () {
          if (pedido !== sequencia) return;
          resultados = [];
          destaque = -1;
          desenhar();
        });
    }

    input.addEventListener("input", function () {
      erro.classList.add("d-none");
      if (normalizar(input.value) !== normalizar(nomeSelecionado())) select.value = "";
      window.clearTimeout(timer);
      timer = window.setTimeout(buscar, 200);
    });
    input.addEventListener("keydown", function (ev) {
      if (lista.classList.contains("d-none") && ev.key !== "Enter") return;
      if (ev.key === "ArrowDown") {
        ev.preventDefault();
        if (!resultados.length) return;
        destaque = Math.min(resultados.length - 1, destaque + 1);
        desenhar();
      } else if (ev.key === "ArrowUp") {
        ev.preventDefault();
        if (!resultados.length) return;
        destaque = Math.max(0, destaque - 1);
        desenhar();
      } else if (ev.key === "Enter") {
        var jaEscolhido = select.value && normalizar(input.value) === normalizar(nomeSelecionado());
        if (jaEscolhido || lista.classList.contains("d-none")) return;
        if (confirmar()) ev.preventDefault();
      } else if (ev.key === "Escape") {
        fechar();
      }
    });
    input.addEventListener("blur", function () {
      window.setTimeout(function () {
        var exato = exatoDigitado();
        if (exato) definirRegistro(exato);
        fechar();
      }, 150);
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    var selects = document.querySelectorAll("[data-busca-cadastro]");
    if (!selects.length) return;
    selects.forEach(initBusca);
    var form = selects[0].closest("form");
    if (!form) return;
    form.addEventListener("submit", function (ev) {
      var invalido = false;
      selects.forEach(function (select) {
        if (select.disabled) return;
        var input = select.parentElement.querySelector("[data-busca-input]");
        var erro = select.parentElement.querySelector("[data-busca-erro]");
        var opcao = select.options[select.selectedIndex];
        var ok = !!(select.value && opcao && input && opcao.textContent.trim() === input.value.trim());
        if (!ok) {
          invalido = true;
          erro.classList.remove("d-none");
          erro.textContent = select.dataset.buscaCadastro === "empresa"
            ? "Selecione uma empresa da lista."
            : "Selecione uma pessoa da lista.";
        } else if (erro) {
          erro.classList.add("d-none");
        }
      });
      if (invalido) {
        ev.preventDefault();
        var visivel = form.querySelector("[data-busca-erro]:not(.d-none)");
        var campo = visivel && visivel.parentElement.querySelector("[data-busca-input]");
        if (campo) campo.focus();
      }
    });
  });
})();
