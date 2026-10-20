(function () {
  function apenasDigitos(texto) {
    return (texto || "").replace(/\D/g, "");
  }

  function limite(input) {
    const max = input.maxLength;
    return max > 0 ? max : 11;
  }

  function aplicar(input) {
    input.addEventListener("paste", function (event) {
      const clipboard = event.clipboardData || window.clipboardData;
      if (!clipboard) {
        return;
      }
      const colado = clipboard.getData("text");
      if (!colado) {
        return;
      }
      event.preventDefault();
      const digitos = apenasDigitos(colado);
      const inicio = input.selectionStart == null ? input.value.length : input.selectionStart;
      const fim = input.selectionEnd == null ? input.value.length : input.selectionEnd;
      const combinado = apenasDigitos(
        input.value.slice(0, inicio) + digitos + input.value.slice(fim)
      ).slice(0, limite(input));
      input.value = combinado;
      const posicao = Math.min(inicio + digitos.length, combinado.length);
      input.setSelectionRange(posicao, posicao);
    });
  }

  document.querySelectorAll("#id_cpf").forEach(aplicar);
})();
