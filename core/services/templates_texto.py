import re
from datetime import date, timedelta
from decimal import Decimal


VARIAVEIS_SUPORTADAS = {
    "nome",
    "nome_curto",
    "cpf",
    "chave_pix",
    "numero_documento",
    "n_doc_sem_ano",
    "data_documento",
    "dia_ano_doc",
    "data_vencimento",
    "empresa",
    "dias_trabalhados",
    "horarios_trabalhados",
    "observacoes",
    "motivo_contratacao",
    "funcao",
    "total_horas",
    "valor_total_formatado",
}

_BRACKET_ALIAS = {
    "NOME": "nome",
    "NOME_CURTO": "nome_curto",
    "CPF": "cpf",
    "CHAVE_PIX": "chave_pix",
    "NUMERO_DOCUMENTO": "numero_documento",
    "N_DOC_SEM_ANO": "n_doc_sem_ano",
    "DATA_DOCUMENTO": "data_documento",
    "DIA_ANO_DOC": "dia_ano_doc",
    "DATA_VENCIMENTO": "data_vencimento",
    "EMPRESA": "empresa",
    "DIAS_TRABALHADOS": "dias_trabalhados",
    "HORARIOS_TRABALHADOS": "horarios_trabalhados",
    "OBSERVACOES": "observacoes",
    "MOTIVO_CONTRATACAO": "motivo_contratacao",
    "FUNCAO": "funcao",
    "TOTAL_HORAS": "total_horas",
    "VALOR_TOTAL_FORMATADO": "valor_total_formatado",
}

_MESES = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}


def formatar_cpf(cpf_digitos: str) -> str:
    digitos = re.sub(r"\D", "", cpf_digitos or "")
    if len(digitos) != 11:
        return digitos
    return f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"


def validar_cpf(cpf: str) -> bool:
    digitos = re.sub(r"\D", "", cpf or "")
    if len(digitos) != 11 or digitos == digitos[0] * 11:
        return False

    def _calc(base: str, weights: list[int]) -> int:
        total = sum(int(d) * w for d, w in zip(base, weights))
        resto = total % 11
        return 0 if resto < 2 else 11 - resto

    d1 = _calc(digitos[:9], list(range(10, 1, -1)))
    d2 = _calc(digitos[:10], list(range(11, 1, -1)))
    return digitos[-2:] == f"{d1}{d2}"


def normalizar_cpf(cpf: str) -> str:
    return re.sub(r"\D", "", cpf or "")


def extrair_nome_curto(nome: str) -> str:
    """Retorna primeiro e último nome (ex.: João da Silva -> João Silva)."""
    partes = [p for p in (nome or "").strip().split() if p]
    if not partes:
        return ""
    if len(partes) == 1:
        return partes[0]
    return f"{partes[0]} {partes[-1]}"


def extrair_n_doc_sem_ano(numero_documento: str) -> str:
    """Retorna o número do documento sem o ano (ex.: 001/2026 -> 001)."""
    texto = (numero_documento or "").strip()
    if "/" in texto:
        return texto.split("/", 1)[0].strip()
    return texto


def formatar_dia_ano_doc(data_documento) -> str:
    """Dia e ano de data_documento no formato DD/AAAA."""
    if isinstance(data_documento, date):
        return f"{data_documento.day:02d}/{data_documento.year}"
    texto = (data_documento or "").strip()
    # Aceita DD/MM/AAAA
    partes = texto.split("/")
    if len(partes) == 3 and partes[0] and partes[2]:
        return f"{partes[0].zfill(2)}/{partes[2]}"
    return texto


def adicionar_dias_uteis(data_base: date, quantidade: int = 1) -> date:
    """
    Soma dias úteis à data (ignora sábado e domingo).
    Ex.: sexta + 1 útil = segunda.
    """
    if quantidade < 0:
        raise ValueError("Quantidade de dias úteis deve ser >= 0.")
    atual = data_base
    restantes = quantidade
    while restantes > 0:
        atual += timedelta(days=1)
        if atual.weekday() < 5:
            restantes -= 1
    return atual


def _fmt_hora(valor) -> str:
    if hasattr(valor, "strftime"):
        return valor.strftime("%H:%M")
    return str(valor)[:5]


def _fmt_data_extenso(valor: date) -> str:
    return f"{valor.day:02d} de {_MESES[valor.month]}"


def _fmt_data_curta(valor: date) -> str:
    return valor.strftime("%d/%m")


def _horario_tuple(dia) -> tuple[str, str]:
    return (_fmt_hora(dia.hora_inicio), _fmt_hora(dia.hora_fim))


def _sao_consecutivos(datas: list[date]) -> bool:
    """True quando as datas selecionadas formam um período corrido, sem lacunas."""
    if len(datas) <= 1:
        return True
    ordenadas = sorted(datas)
    for atual, proximo in zip(ordenadas, ordenadas[1:]):
        if proximo != atual + timedelta(days=1):
            return False
    return True


def _juntar_partes(partes: list[str]) -> str:
    """Junta itens com 'e' (ex.: dias 05/09 e 07/09)."""
    if not partes:
        return ""
    if len(partes) == 1:
        return partes[0]
    if len(partes) == 2:
        return f"{partes[0]} e {partes[1]}"
    return f"{', '.join(partes[:-1])} e {partes[-1]}"


def _juntar_grupos(partes: list[str], separador: str = "\n") -> str:
    """Junta conjuntos de horário.

    Na CI o separador é uma quebra de linha. O gerador do Word converte \\n
    em quebra de linha real do documento.
    """
    return separador.join(partes)


def _periodo_extenso(primeira: date, ultima: date) -> str:
    if primeira == ultima:
        return _fmt_data_extenso(primeira)
    return (
        f"{primeira.day:02d} de {_MESES[primeira.month]} até "
        f"{ultima.day:02d} de {_MESES[ultima.month]}"
    )


def _agrupar_por_horario(dias) -> list[tuple[tuple[str, str], list]]:
    """Mantém a ordem de aparição do primeiro dia de cada horário."""
    grupos: dict[tuple[str, str], list] = {}
    ordem: list[tuple[str, str]] = []
    for dia in sorted(dias, key=lambda d: d.data):
        chave = _horario_tuple(dia)
        if chave not in grupos:
            grupos[chave] = []
            ordem.append(chave)
        grupos[chave].append(dia)
    return [(chave, grupos[chave]) for chave in ordem]


def _datas_curtas(dias) -> str:
    return _juntar_partes([_fmt_data_curta(d.data) for d in sorted(dias, key=lambda x: x.data)])


def formatar_dias_trabalhados(dias) -> str:
    """
    Gera texto corrido para documentos.

    1) Dias corridos + horário igual:
       05 de Setembro até 08 de Setembro das 08:00h às 17:00h.

    2) Dias corridos + horários diferentes (agrupa iguais, um conjunto por linha):
       05 de Setembro até 08 de Setembro 05/09 e 07/09: 08:00h às 17:00h
       06/09 e 08/09: 08:00h às 10:00h.

    3) Dias quebrados + horário igual:
       05/09 e 07/09 das 08:00h às 17:00h.

    4) Dias quebrados + horários divergentes (um conjunto por linha):
       05/09: 08:00h às 17:00h
       07/09: 09:00h às 18:00h.
    """
    lista = sorted(list(dias), key=lambda d: d.data)
    if not lista:
        return ""

    datas = [d.data for d in lista]
    corrido = _sao_consecutivos(datas)
    grupos = _agrupar_por_horario(lista)
    mesma_jornada = len(grupos) == 1
    primeira = lista[0].data
    ultima = lista[-1].data

    # Caso 1 — corridos e horário igual
    if corrido and mesma_jornada:
        inicio, fim = grupos[0][0]
        if len(lista) == 1:
            return f"{_fmt_data_extenso(primeira)} das {inicio}h às {fim}h."
        return f"{_periodo_extenso(primeira, ultima)} das {inicio}h às {fim}h."

    # Caso 2 — corridos e horários diferentes (agrupa dias com mesma jornada)
    if corrido and not mesma_jornada:
        partes = [
            f"{_datas_curtas(dias_grupo)}: {inicio}h às {fim}h"
            for (inicio, fim), dias_grupo in grupos
        ]
        return f"{_periodo_extenso(primeira, ultima)} {_juntar_grupos(partes)}."

    # Caso 3 — quebrados e horário igual
    if not corrido and mesma_jornada:
        inicio, fim = grupos[0][0]
        return f"{_datas_curtas(lista)} das {inicio}h às {fim}h."

    # Caso 4 — quebrados e horários divergentes
    partes = []
    for (inicio, fim), dias_grupo in grupos:
        if len(dias_grupo) == 1:
            partes.append(f"{_fmt_data_curta(dias_grupo[0].data)}: {inicio}h às {fim}h")
        else:
            partes.append(f"{_datas_curtas(dias_grupo)}: {inicio}h às {fim}h")
    return f"{_juntar_grupos(partes)}."


def formatar_datas_trabalhadas(dias) -> str:
    """Datas do serviço, sem horário. Período corrido em extenso; dias soltos em DD/MM."""
    lista = sorted(list(dias), key=lambda item: item.data)
    if not lista:
        return ""
    datas = [item.data for item in lista]
    if _sao_consecutivos(datas):
        if len(lista) == 1:
            return _fmt_data_extenso(lista[0].data)
        return _periodo_extenso(lista[0].data, lista[-1].data)
    return _datas_curtas(lista)


def formatar_horarios_trabalhados(dias) -> str:
    """Horários do serviço, agrupados quando os dias não têm a mesma jornada."""
    lista = list(dias)
    if not lista:
        return ""
    grupos = _agrupar_por_horario(lista)
    if len(grupos) == 1:
        inicio, fim = grupos[0][0]
        return f"{inicio}h às {fim}h"
    partes = [
        f"{_datas_curtas(dias_grupo)}: {inicio}h às {fim}h"
        for (inicio, fim), dias_grupo in grupos
    ]
    return _juntar_grupos(partes, ", ")


def formatar_observacoes(dias) -> str:
    """Observações informadas nos dias, na ordem da data."""
    partes = []
    for dia in sorted(dias, key=lambda item: item.data):
        texto = (getattr(dia, "observacao", "") or "").strip()
        if texto:
            partes.append(texto)
    return "; ".join(partes)


def renderizar_texto(conteudo: str, contexto: dict) -> str:
    """
    Substitui variáveis no padrão principal {{ variavel }} e no alias [ VAR ].

    Valores ausentes são substituídos por string vazia.
    """
    resultado = conteudo or ""

    def _sub_mustache(match: re.Match) -> str:
        chave = match.group(1).strip().lower()
        return str(contexto.get(chave, ""))

    resultado = re.sub(r"\{\{\s*([a-zA-Z0-9_]+)\s*\}\}", _sub_mustache, resultado)

    def _sub_bracket(match: re.Match) -> str:
        raw = match.group(1).strip().upper()
        chave = _BRACKET_ALIAS.get(raw, raw.lower())
        return str(contexto.get(chave, ""))

    resultado = re.sub(r"\[\s*([A-Za-z0-9_]+)\s*\]", _sub_bracket, resultado)
    return resultado
