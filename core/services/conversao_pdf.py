"""Converte planilhas e documentos do Word em PDF e gera os PNG da impressão.

Ordem dos conversores:

1. LibreOffice (`soffice` no PATH), quando estiver instalado.
2. WeasyPrint, a partir de HTML, quando as bibliotecas do sistema existem.
3. Desenho em Python com PyMuPDF, se o WeasyPrint não puder carregar.
4. PDF só com o texto, no mesmo espírito de ``core.services.pdf``, para a
   impressão ainda devolver páginas.

A união dos PDFs fica no ``pypdf``. A rasterização não usa ``pdftoppm``.
"""

from __future__ import annotations

import contextlib
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime, timedelta
from decimal import Decimal
from io import BytesIO, StringIO
from pathlib import Path

import pymupdf
from lxml import etree
from openpyxl.utils import coordinate_to_tuple, get_column_letter, range_boundaries
from pypdf import PdfReader, PdfWriter

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


class ConversaoPdfError(Exception):
    pass


_MENSAGEM = "Não foi possível preparar os documentos para impressão."
_DPI = 140
_PAPEL = {
    1: (612.0, 792.0),
    9: (595.28, 841.89),
    8: (841.89, 1190.55),
}
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_REF = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")
_WEASY = {"html": None, "falhou": False}


def localizar_soffice() -> str | None:
    """Procura o LibreOffice apenas no PATH."""
    return shutil.which("soffice") or shutil.which("soffice.exe")


def converter_para_pdf(pares: list[tuple[Path, Path]]) -> None:
    """Gera um PDF para cada par (origem, destino), na mesma ordem."""
    if not pares:
        return
    if localizar_soffice():
        try:
            _converter_soffice(pares)
            _exigir_pdfs(pares)
            return
        except ConversaoPdfError:
            _limpar_parciais(pares)
    for origem, destino in pares:
        destino.write_bytes(_pdf_python(origem))
    _exigir_pdfs(pares)


def unir_pdfs(pdfs: list[Path], destino: Path) -> None:
    """Junta os PDFs na ordem recebida."""
    if not pdfs:
        raise ConversaoPdfError("Nenhuma página foi gerada para impressão.")
    escritor = PdfWriter()
    try:
        for caminho in pdfs:
            leitor = PdfReader(str(caminho))
            if leitor.is_encrypted:
                leitor.decrypt("")
            if not leitor.pages:
                raise ConversaoPdfError(f"O arquivo {caminho.name} não tem páginas.")
            for pagina in leitor.pages:
                escritor.add_page(pagina)
        buffer = BytesIO()
        escritor.write(buffer)
    except ConversaoPdfError:
        raise
    except Exception as exc:
        raise ConversaoPdfError("Não foi possível juntar os documentos para impressão.") from exc
    destino.write_bytes(buffer.getvalue())


def rasterizar_pdf(origem: Path, *, dpi: int = _DPI) -> list[bytes]:
    """Transforma cada página do PDF em PNG."""
    try:
        documento = pymupdf.open(origem)
    except Exception as exc:
        raise ConversaoPdfError(_MENSAGEM) from exc
    paginas = []
    try:
        for pagina in documento:
            pixmap = pagina.get_pixmap(dpi=dpi, alpha=False)
            paginas.append(pixmap.tobytes("png"))
    except Exception as exc:
        raise ConversaoPdfError(_MENSAGEM) from exc
    finally:
        documento.close()
    if not paginas:
        raise ConversaoPdfError("Nenhuma página foi gerada para impressão.")
    return paginas


def imagem_para_pdf(dados: bytes) -> bytes:
    """Coloca PNG ou JPG em uma página A4, preservando a proporção."""
    pdf = pymupdf.open()
    try:
        pagina = pdf.new_page(width=_PAPEL[9][0], height=_PAPEL[9][1])
        area = pymupdf.Rect(36, 36, pagina.rect.width - 36, pagina.rect.height - 36)
        pagina.insert_image(area, stream=dados, keep_proportion=True)
        return pdf.tobytes(deflate=True)
    except Exception as exc:
        raise ConversaoPdfError("Não foi possível incluir o atesto assinado.") from exc
    finally:
        pdf.close()


def _pdf_python(origem: Path) -> bytes:
    dados = origem.read_bytes()
    tipo = _tipo_conteudo(origem, dados)
    if tipo == "pdf":
        return dados
    if tipo in {"png", "jpeg"}:
        return imagem_para_pdf(dados)
    if tipo in {"xlsx", "docx"}:
        if _weasy_disponivel():
            try:
                return _pdf_weasy(_html_office(origem, tipo))
            except Exception:
                pass
        try:
            if tipo == "xlsx":
                return _pdf_planilha(origem)
            return _pdf_word(origem)
        except ConversaoPdfError:
            raise
        except Exception:
            return _pdf_texto(_linhas_office(origem, tipo))
    return _pdf_texto([origem.name])


def _converter_soffice(pares: list[tuple[Path, Path]]) -> None:
    soffice = localizar_soffice()
    if not soffice:
        raise ConversaoPdfError(_MENSAGEM)
    with tempfile.TemporaryDirectory(prefix="lo-impressao-") as pasta:
        perfil = Path(pasta) / "profile"
        perfil.mkdir()
        saida = pares[0][1].parent
        _executar(
            [
                soffice,
                "--headless",
                "--norestore",
                "--nofirststartwizard",
                "--nologo",
                f"-env:UserInstallation={perfil.resolve().as_uri()}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(saida),
                *[str(origem) for origem, _ in pares],
            ],
            timeout=180,
        )


def _executar(comando: list[str], *, timeout: int) -> None:
    flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
    try:
        resultado = subprocess.run(
            comando,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
            stdin=subprocess.DEVNULL,
            creationflags=flags,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise ConversaoPdfError(_MENSAGEM) from exc
    if resultado.returncode != 0:
        raise ConversaoPdfError(_MENSAGEM)


def _exigir_pdfs(pares: list[tuple[Path, Path]]) -> None:
    ausentes = [
        origem.name
        for origem, destino in pares
        if not destino.is_file() or destino.stat().st_size == 0
    ]
    if ausentes:
        raise ConversaoPdfError(
            "Não foi possível converter todos os documentos: " + ", ".join(ausentes)
        )


def _limpar_parciais(pares: list[tuple[Path, Path]]) -> None:
    for _, destino in pares:
        if destino.exists():
            destino.unlink()


def _tipo_conteudo(origem: Path, dados: bytes) -> str:
    if dados.startswith(b"%PDF"):
        return "pdf"
    if dados.startswith(b"\x89PNG"):
        return "png"
    if dados.startswith(b"\xff\xd8"):
        return "jpeg"
    if dados.startswith(b"PK"):
        with zipfile.ZipFile(BytesIO(dados)) as pacote:
            nomes = pacote.namelist()
        if any(nome.startswith("xl/") for nome in nomes):
            return "xlsx"
        if any(nome.startswith("word/") for nome in nomes):
            return "docx"
    sufixo = origem.suffix.lower()
    return {
        ".xlsx": "xlsx",
        ".xlsm": "xlsx",
        ".docx": "docx",
        ".pdf": "pdf",
        ".png": "png",
        ".jpg": "jpeg",
        ".jpeg": "jpeg",
    }.get(sufixo, "")


def _weasy_disponivel() -> bool:
    if _WEASY["falhou"]:
        return False
    if _WEASY["html"] is not None:
        return True
    try:
        with contextlib.redirect_stdout(StringIO()), contextlib.redirect_stderr(StringIO()):
            from weasyprint import HTML
    except Exception:
        _WEASY["falhou"] = True
        return False
    _WEASY["html"] = HTML
    return True


def _pdf_weasy(html: str) -> bytes:
    html_cls = _WEASY["html"]
    if html_cls is None:
        raise ConversaoPdfError(_MENSAGEM)
    buffer = BytesIO()
    html_cls(string=html).write_pdf(buffer)
    pdf = buffer.getvalue()
    if not pdf.startswith(b"%PDF"):
        raise ConversaoPdfError(_MENSAGEM)
    return pdf


def _html_office(origem: Path, tipo: str) -> str:
    if tipo == "xlsx":
        return _html_planilha(origem)
    return _html_word(origem)


def _html_word(origem: Path) -> str:
    documento = Document(str(origem))
    blocos = []
    for paragrafo in documento.paragraphs:
        texto = _escapar(paragrafo.text)
        if texto:
            blocos.append(f"<p>{texto}</p>")
    for tabela in documento.tables:
        linhas = []
        for linha in tabela.rows:
            vistos = []
            celulas = []
            for celula in linha.cells:
                if celula._tc in vistos:
                    continue
                vistos.append(celula._tc)
                celulas.append(f"<td>{_escapar(celula.text)}</td>")
            linhas.append("<tr>" + "".join(celulas) + "</tr>")
        blocos.append("<table>" + "".join(linhas) + "</table>")
    return _documento_html("".join(blocos))


def _html_planilha(origem: Path) -> str:
    from openpyxl import load_workbook

    livro = load_workbook(origem, data_only=False)
    avaliador = _Avaliador(_grades(livro))
    partes = []
    for folha in livro.worksheets:
        if folha.sheet_state != "visible":
            continue
        min_col, min_row, max_col, max_row = _area_impressao(folha)
        cobertas, mesclas = _mesclas(folha, min_col, min_row, max_col, max_row)
        linhas = []
        for linha in range(min_row, max_row + 1):
            celulas = []
            for coluna in range(min_col, max_col + 1):
                if (linha, coluna) in cobertas:
                    continue
                fim_linha, fim_coluna = mesclas.get((linha, coluna), (linha, coluna))
                texto = _escapar(_texto_celula(folha, avaliador, linha, coluna))
                atributo = ""
                if fim_coluna > coluna:
                    atributo += f' colspan="{fim_coluna - coluna + 1}"'
                if fim_linha > linha:
                    atributo += f' rowspan="{fim_linha - linha + 1}"'
                celulas.append(f"<td{atributo}>{texto}</td>")
            linhas.append("<tr>" + "".join(celulas) + "</tr>")
        partes.append("<table>" + "".join(linhas) + "</table>")
    return _documento_html("".join(partes), quebra=True)


def _documento_html(corpo: str, *, quebra: bool = False) -> str:
    quebra_css = "table { break-after: page; }" if quebra else ""
    return (
        "<!DOCTYPE html><html><head><meta charset='utf-8'><style>"
        "@page { size: A4; margin: 12mm; }"
        "body { font-family: Arial, sans-serif; font-size: 10pt; }"
        "table { border-collapse: collapse; width: 100%; }"
        "td { border: 0.4pt solid #222; padding: 2pt; vertical-align: bottom; }"
        f"{quebra_css}</style></head><body>{corpo}</body></html>"
    )


def _escapar(texto: str) -> str:
    return (
        (texto or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\n", "<br>")
    )


class _Avaliador:
    """Calcula as fórmulas simples usadas nos modelos de previsão e RPA."""

    def __init__(self, grades: dict[str, dict[str, object]]):
        self.grades = grades
        self.memo: dict[tuple[str, str], object] = {}
        self.pilha: set[tuple[str, str]] = set()

    def valor(self, aba: str, coord: str):
        chave = (aba, coord)
        if chave in self.memo:
            return self.memo[chave]
        if chave in self.pilha:
            return 0
        bruto = self.grades.get(aba, {}).get(coord)
        if not isinstance(bruto, str) or not bruto.startswith("="):
            self.memo[chave] = bruto
            return bruto
        self.pilha.add(chave)
        try:
            resultado = self._calcular(bruto[1:], aba)
        except Exception:
            resultado = None
        self.pilha.discard(chave)
        self.memo[chave] = resultado
        return resultado

    def _calcular(self, formula: str, aba: str):
        self.texto = formula.upper()
        self.i = 0
        self.aba = aba
        valor = self._comparacao()
        return valor

    def _comparacao(self):
        esquerda = self._adicao()
        for operador in ("<=", ">=", "<>", "=", "<", ">"):
            if self._come(operador):
                direita = self._adicao()
                return self._comparar(esquerda, direita, operador)
        return esquerda

    def _adicao(self):
        valor = self._produto()
        while True:
            if self._come("+"):
                valor = _numero(valor) + _numero(self._produto())
            elif self._come("-"):
                valor = _numero(valor) - _numero(self._produto())
            else:
                return valor

    def _produto(self):
        valor = self._unario()
        while True:
            if self._come("*"):
                valor = _numero(valor) * _numero(self._unario())
            elif self._come("/"):
                divisor = _numero(self._unario())
                valor = None if divisor == 0 else _numero(valor) / divisor
            else:
                return valor

    def _unario(self):
        if self._come("-"):
            return -_numero(self._posfixo())
        if self._come("+"):
            return _numero(self._posfixo())
        return self._posfixo()

    def _posfixo(self):
        valor = self._primario()
        while self._come("%"):
            valor = _numero(valor) / 100
        return valor

    def _primario(self):
        self._espacos()
        if self._come("("):
            valor = self._comparacao()
            self._come(")")
            return valor
        if self.i < len(self.texto) and self.texto[self.i] == '"':
            return self._texto_aspas()
        if self.i < len(self.texto) and (self.texto[self.i].isdigit() or self.texto[self.i] == "."):
            return self._numero_literal()
        nome = self._nome()
        if nome and self._come("("):
            return self._funcao(nome)
        if nome:
            referencia = _REF.match(nome)
            if referencia:
                coord = f"{referencia.group(1)}{referencia.group(2)}"
                if self._come(":"):
                    fim = self._nome()
                    casamento = _REF.match(fim or "")
                    if not casamento:
                        return None
                    return self._faixa(coord, f"{casamento.group(1)}{casamento.group(2)}")
                return self.valor(self.aba, coord)
        return None

    def _funcao(self, nome: str):
        argumentos = []
        if not self._come(")"):
            while True:
                argumentos.append(self._comparacao())
                if self._come(")"):
                    break
                if not self._come(","):
                    break
        nome = nome.upper()
        if nome == "SUM":
            return _somar(argumentos)
        if nome == "IF":
            condicao = argumentos[0] if argumentos else False
            se_sim = argumentos[1] if len(argumentos) > 1 else 0
            se_nao = argumentos[2] if len(argumentos) > 2 else False
            return se_sim if _verdadeiro(condicao) else se_nao
        if nome == "AND":
            return all(_verdadeiro(item) for item in argumentos)
        if nome == "OR":
            return any(_verdadeiro(item) for item in argumentos)
        return None

    def _faixa(self, inicio: str, fim: str):
        linha1, coluna1 = coordinate_to_tuple(inicio)
        linha2, coluna2 = coordinate_to_tuple(fim)
        valores = []
        for linha in range(min(linha1, linha2), max(linha1, linha2) + 1):
            for coluna in range(min(coluna1, coluna2), max(coluna1, coluna2) + 1):
                coord = f"{get_column_letter(coluna)}{linha}"
                valores.append(self.valor(self.aba, coord))
        return valores

    def _texto_aspas(self) -> str:
        self.i += 1
        partes = []
        while self.i < len(self.texto):
            caractere = self.texto[self.i]
            self.i += 1
            if caractere == '"':
                if self.i < len(self.texto) and self.texto[self.i] == '"':
                    partes.append('"')
                    self.i += 1
                    continue
                break
            partes.append(caractere)
        return "".join(partes)

    def _numero_literal(self) -> float:
        inicio = self.i
        while self.i < len(self.texto) and (self.texto[self.i].isdigit() or self.texto[self.i] == "."):
            self.i += 1
        return float(self.texto[inicio:self.i])

    def _nome(self) -> str:
        self._espacos()
        inicio = self.i
        while self.i < len(self.texto) and (self.texto[self.i].isalnum() or self.texto[self.i] in "$"):
            self.i += 1
        return self.texto[inicio:self.i]

    def _come(self, operador: str) -> bool:
        self._espacos()
        if self.texto.startswith(operador, self.i):
            self.i += len(operador)
            return True
        return False

    def _espacos(self) -> None:
        while self.i < len(self.texto) and self.texto[self.i].isspace():
            self.i += 1

    @staticmethod
    def _comparar(esquerda, direita, operador: str) -> bool:
        if operador == "=":
            return _iguais(esquerda, direita)
        if operador == "<>":
            return not _iguais(esquerda, direita)
        a = _numero(esquerda)
        b = _numero(direita)
        if operador == "<":
            return a < b
        if operador == ">":
            return a > b
        if operador == "<=":
            return a <= b
        return a >= b


def _grades(livro) -> dict[str, dict[str, object]]:
    grades = {}
    for folha in livro.worksheets:
        grade = {}
        for linha in folha.iter_rows():
            for celula in linha:
                if celula.value is not None:
                    grade[celula.coordinate] = celula.value
        grades[folha.title] = grade
    return grades


def _numero(valor) -> float:
    if isinstance(valor, bool) or valor is None or valor == "":
        return 0.0
    if isinstance(valor, list):
        return _numero(valor[0]) if valor else 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, Decimal):
        return float(valor)
    if isinstance(valor, datetime):
        base = datetime(1899, 12, 30)
        return float((valor - base).days)
    if isinstance(valor, str):
        texto = valor.strip().replace("R$", "").replace(" ", "")
        if texto.endswith("%"):
            try:
                return float(texto[:-1].replace(".", "").replace(",", ".")) / 100
            except ValueError:
                return 0.0
        try:
            return float(texto.replace(".", "").replace(",", ".")) if "," in texto else float(texto)
        except ValueError:
            return 0.0
    return 0.0


def _iguais(esquerda, direita) -> bool:
    if _vazio(esquerda) and _vazio(direita):
        return True
    if isinstance(esquerda, str) or isinstance(direita, str):
        return str(esquerda or "").casefold() == str(direita or "").casefold()
    return _numero(esquerda) == _numero(direita)


def _vazio(valor) -> bool:
    return valor is None or valor == ""


def _verdadeiro(valor) -> bool:
    if isinstance(valor, bool):
        return valor
    if isinstance(valor, str):
        return valor != ""
    if valor is None:
        return False
    if isinstance(valor, list):
        return any(_verdadeiro(item) for item in valor)
    return _numero(valor) != 0


def _somar(argumentos) -> float:
    total = 0.0
    for argumento in argumentos:
        itens = argumento if isinstance(argumento, list) else [argumento]
        for item in itens:
            if isinstance(item, bool) or item is None or isinstance(item, str):
                continue
            if isinstance(item, (int, float, Decimal)):
                total += float(item)
    return total


def _texto_celula(folha, avaliador: _Avaliador, linha: int, coluna: int) -> str:
    celula = folha.cell(linha, coluna)
    bruto = celula.value
    if isinstance(bruto, str) and bruto.startswith("="):
        valor = avaliador.valor(folha.title, celula.coordinate)
    else:
        valor = bruto
    return _formatar_valor(valor, celula.number_format or "General")


def _formatar_valor(valor, formato: str) -> str:
    if valor is None:
        return ""
    if isinstance(valor, str):
        return "" if valor.startswith("=") else valor
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, bool):
        return "VERDADEIRO" if valor else "FALSO"
    if isinstance(valor, (int, float, Decimal)):
        return _formatar_numero(float(valor), formato or "General")
    return str(valor)


def _formatar_numero(numero: float, formato: str) -> str:
    if any(token in formato for token in ("#,##0", "0.00", "R$")):
        negativo = numero < 0
        texto = _milhar_brasileiro(abs(numero))
        if "R$" in formato:
            texto = f"R$ {texto}"
        return f"-{texto}" if negativo else texto
    if "%" in formato:
        return _milhar_brasileiro(numero * 100).rstrip("0").rstrip(",") + "%"
    if "d" in formato.lower() and "m" in formato.lower() and 1 <= numero <= 80000:
        data = datetime(1899, 12, 30) + timedelta(days=int(numero))
        return data.strftime("%d/%m/%Y")
    if abs(numero - round(numero)) < 1e-8:
        return str(int(round(numero)))
    texto = f"{numero:.4f}".rstrip("0").rstrip(".")
    return texto.replace(".", ",")


def _milhar_brasileiro(numero: float) -> str:
    inteiro, fracao = f"{numero:.2f}".split(".")
    milhar = f"{int(inteiro):,}".replace(",", ".")
    return f"{milhar},{fracao}"


def _pdf_planilha(origem: Path) -> bytes:
    from openpyxl import load_workbook

    livro = load_workbook(origem, data_only=False)
    avaliador = _Avaliador(_grades(livro))
    pdf = pymupdf.open()
    try:
        desenhada = False
        for folha in livro.worksheets:
            if folha.sheet_state != "visible":
                continue
            _desenhar_folha(pdf, folha, avaliador)
            desenhada = True
        if not desenhada:
            raise ConversaoPdfError(_MENSAGEM)
        return pdf.tobytes(deflate=True, garbage=3)
    finally:
        pdf.close()
        livro.close()


def _area_impressao(folha) -> tuple[int, int, int, int]:
    area = str(folha.print_area or "").split()
    if area:
        intervalo = area[0].split("!")[-1].replace("$", "")
        try:
            return range_boundaries(intervalo)
        except ValueError:
            pass
    max_col = min(folha.max_column or 1, 16)
    max_row = min(folha.max_row or 1, 120)
    return 1, 1, max_col, max_row


def _mesclas(folha, min_col, min_row, max_col, max_row):
    cobertas = set()
    origens = {}
    for faixa in folha.merged_cells.ranges:
        if faixa.max_col < min_col or faixa.min_col > max_col:
            continue
        if faixa.max_row < min_row or faixa.min_row > max_row:
            continue
        origem = (faixa.min_row, faixa.min_col)
        origens[origem] = (faixa.max_row, faixa.max_col)
        for linha in range(faixa.min_row, faixa.max_row + 1):
            for coluna in range(faixa.min_col, faixa.max_col + 1):
                if (linha, coluna) != origem:
                    cobertas.add((linha, coluna))
    return cobertas, origens


def _largura_coluna(folha, coluna: int) -> float:
    letra = get_column_letter(coluna)
    dimensao = folha.column_dimensions.get(letra)
    if dimensao is not None and dimensao.hidden:
        return 0.0
    largura = None
    if dimensao is not None and dimensao.width:
        largura = float(dimensao.width)
    if not largura:
        largura = float(folha.sheet_format.defaultColWidth or 8.43)
    pixels = int(((256 * largura + int(128 / 7)) / 256) * 7)
    return pixels * 72 / 96


def _altura_linha(folha, linha: int) -> float:
    dimensao = folha.row_dimensions.get(linha)
    if dimensao is not None and dimensao.hidden:
        return 0.0
    if dimensao is not None and dimensao.height:
        return float(dimensao.height)
    return float(folha.sheet_format.defaultRowHeight or 15)


def _papel(folha) -> tuple[float, float]:
    try:
        codigo = int(folha.page_setup.paperSize or 9)
    except (TypeError, ValueError):
        codigo = 9
    largura, altura = _PAPEL.get(codigo, _PAPEL[9])
    if folha.page_setup.orientation == "landscape":
        return altura, largura
    return largura, altura


def _margens(folha) -> tuple[float, float, float, float]:
    margens = folha.page_margins
    return (
        float(margens.left or 0.5) * 72,
        float(margens.top or 0.5) * 72,
        float(margens.right or 0.5) * 72,
        float(margens.bottom or 0.5) * 72,
    )


def _ajusta_a_pagina(folha) -> bool:
    propriedades = getattr(folha.sheet_properties, "pageSetUpPr", None)
    return bool(propriedades and propriedades.fitToPage)


def _escala_folha(folha, largura: float, altura: float, util_w: float, util_h: float) -> float:
    if largura <= 0 or altura <= 0:
        return 1.0
    if _ajusta_a_pagina(folha):
        escala = 1.0
        ajustar_largura = int(folha.page_setup.fitToWidth or 0)
        ajustar_altura = int(folha.page_setup.fitToHeight or 0)
        if ajustar_largura >= 1:
            escala = min(escala, util_w / largura)
        if ajustar_altura >= 1:
            escala = min(escala, util_h / altura)
        elif altura * escala > util_h and altura * escala <= util_h * 1.12:
            escala = min(escala, util_h / altura)
        return max(escala, 0.15)
    return max(min(float(folha.page_setup.scale or 100) / 100, 4), 0.15)


def _grupos_de_linhas(alturas: list[float], util_h: float, escala: float) -> list[list[int]]:
    grupos: list[list[int]] = []
    atual: list[int] = []
    usado = 0.0
    for indice, altura in enumerate(alturas):
        ocupacao = altura * escala
        if atual and usado + ocupacao > util_h + 0.5:
            grupos.append(atual)
            atual = [indice]
            usado = ocupacao
        else:
            atual.append(indice)
            usado += ocupacao
    if atual:
        grupos.append(atual)
    return grupos or [[]]


def _desenhar_folha(pdf, folha, avaliador: _Avaliador) -> None:
    min_col, min_row, max_col, max_row = _area_impressao(folha)
    larguras = [_largura_coluna(folha, coluna) for coluna in range(min_col, max_col + 1)]
    alturas = [_altura_linha(folha, linha) for linha in range(min_row, max_row + 1)]
    pagina_w, pagina_h = _papel(folha)
    esquerda, topo, direita, baixo = _margens(folha)
    util_w = max(pagina_w - esquerda - direita, 20)
    util_h = max(pagina_h - topo - baixo, 20)
    escala = _escala_folha(folha, sum(larguras), sum(alturas), util_w, util_h)
    cobertas, mesclas = _mesclas(folha, min_col, min_row, max_col, max_row)
    prefixo_x = [0.0]
    for largura in larguras:
        prefixo_x.append(prefixo_x[-1] + largura)
    absolutos_x = [0.0]
    for coluna in range(1, max_col + 2):
        absolutos_x.append(absolutos_x[-1] + _largura_coluna(folha, coluna))
    absolutos_y = [0.0]
    for linha in range(1, max_row + 2):
        absolutos_y.append(absolutos_y[-1] + _altura_linha(folha, linha))
    imagens = _imagens_planilha(folha)
    for grupo in _grupos_de_linhas(alturas, util_h, escala):
        if not grupo:
            continue
        pagina = pdf.new_page(width=pagina_w, height=pagina_h)
        fontes: dict[tuple[str, bool], str] = {}
        alturas_grupo = [alturas[indice] for indice in grupo]
        prefixo_y = [0.0]
        for altura in alturas_grupo:
            prefixo_y.append(prefixo_y[-1] + altura)
        conteudo_w = prefixo_x[-1] * escala
        origem_x = esquerda
        if getattr(folha.print_options, "horizontalCentered", False):
            origem_x += max(0, (util_w - conteudo_w) / 2)
        origem_y = topo
        linha_inicial = min_row + grupo[0]

        def ponto(coluna: int, linha: int) -> tuple[float, float]:
            x = origem_x + prefixo_x[coluna - min_col] * escala
            y = origem_y + prefixo_y[linha - linha_inicial] * escala
            return x, y

        for deslocamento in grupo:
            linha = min_row + deslocamento
            for coluna in range(min_col, max_col + 1):
                if (linha, coluna) in cobertas:
                    continue
                fim_linha, fim_coluna = mesclas.get((linha, coluna), (linha, coluna))
                fim_linha = min(fim_linha, min_row + grupo[-1])
                fim_coluna = min(fim_coluna, max_col)
                x0, y0 = ponto(coluna, linha)
                x1, y1 = ponto(fim_coluna + 1, fim_linha + 1)
                retangulo = pymupdf.Rect(x0, y0, x1, y1)
                _pintar_celula(pagina, folha, linha, coluna, retangulo)
                _bordas(
                    pagina,
                    folha,
                    linha,
                    coluna,
                    fim_linha,
                    fim_coluna,
                    retangulo,
                )
                texto = _texto_celula(folha, avaliador, linha, coluna).strip()
                if texto:
                    _escrever(
                        pagina,
                        fontes,
                        retangulo,
                        texto,
                        folha.cell(linha, coluna),
                        escala,
                    )
        for imagem in imagens:
            _desenhar_imagem_planilha(
                pagina,
                imagem,
                absolutos_x,
                absolutos_y,
                min_col,
                linha_inicial,
                min_row + grupo[-1],
                origem_x,
                origem_y,
                escala,
            )


def _pintar_celula(pagina, folha, linha, coluna, retangulo) -> None:
    cor = _cor_preenchimento(folha.cell(linha, coluna))
    if cor is None:
        return
    pagina.draw_rect(retangulo, color=cor, fill=cor, width=0)


def _cor_preenchimento(celula):
    preenchimento = celula.fill
    if not preenchimento or preenchimento.patternType != "solid":
        return None
    cor = preenchimento.fgColor
    if cor is None or getattr(cor, "type", None) != "rgb" or not cor.rgb:
        return None
    rgb = str(cor.rgb).upper()
    if rgb in {"FFFFFFFF", "00000000", "FFFFFF"}:
        return None
    if len(rgb) == 8:
        rgb = rgb[2:]
    if rgb == "FFFFFF":
        return None
    try:
        return tuple(int(rgb[indice : indice + 2], 16) / 255 for indice in (0, 2, 4))
    except ValueError:
        return None


def _bordas(pagina, folha, linha, coluna, fim_linha, fim_coluna, retangulo) -> None:
    lados = {
        "top": (folha.cell(linha, coluna).border.top, (retangulo.x0, retangulo.y0), (retangulo.x1, retangulo.y0)),
        "bottom": (
            folha.cell(fim_linha, coluna).border.bottom,
            (retangulo.x0, retangulo.y1),
            (retangulo.x1, retangulo.y1),
        ),
        "left": (folha.cell(linha, coluna).border.left, (retangulo.x0, retangulo.y0), (retangulo.x0, retangulo.y1)),
        "right": (
            folha.cell(linha, fim_coluna).border.right,
            (retangulo.x1, retangulo.y0),
            (retangulo.x1, retangulo.y1),
        ),
    }
    for estilo, inicio, fim in lados.values():
        largura = _espessura_borda(getattr(estilo, "style", None))
        if not largura:
            continue
        pagina.draw_line(inicio, fim, color=(0, 0, 0), width=largura)


def _espessura_borda(estilo) -> float:
    if not estilo:
        return 0
    if estilo in {"medium", "mediumDashed", "mediumDashDot"}:
        return 1.15
    if estilo in {"thick", "double"}:
        return 1.35
    return 0.55


def _escrever(pagina, fontes, retangulo, texto: str, celula, escala: float) -> None:
    fonte = celula.font
    nome = fonte.name or "Arial"
    negrito = bool(fonte.bold)
    altura = max(retangulo.height - 1.0, 4)
    tamanho = min(float(fonte.size or 10) * escala, altura / 1.15)
    tamanho = max(5.0, tamanho)
    fontname = _registrar_fonte(pagina, fontes, nome, negrito)
    alinhamento = _alinhamento(celula, texto)
    caixa = pymupdf.Rect(retangulo.x0 + 1, retangulo.y0 + 0.3, retangulo.x1 - 1, retangulo.y1 - 0.2)
    cor = _cor_fonte(celula)
    for _ in range(5):
        sobra = pagina.insert_textbox(
            caixa,
            texto,
            fontsize=tamanho,
            fontname=fontname,
            align=alinhamento,
            color=cor,
            lineheight=1.0,
        )
        if sobra >= -0.15:
            break
        tamanho = max(4.0, tamanho * caixa.height / (caixa.height - sobra))


def _alinhamento(celula, texto: str) -> int:
    horizontal = celula.alignment.horizontal
    if horizontal in {"center", "centerContinuous"}:
        return pymupdf.TEXT_ALIGN_CENTER
    if horizontal == "right":
        return pymupdf.TEXT_ALIGN_RIGHT
    if horizontal == "justify":
        return pymupdf.TEXT_ALIGN_JUSTIFY
    if horizontal in {None, "general"} and _parece_numero(texto):
        return pymupdf.TEXT_ALIGN_RIGHT
    return pymupdf.TEXT_ALIGN_LEFT


def _parece_numero(texto: str) -> bool:
    limpo = texto.replace("R$", "").replace(" ", "").replace("%", "")
    return bool(re.fullmatch(r"-?\d{1,3}(\.\d{3})*,\d{2}", limpo) or re.fullmatch(r"-?\d+(,\d+)?", limpo))


def _cor_fonte(celula):
    cor = celula.font.color
    if cor is None or not getattr(cor, "type", None):
        return (0, 0, 0)
    if cor.type == "rgb" and cor.rgb:
        rgb = str(cor.rgb).upper()
        if len(rgb) == 8:
            if rgb.startswith("00"):
                return (0, 0, 0)
            rgb = rgb[2:]
        try:
            return tuple(int(rgb[indice : indice + 2], 16) / 255 for indice in (0, 2, 4))
        except ValueError:
            return (0, 0, 0)
    return (0, 0, 0)


def _registrar_fonte(pagina, fontes: dict, nome: str, negrito: bool) -> str:
    chave = ((nome or "Arial").lower(), bool(negrito))
    if chave in fontes:
        return fontes[chave]
    caminho = _caminho_fonte(*chave)
    interno = f"F{len(fontes)}"
    if caminho is not None:
        pagina.insert_font(fontname=interno, fontfile=str(caminho))
        fontes[chave] = interno
        return interno
    interno = "hebo" if negrito else "helv"
    fontes[chave] = interno
    return interno


def _caminho_fonte(nome: str, negrito: bool) -> Path | None:
    arquivos = {
        ("arial", False): "arial.ttf",
        ("arial", True): "arialbd.ttf",
        ("calibri", False): "calibri.ttf",
        ("calibri", True): "calibrib.ttf",
        ("times new roman", False): "times.ttf",
        ("times new roman", True): "timesbd.ttf",
        ("cambria", False): "times.ttf",
        ("cambria", True): "timesbd.ttf",
        ("liberation sans", False): "LiberationSans-Regular.ttf",
        ("liberation sans", True): "LiberationSans-Bold.ttf",
        ("ubuntu", False): "arial.ttf",
        ("ubuntu", True): "arialbd.ttf",
    }
    arquivo = arquivos.get((nome, negrito)) or arquivos.get((nome, False))
    if negrito and arquivo is None:
        arquivo = "arialbd.ttf"
    if arquivo is None:
        arquivo = "arial.ttf"
    pastas = [
        Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
        Path("/usr/share/fonts/truetype/dejavu"),
        Path("/usr/share/fonts/truetype/liberation"),
        Path("/usr/share/fonts/truetype/liberation2"),
    ]
    if arquivo == "arial.ttf":
        alternativas = [arquivo, "DejaVuSans.ttf", "LiberationSans-Regular.ttf"]
    elif arquivo == "arialbd.ttf":
        alternativas = [arquivo, "DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf"]
    else:
        alternativas = [arquivo]
    for pasta in pastas:
        for candidato in alternativas:
            caminho = pasta / candidato
            if caminho.is_file():
                return caminho
    return None


def _imagens_planilha(folha) -> list[dict]:
    imagens = []
    for imagem in getattr(folha, "_images", []):
        ancora = imagem.anchor
        origem = getattr(ancora, "_from", None)
        if origem is None:
            continue
        destino = getattr(ancora, "to", None)
        imagens.append(
            {
                "dados": imagem._data(),
                "col": int(origem.col),
                "lin": int(origem.row),
                "col_off": int(getattr(origem, "colOff", 0) or 0),
                "lin_off": int(getattr(origem, "rowOff", 0) or 0),
                "col2": int(destino.col) if destino is not None else int(origem.col) + 1,
                "lin2": int(destino.row) if destino is not None else int(origem.row) + 1,
                "col2_off": int(getattr(destino, "colOff", 0) or 0) if destino is not None else 0,
                "lin2_off": int(getattr(destino, "rowOff", 0) or 0) if destino is not None else 0,
            }
        )
    return imagens


def _desenhar_imagem_planilha(
    pagina,
    imagem: dict,
    absolutos_x,
    absolutos_y,
    min_col: int,
    linha_inicial: int,
    linha_final: int,
    origem_x: float,
    origem_y: float,
    escala: float,
) -> None:
    linha = imagem["lin"] + 1
    if linha < linha_inicial or linha > linha_final:
        return
    base_x = absolutos_x[min_col - 1]
    base_y = absolutos_y[linha_inicial - 1]

    def coordenada(coluna: int, deslocamento: int, linha_zero: int, deslocamento_linha: int):
        x = absolutos_x[coluna] + deslocamento / 12700 - base_x
        y = absolutos_y[linha_zero] + deslocamento_linha / 12700 - base_y
        return x, y

    x0, y0 = coordenada(imagem["col"], imagem["col_off"], imagem["lin"], imagem["lin_off"])
    x1, y1 = coordenada(imagem["col2"], imagem["col2_off"], imagem["lin2"], imagem["lin2_off"])
    retangulo = pymupdf.Rect(
        origem_x + min(x0, x1) * escala,
        origem_y + min(y0, y1) * escala,
        origem_x + max(x0, x1) * escala,
        origem_y + max(y0, y1) * escala,
    )
    if retangulo.width < 2 or retangulo.height < 2:
        return
    pagina.insert_image(retangulo, stream=imagem["dados"], keep_proportion=True)


def _pdf_word(origem: Path) -> bytes:
    documento = Document(str(origem))
    secao = documento.sections[0]
    largura = float(secao.page_width.pt)
    altura = float(secao.page_height.pt)
    esquerda = float(secao.left_margin.pt)
    direita = float(secao.right_margin.pt)
    topo = float(secao.top_margin.pt)
    baixo = float(secao.bottom_margin.pt)
    imagens = _imagens_docx(origem)
    pdf = pymupdf.open()
    try:
        pagina = pdf.new_page(width=largura, height=altura)
        fontes: dict[tuple[str, bool], str] = {}
        y = topo
        limite = altura - baixo
        for bloco in documento.element.body:
            nome = etree.QName(bloco).localname
            if nome == "p":
                y = _paragrafo_word(
                    pagina,
                    Paragraph(bloco, documento),
                    fontes,
                    esquerda,
                    y,
                    largura - direita,
                    limite,
                )
            elif nome == "tbl":
                y = _tabela_word(
                    pagina,
                    Table(bloco, documento),
                    fontes,
                    imagens,
                    esquerda,
                    y,
                    largura - direita,
                )
            if y > limite:
                _rodape_word(pagina, secao, fontes)
                pagina = pdf.new_page(width=largura, height=altura)
                fontes = {}
                y = topo
        _rodape_word(pagina, secao, fontes)
        return pdf.tobytes(deflate=True, garbage=3)
    finally:
        pdf.close()


def _paragrafo_word(pagina, paragrafo: Paragraph, fontes, x0, y, x1, limite) -> float:
    texto = paragrafo.text.strip()
    tamanho = _tamanho_fonte(paragrafo)
    antes = _espaco(paragrafo.paragraph_format.space_before)
    depois = _espaco(paragrafo.paragraph_format.space_after)
    y += antes
    if not texto:
        return y + max(tamanho, 12) * 0.85 + depois
    negrito = any(run.bold for run in paragrafo.runs)
    nome = next((run.font.name for run in paragrafo.runs if run.font.name), "Arial")
    fontname = _registrar_fonte(pagina, fontes, nome, bool(negrito))
    alinhamento = {
        1: pymupdf.TEXT_ALIGN_CENTER,
        2: pymupdf.TEXT_ALIGN_RIGHT,
        3: pymupdf.TEXT_ALIGN_JUSTIFY,
    }.get(int(paragrafo.alignment or 0), pymupdf.TEXT_ALIGN_LEFT)
    linhas = max(1, math.ceil((len(texto) * tamanho * 0.5) / max(x1 - x0, 20)))
    altura = linhas * tamanho * 1.45 + 2
    if y + altura > limite:
        return limite + 1
    caixa = pymupdf.Rect(x0, y, x1, y + altura)
    pagina.insert_textbox(
        caixa,
        texto,
        fontsize=tamanho,
        fontname=fontname,
        align=alinhamento,
        color=(0, 0, 0),
    )
    return y + altura + depois


def _tamanho_fonte(paragrafo: Paragraph) -> float:
    for run in paragrafo.runs:
        if run.font.size:
            return float(run.font.size.pt)
    return 12.0


def _espaco(valor) -> float:
    if valor is None:
        return 0.0
    try:
        return float(valor.pt)
    except AttributeError:
        return 0.0


def _tabela_word(pagina, tabela: Table, fontes, imagens, x0, y, x1) -> float:
    larguras = _larguras_tabela(tabela)
    disponivel = max(x1 - x0, 20)
    total = sum(larguras) or disponivel
    fator = min(1.0, disponivel / total)
    larguras = [item * fator for item in larguras]
    linhas = [_celulas_da_linha(linha) for linha in tabela.rows]
    alturas = [max(_altura_linha_word(linha), 12) for linha in tabela.rows]
    for indice, celulas in enumerate(linhas):
        coluna = 0
        x = x0
        for celula, alcance, vmerge in celulas:
            largura = sum(larguras[coluna : coluna + alcance]) if larguras else disponivel
            if vmerge == "continue":
                x += largura
                coluna += alcance
                continue
            repeticoes = _linhas_mescladas(linhas, indice, coluna) if vmerge == "restart" else 1
            altura = sum(alturas[indice : indice + repeticoes])
            retangulo = pymupdf.Rect(x, y, x + largura, y + altura)
            pagina.draw_rect(retangulo, color=(0, 0, 0), width=0.6)
            _imagem_na_celula(pagina, celula, imagens, retangulo)
            texto = celula.text.strip()
            if texto:
                fontname = _registrar_fonte(pagina, fontes, "Arial", False)
                tamanho = min(9.0, max(6.0, (altura - 3) / 1.25))
                pagina.insert_textbox(
                    pymupdf.Rect(retangulo.x0 + 2, retangulo.y0 + 1, retangulo.x1 - 2, retangulo.y1 - 1),
                    texto,
                    fontsize=tamanho,
                    fontname=fontname,
                    align=pymupdf.TEXT_ALIGN_LEFT,
                    color=(0, 0, 0),
                    lineheight=1.05,
                )
            x += largura
            coluna += alcance
        y += alturas[indice]
    return y + 8


def _linhas_mescladas(linhas, indice: int, coluna: int) -> int:
    total = 1
    for seguinte in linhas[indice + 1 :]:
        cursor = 0
        continua = False
        for _celula, alcance, vmerge in seguinte:
            if cursor == coluna:
                continua = vmerge == "continue"
                break
            cursor += alcance
        if not continua:
            break
        total += 1
    return total


def _celulas_da_linha(linha):
    vistos = []
    saida = []
    for celula in linha.cells:
        if celula._tc in vistos:
            continue
        vistos.append(celula._tc)
        alcance = 1
        propriedades = celula._tc.tcPr
        vmerge = None
        if propriedades is not None:
            if propriedades.gridSpan is not None:
                alcance = int(propriedades.gridSpan.val)
            if propriedades.vMerge is not None:
                valor = propriedades.vMerge.val
                vmerge = "continue" if not valor or valor == "continue" else "restart"
        saida.append((celula, alcance, vmerge))
    return saida


def _larguras_tabela(tabela: Table) -> list[float]:
    grade = tabela._tbl.tblGrid
    if grade is None:
        return []
    larguras = []
    for coluna in grade.gridCol_lst:
        valor = float(coluna.w or 0)
        larguras.append(valor / 12700 if valor > 5000 else valor / 20)
    return larguras


def _altura_linha_word(linha) -> float:
    propriedades = linha._tr.trPr
    if propriedades is None or propriedades.trHeight is None:
        return 18.0
    valor = float(propriedades.trHeight.val)
    return valor / 12700 if valor > 5000 else valor / 20


def _imagens_docx(origem: Path) -> dict[str, bytes]:
    imagens = {}
    with zipfile.ZipFile(origem) as pacote:
        relacoes = "word/_rels/document.xml.rels"
        if relacoes not in pacote.namelist():
            return imagens
        raiz = etree.fromstring(pacote.read(relacoes))
        for relacao in raiz:
            if not (relacao.get("Type") or "").endswith("/image"):
                continue
            alvo = relacao.get("Target") or ""
            if alvo.startswith("/"):
                nome = alvo.lstrip("/")
            elif alvo.startswith("word/"):
                nome = alvo
            else:
                nome = "word/" + alvo
            if nome in pacote.namelist():
                imagens[relacao.get("Id")] = pacote.read(nome)
    return imagens


def _imagem_na_celula(pagina, celula, imagens: dict[str, bytes], retangulo) -> None:
    for blip in celula._tc.findall(f".//{{{_NS_A}}}blip"):
        identificador = blip.get(f"{{{_NS_REL}}}embed")
        dados = imagens.get(identificador or "")
        if not dados:
            continue
        area = pymupdf.Rect(
            retangulo.x0 + 4,
            retangulo.y0 + 3,
            retangulo.x1 - 4,
            retangulo.y1 - 3,
        )
        pagina.insert_image(area, stream=dados, keep_proportion=True)


def _rodape_word(pagina, secao, fontes) -> None:
    textos = [paragrafo.text.strip() for paragrafo in secao.footer.paragraphs if paragrafo.text.strip()]
    if not textos:
        return
    y = pagina.rect.height - float(secao.bottom_margin.pt) + 6
    pagina.draw_line(
        (float(secao.left_margin.pt), y),
        (pagina.rect.width - float(secao.right_margin.pt), y),
        color=(0, 0, 0),
        width=0.6,
    )
    fontname = _registrar_fonte(pagina, fontes, "Arial", False)
    pagina.insert_textbox(
        pymupdf.Rect(float(secao.left_margin.pt), y + 2, pagina.rect.width - float(secao.right_margin.pt), y + 28),
        "\n".join(textos),
        fontsize=9,
        fontname=fontname,
        align=pymupdf.TEXT_ALIGN_CENTER,
        color=(0, 0, 0),
    )


def _linhas_office(origem: Path, tipo: str) -> list[str]:
    if tipo == "docx":
        documento = Document(str(origem))
        return [paragrafo.text for paragrafo in documento.paragraphs if paragrafo.text.strip()]
    from openpyxl import load_workbook

    livro = load_workbook(origem, data_only=False)
    linhas = []
    for folha in livro.worksheets:
        if folha.sheet_state != "visible":
            continue
        for linha in folha.iter_rows(max_row=min(folha.max_row or 1, 80), max_col=12):
            textos = [str(celula.value) for celula in linha if celula.value not in (None, "")]
            if textos:
                linhas.append(" | ".join(textos))
    livro.close()
    return linhas


def _pdf_texto(linhas: list[str]) -> bytes:
    """PDF mínimo com o texto do documento, quando o desenho também falha."""
    pdf = pymupdf.open()
    try:
        fonte = _caminho_fonte("arial", False)
        conteudo = linhas or ["Documento"]
        pagina = None
        y = 800.0
        for linha in conteudo:
            if pagina is None or y > 800:
                pagina = pdf.new_page(width=_PAPEL[9][0], height=_PAPEL[9][1])
                if fonte is not None:
                    pagina.insert_font(fontname="FTexto", fontfile=str(fonte))
                y = 48
            pagina.insert_text(
                (40, y),
                (linha or "")[:180],
                fontsize=10,
                fontname="FTexto" if fonte is not None else "helv",
                color=(0, 0, 0),
            )
            y += 14
        return pdf.tobytes(deflate=True)
    finally:
        pdf.close()

