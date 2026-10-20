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
_WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
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
    if tipo == "docx":
        try:
            return _pdf_word(origem)
        except ConversaoPdfError:
            raise
        except Exception:
            return _pdf_texto(_linhas_office(origem, tipo))
    if tipo == "xlsx":
        if _weasy_disponivel():
            try:
                return _pdf_weasy(_html_planilha(origem))
            except Exception:
                pass
        try:
            return _pdf_planilha(origem)
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


def _registrar_fonte(pagina, fontes: dict, nome: str, negrito: bool, italico: bool = False) -> str:
    chave = ((nome or "Arial").lower(), bool(negrito), bool(italico))
    if chave in fontes:
        return fontes[chave]
    caminho = _caminho_fonte(chave[0], negrito, italico)
    interno = f"F{len(fontes)}"
    if caminho is not None:
        try:
            pagina.insert_font(fontname=interno, fontfile=str(caminho))
        except Exception:
            caminho = None
        else:
            fontes[chave] = interno
            return interno
    interno = "hebo" if negrito else "helv"
    fontes[chave] = interno
    return interno


def _caminho_fonte(nome: str, negrito: bool, italico: bool = False) -> Path | None:
    apelidos = {
        "times": "times new roman",
        "liberation serif": "times new roman",
        "liberation sans": "arial",
        "noto serif cjk sc": "times new roman",
        "noto sans devanagari": "arial",
    }
    nome = apelidos.get((nome or "").lower(), (nome or "").lower())
    arquivos = {
        ("arial", False, False): "arial.ttf",
        ("arial", True, False): "arialbd.ttf",
        ("arial", False, True): "ariali.ttf",
        ("arial", True, True): "arialbi.ttf",
        ("arial narrow", False, False): "arialn.ttf",
        ("arial narrow", True, False): "arialnb.ttf",
        ("arial narrow", False, True): "arialni.ttf",
        ("arial narrow", True, True): "arialnbi.ttf",
        ("calibri", False, False): "calibri.ttf",
        ("calibri", True, False): "calibrib.ttf",
        ("calibri", False, True): "calibrii.ttf",
        ("calibri", True, True): "calibriz.ttf",
        ("times new roman", False, False): "times.ttf",
        ("times new roman", True, False): "timesbd.ttf",
        ("times new roman", False, True): "timesi.ttf",
        ("times new roman", True, True): "timesbi.ttf",
        ("cambria", False, False): "cambria.ttc",
        ("cambria", True, False): "cambriab.ttf",
        ("cambria", False, True): "cambriai.ttf",
        ("cambria", True, True): "cambriaz.ttf",
        ("ubuntu", False, False): "arial.ttf",
        ("ubuntu", True, False): "arialbd.ttf",
    }
    arquivo = (
        arquivos.get((nome, negrito, italico))
        or arquivos.get((nome, negrito, False))
        or arquivos.get((nome, False, False))
    )
    if arquivo is None:
        arquivo = "arialbd.ttf" if negrito else "arial.ttf"
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
    elif arquivo == "times.ttf":
        alternativas = [arquivo, "LiberationSerif-Regular.ttf", "DejaVuSerif.ttf"]
    elif arquivo == "timesbd.ttf":
        alternativas = [arquivo, "LiberationSerif-Bold.ttf", "DejaVuSerif-Bold.ttf"]
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
    """Desenha o DOCX armazenado, sem reconstruí-lo a partir de outro modelo."""
    with zipfile.ZipFile(origem) as pacote:
        if "word/document.xml" not in pacote.namelist():
            raise ConversaoPdfError(_MENSAGEM)
        documento = etree.fromstring(pacote.read("word/document.xml"))
        base, estilos = _estilos_docx(pacote)
        imagens = _imagens_docx(origem)
        rodape = _paragrafos_rodape(pacote)
    corpo = documento.find(_w("body"))
    if corpo is None:
        raise ConversaoPdfError(_MENSAGEM)
    pdf = pymupdf.open()
    try:
        desenho = _DocxFiel(pdf, _geometria_docx(corpo), base, estilos, imagens)
        desenho.renderizar(corpo, rodape)
        return pdf.tobytes(deflate=True, garbage=3)
    finally:
        pdf.close()


def _w(tag: str) -> str:
    return f"{{{_WML}}}{tag}"


def _attr(elemento, nome: str):
    if elemento is None:
        return None
    return elemento.get(_w(nome))


def _twips(valor) -> float:
    return int(valor) / 20


def _ligado(elemento, tag: str):
    nodo = elemento.find(_w(tag)) if elemento is not None else None
    if nodo is None:
        return None
    valor = _attr(nodo, "val")
    if valor is None:
        return True
    return str(valor).lower() not in {"0", "false", "off", "none"}


def _cor_hex(valor: str | None):
    if not valor or valor == "auto":
        return None
    texto = valor[-6:]
    try:
        return tuple(int(texto[indice : indice + 2], 16) / 255 for indice in (0, 2, 4))
    except ValueError:
        return None


def _ler_rpr(elemento) -> dict:
    if elemento is None:
        return {}
    dados = {}
    fontes = elemento.find(_w("rFonts"))
    if fontes is not None:
        nome = _attr(fontes, "ascii") or _attr(fontes, "hAnsi") or _attr(fontes, "cs")
        if nome:
            dados["fonte"] = nome
    for tag, chave in (("b", "negrito"), ("i", "italico"), ("smallCaps", "maiusculas")):
        flag = _ligado(elemento, tag)
        if flag is not None:
            dados[chave] = flag
    tamanho = elemento.find(_w("sz"))
    if tamanho is not None and _attr(tamanho, "val"):
        dados["tamanho"] = int(_attr(tamanho, "val")) / 2
    cor = _cor_hex(_attr(elemento.find(_w("color")), "val"))
    if cor is not None:
        dados["cor"] = cor
    return dados


def _ler_ppr(elemento) -> dict:
    if elemento is None:
        return {}
    dados = {}
    alinhamento = elemento.find(_w("jc"))
    if alinhamento is not None and _attr(alinhamento, "val"):
        dados["alinhamento"] = _attr(alinhamento, "val")
    espaco = elemento.find(_w("spacing"))
    if espaco is not None:
        for origem, destino in (("before", "antes"), ("after", "depois"), ("line", "linha")):
            if _attr(espaco, origem):
                dados[destino] = int(_attr(espaco, origem))
        if _attr(espaco, "lineRule"):
            dados["regra"] = _attr(espaco, "lineRule")
    recuo = elemento.find(_w("ind"))
    if recuo is not None:
        inicio = _attr(recuo, "start") or _attr(recuo, "left") or "0"
        pendente = _attr(recuo, "hanging")
        if pendente:
            dados["recuo"] = (int(inicio) - int(pendente)) / 20
        elif _attr(recuo, "start") or _attr(recuo, "left"):
            dados["recuo"] = int(inicio) / 20
    estilo = elemento.find(_w("pStyle"))
    if estilo is not None and _attr(estilo, "val"):
        dados["estilo"] = _attr(estilo, "val")
    return dados


def _estilos_docx(pacote) -> tuple[dict, dict]:
    if "word/styles.xml" not in pacote.namelist():
        return {}, {}
    raiz = etree.fromstring(pacote.read("word/styles.xml"))
    padrao = raiz.find(f"{_w('docDefaults')}/{_w('rPrDefault')}/{_w('rPr')}")
    base = _ler_rpr(padrao)
    estilos = {}
    for estilo in raiz.findall(_w("style")):
        identificador = _attr(estilo, "styleId")
        if not identificador:
            continue
        base_em = estilo.find(_w("basedOn"))
        estilos[identificador] = {
            "based": _attr(base_em, "val"),
            "r": _ler_rpr(estilo.find(_w("rPr"))),
            "p": _ler_ppr(estilo.find(_w("pPr"))),
        }
    return base, estilos


def _cadeia(estilos: dict, identificador: str | None, campo: str) -> dict:
    vistos = set()
    ordem = []
    while identificador and identificador not in vistos and identificador in estilos:
        vistos.add(identificador)
        ordem.append(identificador)
        identificador = estilos[identificador]["based"]
    resultado = {}
    for item in reversed(ordem):
        resultado.update(estilos[item][campo])
    return resultado


def _geometria_docx(corpo) -> dict:
    secao = corpo.find(_w("sectPr"))
    caixa = {"largura": 612.0, "altura": 792.0, "esquerda": 72.0, "topo": 72.0, "direita": 72.0, "baixo": 72.0, "rodape": 36.0}
    if secao is None:
        return caixa
    tamanho = secao.find(_w("pgSz"))
    if tamanho is not None and _attr(tamanho, "w") and _attr(tamanho, "h"):
        caixa["largura"] = _twips(_attr(tamanho, "w"))
        caixa["altura"] = _twips(_attr(tamanho, "h"))
        if _attr(tamanho, "orient") == "landscape" and caixa["largura"] < caixa["altura"]:
            caixa["largura"], caixa["altura"] = caixa["altura"], caixa["largura"]
    margens = secao.find(_w("pgMar"))
    if margens is not None:
        for chave, campo in (
            ("left", "esquerda"),
            ("right", "direita"),
            ("top", "topo"),
            ("bottom", "baixo"),
            ("footer", "rodape"),
        ):
            if _attr(margens, chave):
                caixa[campo] = _twips(_attr(margens, chave))
    return caixa


def _paragrafos_rodape(pacote) -> list:
    relacoes = "word/_rels/document.xml.rels"
    if relacoes not in pacote.namelist():
        return []
    raiz = etree.fromstring(pacote.read(relacoes))
    for relacao in raiz:
        if not (relacao.get("Type") or "").endswith("/footer"):
            continue
        alvo = relacao.get("Target") or ""
        nome = alvo if alvo.startswith("word/") else "word/" + alvo.lstrip("/")
        if nome not in pacote.namelist():
            continue
        return etree.fromstring(pacote.read(nome)).findall(_w("p"))
    return []


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


def _formato_base() -> dict:
    return {
        "fonte": "Times New Roman",
        "tamanho": 12.0,
        "negrito": False,
        "italico": False,
        "cor": (0, 0, 0),
        "maiusculas": False,
    }


def _ancoras(paragrafo) -> list[dict]:
    imagens = []
    for ancora in paragrafo.findall(f".//{{{_WP}}}anchor"):
        extensao = ancora.find(f"{{{_WP}}}extent")
        blip = ancora.find(f".//{{{_NS_A}}}blip")
        if extensao is None or blip is None:
            continue
        horizontal = ancora.find(f"{{{_WP}}}positionH/{{{_WP}}}posOffset")
        vertical = ancora.find(f"{{{_WP}}}positionV/{{{_WP}}}posOffset")
        imagens.append(
            {
                "id": blip.get(f"{{{_NS_REL}}}embed"),
                "x": int(horizontal.text or 0) / 12700 if horizontal is not None else 0,
                "y": int(vertical.text or 0) / 12700 if vertical is not None else 0,
                "largura": int(extensao.get("cx") or 0) / 12700,
                "altura": int(extensao.get("cy") or 0) / 12700,
            }
        )
    return imagens


def _bordas_celula(propriedades) -> dict:
    elemento = propriedades.find(_w("tcBorders")) if propriedades is not None else None
    if elemento is None:
        return {}
    nomes = {"top": "top", "bottom": "bottom", "start": "left", "left": "left", "end": "right", "right": "right"}
    bordas = {}
    for tag, lado in nomes.items():
        borda = elemento.find(_w(tag))
        if borda is None:
            continue
        estilo = _attr(borda, "val") or "single"
        if estilo in {"nil", "none"}:
            continue
        cor = _cor_hex(_attr(borda, "color")) or (0, 0, 0)
        bordas[lado] = {"estilo": estilo, "sz": int(_attr(borda, "sz") or 4), "cor": cor}
    return bordas


class _DocxFiel:
    """Reproduz o DOCX guardado: tabelas, bordas, fontes e imagens do próprio arquivo."""

    def __init__(self, pdf, caixa: dict, base: dict, estilos: dict, imagens: dict[str, bytes]):
        self.pdf = pdf
        self.caixa = caixa
        self.base = base
        self.estilos = estilos
        self.imagens = imagens
        self.medidores: dict = {}
        self.pagina = None
        self.fontes: dict = {}
        self.paginas = []
        self.y = 0.0

    def renderizar(self, corpo, rodape) -> None:
        self._nova_pagina()
        for bloco in corpo:
            nome = etree.QName(bloco).localname
            if nome == "p":
                self._fluir_paragrafo(bloco, self._largura_util())
            elif nome == "tbl":
                self._fluir_tabela(bloco)
        for pagina in self.paginas:
            self._desenhar_rodape(pagina, rodape)

    def _nova_pagina(self) -> None:
        self.pagina = self.pdf.new_page(width=self.caixa["largura"], height=self.caixa["altura"])
        self.fontes = {}
        self.paginas.append(self.pagina)
        self.y = self.caixa["topo"]

    def _limite(self) -> float:
        return self.caixa["altura"] - self.caixa["baixo"]

    def _largura_util(self) -> float:
        return max(self.caixa["largura"] - self.caixa["esquerda"] - self.caixa["direita"], 40)

    def _fluir_paragrafo(self, elemento, largura: float) -> None:
        bloco = self._preparar_paragrafo(elemento, largura)
        if self.y + bloco["altura"] > self._limite() and self.y > self.caixa["topo"] + 8:
            self._nova_pagina()
        self._desenhar_paragrafo(self.pagina, bloco, self.caixa["esquerda"], self.y)
        self.y += bloco["altura"]

    def _fluir_tabela(self, tabela) -> None:
        preparada = self._preparar_tabela(tabela)
        if self.y + preparada["altura"] > self._limite() and self.y > self.caixa["topo"] + 8:
            self._nova_pagina()
        self.y = self._desenhar_tabela(self.pagina, preparada, self.caixa["esquerda"], self.y)

    def _preparar_paragrafo(self, elemento, largura: float) -> dict:
        propriedades = elemento.find(_w("pPr"))
        direto = _ler_ppr(propriedades)
        paragrafo = _cadeia(self.estilos, direto.get("estilo"), "p")
        paragrafo.update({chave: valor for chave, valor in direto.items() if chave != "estilo"})
        paragrafo.setdefault("alinhamento", "left")
        paragrafo.setdefault("antes", 0)
        paragrafo.setdefault("depois", 0)
        paragrafo.setdefault("recuo", 0)
        rpr = dict(self.base)
        rpr.update(_cadeia(self.estilos, direto.get("estilo"), "r"))
        if propriedades is not None:
            rpr.update(_ler_rpr(propriedades.find(_w("rPr"))))
        trechos = []
        for run in self._runs(elemento):
            fmt = _formato_base()
            fmt.update(rpr)
            fmt.update(_ler_rpr(run.find(_w("rPr"))))
            for neto in run:
                tag = etree.QName(neto).localname
                if tag == "br":
                    trechos.append(("\n", fmt))
                elif tag == "tab":
                    trechos.append(("    ", fmt))
                elif tag == "t" and neto.text:
                    texto = neto.text.upper() if fmt["maiusculas"] else neto.text
                    trechos.append((texto, fmt))
        tamanho = max((fmt["tamanho"] for _, fmt in trechos), default=fmt_padrao(rpr))
        if not any(texto.strip() for texto, _ in trechos):
            linhas = [[]]
        else:
            linhas = self._quebrar(trechos, max(largura - paragrafo["recuo"], 12))
        altura_linha = _altura_da_linha(paragrafo, tamanho)
        return {
            "linhas": linhas,
            "alinhamento": paragrafo["alinhamento"],
            "antes": paragrafo["antes"] / 20,
            "depois": paragrafo["depois"] / 20,
            "recuo": paragrafo["recuo"],
            "altura_linha": altura_linha,
            "altura": paragrafo["antes"] / 20 + max(len(linhas), 1) * altura_linha + paragrafo["depois"] / 20,
            "largura": largura,
        }

    def _runs(self, elemento):
        for filho in elemento:
            tag = etree.QName(filho).localname
            if tag == "r":
                yield filho
            elif tag in {"hyperlink", "ins"}:
                yield from self._runs(filho)

    def _quebrar(self, trechos, largura: float) -> list[list]:
        linhas = [[]]
        usado = 0.0
        for texto, fmt in trechos:
            partes = texto.split("\n")
            for indice, parte in enumerate(partes):
                if indice:
                    linhas.append([])
                    usado = 0.0
                for token in re.findall(r"\s+|\S+", parte):
                    ocupacao = self._medir(token, fmt)
                    if linhas[-1] and token.strip() and usado + ocupacao > largura + 0.5:
                        linhas.append([])
                        usado = 0.0
                        if not token.strip():
                            continue
                    linhas[-1].append((token, fmt, ocupacao))
                    usado += ocupacao
        return linhas

    def _medir(self, texto: str, fmt: dict) -> float:
        caminho = _caminho_fonte(fmt["fonte"], fmt["negrito"], fmt["italico"])
        chave = str(caminho) if caminho is not None else "helv"
        if chave not in self.medidores:
            try:
                self.medidores[chave] = (
                    pymupdf.Font(fontfile=str(caminho)) if caminho is not None else pymupdf.Font("helv")
                )
            except Exception:
                self.medidores[chave] = pymupdf.Font("helv")
        return self.medidores[chave].text_length(texto, fontsize=fmt["tamanho"])

    def _desenhar_paragrafo(self, pagina, bloco: dict, x: float, y: float) -> None:
        y += bloco["antes"]
        for indice, linha in enumerate(bloco["linhas"]):
            self._desenhar_linha(
                pagina,
                linha,
                x + (bloco["recuo"] if indice == 0 else 0),
                y,
                bloco["largura"] - (bloco["recuo"] if indice == 0 else 0),
                bloco["alinhamento"],
                bloco["altura_linha"],
                indice == len(bloco["linhas"]) - 1,
            )
            y += bloco["altura_linha"]

    def _desenhar_linha(self, pagina, linha, x, y, largura, alinhamento, altura_linha, ultima) -> None:
        if not linha:
            return
        usada = sum(item[2] for item in linha)
        escala = min(1.0, largura / usada) if usada > largura > 0 else 1.0
        if alinhamento in {"center", "right"} or (alinhamento in {"both", "justify"} and ultima):
            folga = max(largura - usada, 0)
            cursor = x + (folga / 2 if alinhamento == "center" else folga if alinhamento == "right" else 0)
            extra = 0
        elif alinhamento in {"both", "justify"}:
            espacos = sum(token.count(" ") for token, _, _ in linha)
            extra = max(largura - usada, 0) / espacos if espacos else 0
            cursor = x
        else:
            cursor = x
            extra = 0
        tamanho = max(fmt["tamanho"] for _, fmt, _ in linha)
        baseline = y + min(tamanho * 0.82, altura_linha * 0.86)
        for token, fmt, ocupacao in linha:
            nome = _registrar_fonte(pagina, self.fontes, fmt["fonte"], fmt["negrito"], fmt["italico"])
            pagina.insert_text(
                (cursor, baseline),
                token,
                fontsize=fmt["tamanho"] * escala,
                fontname=nome,
                color=fmt["cor"],
            )
            cursor += (ocupacao * escala) + (token.count(" ") * extra)

    def _preparar_tabela(self, tabela) -> dict:
        grade = []
        grelha = tabela.find(_w("tblGrid"))
        if grelha is not None:
            for coluna in grelha.findall(_w("gridCol")):
                grade.append(_twips(_attr(coluna, "w") or 0))
        propriedades = tabela.find(_w("tblPr"))
        recuo = 0.0
        margens = {"top": 1.0, "bottom": 1.0, "left": 3.6, "right": 3.6}
        if propriedades is not None:
            indentacao = propriedades.find(_w("tblInd"))
            if indentacao is not None and _attr(indentacao, "w"):
                recuo = _twips(_attr(indentacao, "w"))
            margem = propriedades.find(_w("tblCellMar"))
            if margem is not None:
                for tag, lado in (("top", "top"), ("bottom", "bottom"), ("start", "left"), ("left", "left"), ("end", "right"), ("right", "right")):
                    item = margem.find(_w(tag))
                    if item is not None and _attr(item, "w"):
                        margens[lado] = _twips(_attr(item, "w"))
        linhas = []
        for tr in tabela.findall(_w("tr")):
            altura, regra = 0.0, "atLeast"
            tr_pr = tr.find(_w("trPr"))
            if tr_pr is not None and tr_pr.find(_w("trHeight")) is not None:
                marca = tr_pr.find(_w("trHeight"))
                if _attr(marca, "val"):
                    altura = _twips(_attr(marca, "val"))
                regra = _attr(marca, "hRule") or "atLeast"
            celulas = []
            coluna = 0
            for tc in tr.findall(_w("tc")):
                tc_pr = tc.find(_w("tcPr"))
                alcance = 1
                vmerge = None
                valign = "top"
                if tc_pr is not None:
                    span = tc_pr.find(_w("gridSpan"))
                    if span is not None and _attr(span, "val"):
                        alcance = int(_attr(span, "val"))
                    mescla = tc_pr.find(_w("vMerge"))
                    if mescla is not None:
                        vmerge = "restart" if _attr(mescla, "val") == "restart" else "continue"
                    vertical = tc_pr.find(_w("vAlign"))
                    if vertical is not None and _attr(vertical, "val"):
                        valign = _attr(vertical, "val")
                largura = sum(grade[coluna : coluna + alcance]) if grade else 80.0
                interna = max(largura - margens["left"] - margens["right"], 8)
                paragrafos = [self._preparar_paragrafo(p, interna) for p in tc.findall(_w("p"))]
                figuras = []
                for paragrafo in tc.findall(_w("p")):
                    figuras.extend(_ancoras(paragrafo))
                celulas.append(
                    {
                        "alcance": alcance,
                        "vmerge": vmerge,
                        "valign": valign,
                        "bordas": _bordas_celula(tc_pr),
                        "paragrafos": paragrafos,
                        "figuras": figuras,
                        "conteudo": sum(item["altura"] for item in paragrafos),
                    }
                )
                coluna += alcance
            linhas.append({"altura": altura, "regra": regra, "celulas": celulas})
        self._ajustar_alturas(linhas, margens)
        return {"recuo": recuo, "margens": margens, "grade": grade, "linhas": linhas, "altura": sum(item["altura"] for item in linhas)}

    def _ajustar_alturas(self, linhas, margens) -> None:
        for indice, linha in enumerate(linhas):
            coluna = 0
            for celula in linha["celulas"]:
                if celula["vmerge"] == "continue":
                    coluna += celula["alcance"]
                    continue
                repetidas = _linhas_verticais(linhas, indice, coluna)
                necessario = celula["conteudo"] + margens["top"] + margens["bottom"]
                if repetidas == 1 and linha["regra"] != "exact":
                    linha["altura"] = max(linha["altura"], necessario)
                elif repetidas > 1:
                    total = sum(linhas[indice + salto]["altura"] for salto in range(repetidas))
                    if necessario > total:
                        extra = necessario - total
                        for salto in range(repetidas - 1, -1, -1):
                            alvo = linhas[indice + salto]
                            if alvo["regra"] != "exact":
                                alvo["altura"] += extra
                                break
                coluna += celula["alcance"]
            if linha["altura"] <= 0:
                linha["altura"] = 14

    def _desenhar_tabela(self, pagina, tabela: dict, x: float, y: float) -> float:
        margens = tabela["margens"]
        origem_y = y
        arestas = {}
        for indice, linha in enumerate(tabela["linhas"]):
            cursor = x + tabela["recuo"]
            coluna = 0
            for celula in linha["celulas"]:
                largura = sum(tabela["grade"][coluna : coluna + celula["alcance"]]) if tabela["grade"] else 80
                if celula["vmerge"] == "continue":
                    cursor += largura
                    coluna += celula["alcance"]
                    continue
                repetidas = _linhas_verticais(tabela["linhas"], indice, coluna)
                altura = sum(tabela["linhas"][indice + salto]["altura"] for salto in range(repetidas))
                retangulo = pymupdf.Rect(cursor, y, cursor + largura, y + altura)
                livre = altura - margens["top"] - margens["bottom"] - celula["conteudo"]
                if celula["valign"] == "center":
                    texto_y = y + margens["top"] + max(livre, 0) / 2
                elif celula["valign"] == "bottom":
                    texto_y = y + altura - margens["bottom"] - celula["conteudo"]
                else:
                    texto_y = y + margens["top"]
                for paragrafo in celula["paragrafos"]:
                    self._desenhar_paragrafo(pagina, paragrafo, cursor + margens["left"], texto_y)
                    texto_y += paragrafo["altura"]
                for figura in celula["figuras"]:
                    dados = self.imagens.get(figura["id"] or "")
                    if not dados or figura["largura"] < 1 or figura["altura"] < 1:
                        continue
                    area = pymupdf.Rect(
                        retangulo.x0 + figura["x"],
                        retangulo.y0 + figura["y"],
                        retangulo.x0 + figura["x"] + figura["largura"],
                        retangulo.y0 + figura["y"] + figura["altura"],
                    )
                    pagina.insert_image(area, stream=dados, keep_proportion=True)
                for lado, borda in celula["bordas"].items():
                    chave = _chave_aresta(lado, retangulo)
                    if chave not in arestas or borda["sz"] >= arestas[chave][1]["sz"]:
                        arestas[chave] = (lado, borda, retangulo)
                cursor += largura
                coluna += celula["alcance"]
            y += linha["altura"]
        for lado, borda, retangulo in arestas.values():
            _tracar_borda(pagina, lado, retangulo, borda)
        return origem_y + tabela["altura"]

    def _desenhar_rodape(self, pagina, paragrafos) -> None:
        if not paragrafos:
            return
        self.pagina = pagina
        self.fontes = {}
        blocos = [self._preparar_paragrafo(item, self._largura_util()) for item in paragrafos]
        y = self.caixa["altura"] - self.caixa["rodape"] - sum(item["altura"] for item in blocos)
        for bloco in blocos:
            self._desenhar_paragrafo(pagina, bloco, self.caixa["esquerda"], y)
            y += bloco["altura"]


def fmt_padrao(rpr: dict) -> float:
    return float(rpr.get("tamanho") or 12)


def _altura_da_linha(paragrafo: dict, tamanho: float) -> float:
    linha = paragrafo.get("linha")
    regra = paragrafo.get("regra") or "auto"
    if linha and regra == "exact":
        return max(linha / 20, 8)
    if linha and regra == "auto":
        return max(tamanho * (linha / 240), 8)
    if linha and regra == "atLeast":
        return max(tamanho * 1.12, linha / 20)
    return max(tamanho * 1.12, 8)


def _linhas_verticais(linhas, indice: int, coluna: int) -> int:
    total = 1
    for seguinte in linhas[indice + 1 :]:
        cursor = 0
        continua = False
        for celula in seguinte["celulas"]:
            if cursor == coluna and celula["vmerge"] == "continue":
                continua = True
                break
            cursor += celula["alcance"]
        if not continua:
            break
        total += 1
    return total


def _chave_aresta(lado: str, retangulo) -> tuple:
    if lado in {"top", "bottom"}:
        y = retangulo.y0 if lado == "top" else retangulo.y1
        return ("h", round(y, 1), round(retangulo.x0, 1), round(retangulo.x1, 1))
    x = retangulo.x0 if lado == "left" else retangulo.x1
    return ("v", round(x, 1), round(retangulo.y0, 1), round(retangulo.y1, 1))


def _tracar_borda(pagina, lado: str, retangulo, borda: dict) -> None:
    estilo = borda["estilo"]
    cor = borda["cor"]
    total = max(borda["sz"] / 8, 0.5)
    if estilo == "double":
        for deslocamento in (-0.7, 0.7):
            _traco(pagina, lado, retangulo, cor, 0.65, deslocamento)
        return
    if estilo in {"thinThickSmallGap", "thickThinSmallGap"}:
        fino = max(total * 0.18, 0.45)
        vao = max(total * 0.22, 0.55)
        grosso = max(total - fino - vao, 0.9)
        partes = (
            [(grosso, False), (vao, True), (fino, False)]
            if estilo == "thickThinSmallGap"
            else [(fino, False), (vao, True), (grosso, False)]
        )
        cursor = 0.0
        for parte, intervalo in partes:
            if intervalo:
                cursor += parte
                continue
            _traco(pagina, lado, retangulo, cor, parte, cursor + parte / 2)
            cursor += parte
        return
    _traco(pagina, lado, retangulo, cor, max(total, 0.55), max(total, 0.55) / 2)


def _traco(pagina, lado: str, retangulo, cor, largura: float, deslocamento: float) -> None:
    extra = max(deslocamento, 0)
    if lado == "top":
        y = retangulo.y0 + deslocamento
        inicio, fim = (retangulo.x0 - extra, y), (retangulo.x1 + extra, y)
    elif lado == "bottom":
        y = retangulo.y1 - deslocamento
        inicio, fim = (retangulo.x0 - extra, y), (retangulo.x1 + extra, y)
    elif lado == "left":
        x = retangulo.x0 + deslocamento
        inicio, fim = (x, retangulo.y0 - extra), (x, retangulo.y1 + extra)
    else:
        x = retangulo.x1 - deslocamento
        inicio, fim = (x, retangulo.y0 - extra), (x, retangulo.y1 + extra)
    pagina.draw_line(inicio, fim, color=cor, width=largura, lineCap=1)


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

