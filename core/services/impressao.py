"""Converte previsão de pagamento, CI, RPA e o atesto assinado em páginas PNG."""

import re
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path

from core.services.conversao_pdf import (
    ConversaoPdfError,
    converter_para_pdf,
    imagem_para_pdf,
    rasterizar_pdf,
    unir_pdfs,
)
from core.services.excel_template import gerar_excel_previsao, gerar_excel_rpa
from core.services.servico_documentos import montar_contexto_de_servico


class ImpressaoError(Exception):
    pass


_ABA_IR = re.compile(br"<sheet\b[^>]*name=\"IR-2024\"[^>]*/?>")


def _ocultar_aba_ir(conteudo: bytes) -> bytes:
    """A aba IR-2024 é tabela de apoio; a impressão fica só com o formulário RPA.

    Se ela já está oculta, o arquivo não é regravado.
    """
    with zipfile.ZipFile(BytesIO(conteudo)) as origem:
        if "xl/workbook.xml" not in origem.namelist():
            return conteudo
        xml = origem.read("xl/workbook.xml")
    marcador = _ABA_IR.search(xml)
    if marcador is None or b'state="hidden"' in marcador.group(0) or b"state='hidden'" in marcador.group(0):
        return conteudo
    tag = marcador.group(0)
    if b"state=" in tag:
        nova = re.sub(br'state="[^"]*"', b'state="hidden"', tag, count=1)
    elif tag.endswith(b"/>"):
        nova = tag[:-2] + b' state="hidden"/>'
    else:
        nova = tag[:-1] + b' state="hidden">'
    xml = xml[: marcador.start()] + nova + xml[marcador.end() :]
    saida = BytesIO()
    with zipfile.ZipFile(BytesIO(conteudo)) as origem, zipfile.ZipFile(
        saida, "w", compression=zipfile.ZIP_DEFLATED
    ) as destino:
        for info in origem.infolist():
            dados = xml if info.filename == "xl/workbook.xml" else origem.read(info.filename)
            destino.writestr(info, dados)
    return saida.getvalue()


def _preparar_atesto(servico, destino: Path) -> None:
    arquivo = servico.atesto_assinado
    with arquivo.open("rb") as handle:
        dados = handle.read()
    nome = (arquivo.name or "").lower()
    if nome.endswith(".pdf") or dados.startswith(b"%PDF"):
        destino.write_bytes(dados)
        return
    if nome.endswith((".png", ".jpg", ".jpeg")) or dados.startswith(b"\x89PNG") or dados.startswith(b"\xff\xd8"):
        destino.write_bytes(imagem_para_pdf(dados))
        return
    raise ImpressaoError("O atesto assinado precisa ser um PDF ou uma imagem.")


def paginas_png_documentos(documento) -> list[bytes]:
    """
    Gera as páginas da previsão de pagamento, da CI, do RPA e, se houver, do atesto.

    Cada item é o PNG de uma página, pronto para a tela de impressão.
    O atesto assinado só entra quando estiver anexado ao serviço.
    """
    if not documento.arquivo_pdf:
        raise ImpressaoError("A CI deste serviço ainda não foi gerada.")

    servico = documento.servico
    contexto = montar_contexto_de_servico(servico)
    previsao = gerar_excel_previsao(
        contexto=contexto,
        numero_documento=servico.numero_documento,
    )
    rpa = gerar_excel_rpa(
        contexto=contexto,
        numero_documento=servico.numero_rpa_documento,
    )

    with tempfile.TemporaryDirectory(prefix="atesto-impressao-") as pasta:
        base = Path(pasta)
        previsao_path = base / "01-previsao.xlsx"
        ci_path = base / "02-ci.docx"
        rpa_path = base / "03-rpa.xlsx"
        with documento.arquivo_pdf.open("rb") as arquivo_ci:
            ci_path.write_bytes(arquivo_ci.read())
        previsao_path.write_bytes(previsao.read())
        rpa_path.write_bytes(_ocultar_aba_ir(rpa.read()))

        pdf_dir = base / "pdf"
        pdf_dir.mkdir()
        atesto_pdf = None
        if getattr(servico, "atesto_assinado", None):
            atesto_pdf = pdf_dir / "04-atesto.pdf"
            try:
                _preparar_atesto(servico, atesto_pdf)
            except ImpressaoError:
                raise
            except OSError as exc:
                raise ImpressaoError("Não foi possível incluir o atesto assinado.") from exc

        pares = [
            (previsao_path, pdf_dir / "01-previsao.pdf"),
            (ci_path, pdf_dir / "02-ci.pdf"),
            (rpa_path, pdf_dir / "03-rpa.pdf"),
        ]
        try:
            converter_para_pdf(pares)
        except ConversaoPdfError as exc:
            raise ImpressaoError(str(exc)) from exc

        pdfs = [destino for _, destino in pares]
        if atesto_pdf is not None:
            pdfs.append(atesto_pdf)
        ausentes = [caminho.name for caminho in pdfs if not caminho.exists()]
        if ausentes:
            raise ImpressaoError(
                "Não foi possível converter todos os documentos: " + ", ".join(ausentes)
            )

        unido = base / "documentos.pdf"
        try:
            unir_pdfs(pdfs, unido)
            return rasterizar_pdf(unido)
        except ConversaoPdfError as exc:
            raise ImpressaoError(str(exc)) from exc
