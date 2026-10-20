"""Junta previsão de pagamento, CI, RPA e o atesto assinado em um PDF para impressão."""

import tempfile
import zipfile
from io import BytesIO
from pathlib import Path

from django.core.exceptions import ValidationError
from lxml import etree

from core.services.conversao_pdf import ConversaoPdfError, converter_para_pdf, unir_pdfs
from core.services.excel_template import gerar_excel_previsao, gerar_excel_rpa
from core.services.servico_documentos import montar_contexto_de_servico
from servicos.atesto_assinado import preparar_pdf_atesto_assinado

_SSML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


class ImpressaoError(Exception):
    pass


def _ocultar_aba_ir(conteudo: bytes) -> bytes:
    """A aba IR-2024 é tabela de apoio; a impressão fica só com o formulário RPA.

    A edição fica no XML para não descartar desenhos nem a configuração de página.
    """
    saida = BytesIO()
    with zipfile.ZipFile(BytesIO(conteudo)) as origem, zipfile.ZipFile(
        saida, "w", compression=zipfile.ZIP_DEFLATED
    ) as destino:
        for info in origem.infolist():
            dados = origem.read(info.filename)
            if info.filename == "xl/workbook.xml":
                raiz = etree.fromstring(dados)
                for planilha in raiz.findall(f"{{{_SSML}}}sheets/{{{_SSML}}}sheet"):
                    if planilha.get("name") == "IR-2024":
                        planilha.set("state", "hidden")
                dados = etree.tostring(
                    raiz, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            destino.writestr(info, dados)
    return saida.getvalue()


def montar_pdfs_impressao(pdf_dir: Path, atesto: Path | None) -> list[Path]:
    """Ordem da impressão: previsão, CI, RPA e, se houver, o atesto assinado."""
    pdfs = [
        pdf_dir / "01-previsao.pdf",
        pdf_dir / "02-ci.pdf",
        pdf_dir / "03-rpa.pdf",
    ]
    if atesto is not None:
        pdfs.append(atesto)
    return pdfs


def pdf_impressao(documento) -> bytes:
    """
    Gera o PDF da previsão de pagamento, da CI, do RPA e do atesto assinado.

    O atesto assinado só entra quando estiver anexado ao serviço.
    O arquivo permanece em PDF, sem conversão para imagem.
    A área de impressão, a escala, o ajuste à página e as margens de cada
    modelo são os da planilha ou do documento.
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
        ci_path.write_bytes(documento.arquivo_pdf.read())
        previsao_path.write_bytes(previsao.read())
        rpa_path.write_bytes(_ocultar_aba_ir(rpa.read()))

        pdf_dir = base / "pdf"
        pdf_dir.mkdir()
        atesto_pdf = None
        if servico.atesto_assinado:
            atesto_pdf = pdf_dir / "04-atesto.pdf"
            try:
                preparar_pdf_atesto_assinado(servico, atesto_pdf)
            except ValidationError as exc:
                raise ImpressaoError(" ".join(exc.messages)) from exc
        pares = [
            (previsao_path, pdf_dir / "01-previsao.pdf"),
            (ci_path, pdf_dir / "02-ci.pdf"),
            (rpa_path, pdf_dir / "03-rpa.pdf"),
        ]
        try:
            converter_para_pdf(pares)
        except ConversaoPdfError as exc:
            raise ImpressaoError(str(exc)) from exc

        pdfs = montar_pdfs_impressao(pdf_dir, atesto_pdf)
        ausentes = [caminho.name for caminho in pdfs if not caminho.exists()]
        if ausentes:
            raise ImpressaoError(
                "Não foi possível converter todos os documentos: "
                + ", ".join(ausentes)
            )

        unido = base / "documentos.pdf"
        try:
            unir_pdfs(pdfs, unido)
        except ConversaoPdfError as exc:
            raise ImpressaoError(str(exc)) from exc
        if not unido.exists():
            raise ImpressaoError("Nenhuma página foi gerada para impressão.")
        return unido.read_bytes()
