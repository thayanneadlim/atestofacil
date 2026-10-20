import base64
from datetime import date
from io import BytesIO
from pathlib import Path

import pymupdf
from django.contrib.auth import get_user_model
from django.core.files.base import ContentFile
from django.test import TestCase
from django.urls import reverse
from PIL import Image

from core.services.conversao_pdf import converter_para_pdf, localizar_soffice
from core.services.impressao import ImpressaoError, paginas_png_documentos
from core.services.servico_documentos import criar_servico_com_documentos
from core.tests import _cpf_valido
from empresas.models import Empresa
from pessoas.models import Pessoa
from textos_padroes.models import TextoPadrao

MODELOS = Path(__file__).resolve().parents[1] / "media" / "documentos_manuais"


def _png_minimo() -> bytes:
    imagem = Image.new("RGB", (40, 60), "white")
    buffer = BytesIO()
    imagem.save(buffer, format="PNG")
    return buffer.getvalue()


def _texto_pdf(caminho: Path) -> str:
    documento = pymupdf.open(caminho)
    try:
        texto = "\n".join(pagina.get_text() for pagina in documento)
    finally:
        documento.close()
    return " ".join(texto.replace("\xa0", " ").replace("\xad", "").split())


class ConversaoSemLibreOfficeTests(TestCase):
    def test_modelos_viram_pdf_sem_soffice(self):
        self.assertIsNone(localizar_soffice())
        from tempfile import TemporaryDirectory

        with TemporaryDirectory() as pasta:
            base = Path(pasta)
            pares = [
                (MODELOS / "modelo_previsaodepagamento.xlsx", base / "previsao.pdf"),
                (MODELOS / "ci_padrao.docx", base / "ci.pdf"),
                (MODELOS / "modelo_RPA.xlsx", base / "rpa.pdf"),
            ]
            converter_para_pdf(pares)
            previsao = _texto_pdf(base / "previsao.pdf")
            ci = _texto_pdf(base / "ci.pdf")
            rpa = _texto_pdf(base / "rpa.pdf")
            ci_pdf = pymupdf.open(base / "ci.pdf")
            try:
                self.assertEqual(ci_pdf.page_count, 1)
                self.assertTrue(ci_pdf[0].get_images())
                verdes = [
                    item
                    for item in ci_pdf[0].get_drawings()
                    if item.get("color")
                    and item["color"][1] > 0.35
                    and item["color"][0] < 0.2
                ]
                self.assertTrue(verdes)
            finally:
                ci_pdf.close()
        self.assertIn("Previsão de Pagamento", previsao)
        self.assertIn("Comunicação Interna", ci)
        self.assertIn("Atenciosamente", ci)
        self.assertIn("NOSSO CRESCIMENTO", ci)
        self.assertIn("RECIBO DE PAGAMENTO", rpa)
        self.assertIn("1,33", rpa)
        self.assertNotIn("TRANSPORTAR RPA", rpa)
        self.assertNotIn("VALOR MIN", rpa)


class ImpressaoViewTests(TestCase):
    def setUp(self):
        self.pessoa = Pessoa.objects.create(
            nome="Maria",
            cpf=_cpf_valido(),
            chave_pix="maria@email.com",
        )
        self.empresa = Empresa.objects.create(nome="Empresa B")
        self.texto = TextoPadrao.objects.create(
            nome="CI",
            conteudo="Doc {{ nome }}",
            tipo_documento="DOCUMENTO_1",
        )
        self.user = get_user_model().objects.create_user("admin", password="admin123")
        self.client.login(username="admin", password="admin123")
        self.servico = criar_servico_com_documentos(
            pessoa=self.pessoa,
            empresa=self.empresa,
            motivo_contratacao="Cobertura",
            funcao="Assistente",
            data_servico=date(2026, 9, 1),
            dias_trabalhados=[
                {"data": date(2026, 9, 1), "hora_inicio": "08:00", "hora_fim": "16:00"}
            ],
            textos_por_tipo={"DOCUMENTO_1": self.texto},
            usuario=self.user,
        )
        self.documento = self.servico.documentos.get()

    def test_imprimir_abre_dialogo_com_png(self):
        resposta = self.client.get(reverse("documentos:print", args=[self.documento.pk]))
        self.assertEqual(resposta.status_code, 200)
        self.assertContains(resposta, 'class="barra no-print"')
        self.assertContains(resposta, 'onload="window.print()"')
        self.assertContains(resposta, "data:image/png;base64,")
        self.assertContains(resposta, "page-break-after: always")
        png = base64.b64decode(resposta.context["paginas"][0])
        self.assertTrue(png.startswith(b"\x89PNG"))
        self.assertGreaterEqual(len(resposta.context["paginas"]), 3)

    def test_sem_ci_nao_chama_impressao(self):
        self.documento.arquivo_pdf.delete(save=True)
        resposta = self.client.get(reverse("documentos:print", args=[self.documento.pk]))
        self.assertEqual(resposta.status_code, 500)
        self.assertNotContains(resposta, "window.print()", status_code=500)
        self.assertContains(resposta, "ainda não foi gerada", status_code=500)

    def test_atesto_assinado_entra_no_pacote(self):
        antes = len(paginas_png_documentos(self.documento))
        self.servico.atesto_assinado.save("atesto.png", ContentFile(_png_minimo()), save=True)
        self.documento.refresh_from_db()
        depois = paginas_png_documentos(self.documento)
        self.assertGreater(len(depois), antes)
        self.assertTrue(depois[-1].startswith(b"\x89PNG"))

    def test_impressao_reaproveita_ci_armazenada(self):
        from tempfile import TemporaryDirectory

        from core.services.conversao_pdf import _pdf_python

        with self.documento.arquivo_pdf.open("rb") as arquivo:
            dados = arquivo.read()
        with TemporaryDirectory() as pasta:
            origem = Path(pasta) / "ci.docx"
            origem.write_bytes(dados)
            pdf = Path(pasta) / "ci.pdf"
            pdf.write_bytes(_pdf_python(origem))
            texto = _texto_pdf(pdf)
            documento = pymupdf.open(pdf)
            try:
                self.assertEqual(documento.page_count, 1)
                self.assertTrue(documento[0].get_images())
            finally:
                documento.close()
        self.assertIn("Maria", texto)
        self.assertIn("NOSSO CRESCIMENTO", texto)
        self.assertNotIn("{{", texto)

    def test_sem_arquivo_levanta_erro(self):
        self.documento.arquivo_pdf = None
        with self.assertRaises(ImpressaoError):
            paginas_png_documentos(self.documento)
