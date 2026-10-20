from io import BytesIO

from django.conf import settings
from django.core.files.base import ContentFile
from django.template.loader import render_to_string


def gerar_pdf_documento(*, documento, titulo: str) -> ContentFile:
    """Gera PDF a partir de template HTML. Fallback textual se WeasyPrint falhar."""
    html = render_to_string(
        "documentos/pdf.html",
        {
            "documento": documento,
            "titulo": titulo,
            "conteudo": documento.conteudo_renderizado,
        },
    )

    try:
        from weasyprint import HTML

        buffer = BytesIO()
        HTML(string=html, base_url=str(settings.BASE_DIR)).write_pdf(buffer)
        pdf_bytes = buffer.getvalue()
    except Exception:
        # Ambiente sem libs do sistema do WeasyPrint: gera PDF mínimo válido.
        pdf_bytes = _pdf_minimo(documento.conteudo_renderizado or titulo)

    nome = f"{documento.numero_documento.replace('/', '-')}_{documento.tipo_documento}_v{documento.versao_atual}.pdf"
    return ContentFile(pdf_bytes, name=nome)


def _pdf_minimo(texto: str) -> bytes:
    """PDF simples para ambientes sem WeasyPrint/system deps."""
    safe = (texto or "Documento").encode("latin-1", errors="replace")[:1800]
    # Conteúdo stream com texto básico
    stream = b"BT /F1 10 Tf 50 750 Td (" + safe.replace(b"(", b"[").replace(b")", b"]") + b") Tj ET"
    objects = []
    objects.append(b"1 0 obj<< /Type /Catalog /Pages 2 0 R >>endobj\n")
    objects.append(b"2 0 obj<< /Type /Pages /Kids [3 0 R] /Count 1 >>endobj\n")
    objects.append(
        b"3 0 obj<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] "
        b"/Contents 4 0 R /Resources<< /Font<< /F1 5 0 R >> >> >>endobj\n"
    )
    objects.append(
        f"4 0 obj<< /Length {len(stream)} >>stream\n".encode() + stream + b"\nendstream endobj\n"
    )
    objects.append(b"5 0 obj<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>endobj\n")

    out = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for obj in objects:
        offsets.append(len(out))
        out.extend(obj)
    xref_pos = len(out)
    out.extend(f"xref\n0 {len(offsets)}\n".encode())
    out.extend(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        out.extend(f"{off:010d} 00000 n \n".encode())
    out.extend(
        f"trailer<< /Size {len(offsets)} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n".encode()
    )
    return bytes(out)
