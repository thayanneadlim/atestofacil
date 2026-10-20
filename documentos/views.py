import base64

from django.contrib.auth.decorators import login_required
from django.db.models import Q
from django.http import FileResponse, Http404
from django.shortcuts import get_object_or_404, render

from core.services.impressao import ImpressaoError, paginas_png_documentos
from core.services.nomes_arquivo import nome_arquivo_download
from documentos.forms import DocumentoFilterForm
from documentos.models import Documento, DocumentoVersao


def _content_type_arquivo(nome: str) -> str:
    lower = (nome or "").lower()
    if lower.endswith(".docx"):
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if lower.endswith(".doc"):
        return "application/msword"
    if lower.endswith(".pdf"):
        return "application/pdf"
    return "application/octet-stream"


@login_required
def documento_list(request):
    form = DocumentoFilterForm(request.GET or None)
    qs = Documento.objects.select_related("pessoa", "servico__empresa", "remessa")
    if form.is_valid():
        q = form.cleaned_data.get("q") or ""
        if q:
            cpf = q.replace(".", "").replace("-", "")
            qs = qs.filter(
                Q(numero_documento__icontains=q)
                | Q(pessoa__nome__icontains=q)
                | Q(pessoa__cpf__icontains=cpf)
            )
        empresa = form.cleaned_data.get("empresa")
        if empresa:
            qs = qs.filter(servico__empresa=empresa)
        if form.cleaned_data.get("data_inicio"):
            qs = qs.filter(data_documento__gte=form.cleaned_data["data_inicio"])
        if form.cleaned_data.get("data_fim"):
            qs = qs.filter(data_documento__lte=form.cleaned_data["data_fim"])
    return render(request, "documentos/list.html", {"documentos": qs, "form": form})


@login_required
def documento_detail(request, pk):
    documento = get_object_or_404(
        Documento.objects.select_related(
            "pessoa", "servico__empresa", "remessa", "texto_padrao"
        ),
        pk=pk,
    )
    versoes = documento.versoes.all()
    return render(
        request,
        "documentos/detail.html",
        {"documento": documento, "versoes": versoes},
    )


@login_required
def documento_pdf(request, pk):
    """Mantém a URL histórica /pdf/, mas agora serve o arquivo gerado (.docx)."""
    documento = get_object_or_404(Documento, pk=pk)
    if not documento.arquivo_pdf:
        raise Http404("Arquivo não encontrado.")
    armazenado = documento.arquivo_pdf.name
    nome = nome_arquivo_download(
        "CI",
        documento.numero_documento,
        documento.pessoa.nome,
        "docx",
    )
    return FileResponse(
        documento.arquivo_pdf.open("rb"),
        content_type=_content_type_arquivo(armazenado),
        as_attachment=True,
        filename=nome,
    )


@login_required
def documento_versao(request, pk, versao):
    versao_obj = get_object_or_404(DocumentoVersao, documento_id=pk, versao=versao)
    return render(request, "documentos/versao.html", {"versao": versao_obj})


@login_required
def documento_print(request, pk):
    documento = get_object_or_404(
        Documento.objects.select_related("pessoa", "servico__empresa").prefetch_related(
            "servico__dias_trabalhados",
            "servico__documentos",
        ),
        pk=pk,
    )
    try:
        paginas = paginas_png_documentos(documento)
    except (ImpressaoError, FileNotFoundError) as exc:
        return render(
            request,
            "documentos/print.html",
            {"documento": documento, "erro": str(exc)},
            status=500,
        )
    return render(
        request,
        "documentos/print.html",
        {
            "documento": documento,
            "paginas": [base64.b64encode(pagina).decode("ascii") for pagina in paginas],
        },
    )
