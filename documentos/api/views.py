from django.http import FileResponse, Http404
from rest_framework import viewsets
from rest_framework.decorators import action

from core.services.nomes_arquivo import nome_arquivo_download
from documentos.api.serializers import DocumentoListSerializer, DocumentoSerializer
from documentos.models import Documento
from documentos.views import _content_type_arquivo


class DocumentoViewSet(viewsets.ReadOnlyModelViewSet):
    queryset = Documento.objects.select_related(
        "pessoa", "servico__empresa", "remessa"
    ).prefetch_related("versoes")
    search_fields = ["numero_documento", "pessoa__nome", "pessoa__cpf"]
    filterset_fields = ["tipo_documento", "pessoa", "remessa"]

    def get_serializer_class(self):
        if self.action == "list":
            return DocumentoListSerializer
        return DocumentoSerializer

    @action(detail=True, methods=["get"])
    def pdf(self, request, pk=None):
        documento = self.get_object()
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
