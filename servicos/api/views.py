from datetime import datetime

from rest_framework import status, viewsets
from rest_framework.response import Response

from core.services.servico_documentos import (
    ServicoDocumentoError,
    criar_servico_com_documentos,
    editar_servico_com_documentos,
)
from empresas.models import Empresa
from pessoas.models import Pessoa
from servicos.api.serializers import ServicoCreateSerializer, ServicoSerializer
from servicos.models import Servico
from textos_padroes.models import TextoPadrao


def _parse_dias(raw):
    dias = []
    for item in raw:
        data = item["data"]
        if isinstance(data, str):
            data = datetime.strptime(data, "%Y-%m-%d").date()
        dias.append(
            {
                "data": data,
                "hora_inicio": item["hora_inicio"],
                "hora_fim": item["hora_fim"],
                "observacao": item.get("observacao") or "",
            }
        )
    return dias


def _textos_selecionados(data: dict) -> dict:
    textos = {
        "DOCUMENTO_1": TextoPadrao.objects.get(pk=data["texto_documento_1"]),
    }
    if data.get("texto_documento_2"):
        textos["DOCUMENTO_2"] = TextoPadrao.objects.get(pk=data["texto_documento_2"])
    if data.get("texto_documento_3"):
        textos["DOCUMENTO_3"] = TextoPadrao.objects.get(pk=data["texto_documento_3"])
    return textos


class ServicoViewSet(viewsets.ModelViewSet):
    queryset = Servico.objects.select_related("pessoa", "empresa").prefetch_related(
        "dias_trabalhados", "documentos"
    )
    serializer_class = ServicoSerializer
    search_fields = ["numero_documento", "pessoa__nome", "empresa__nome"]
    filterset_fields = ["pessoa", "empresa", "ano"]
    http_method_names = ["get", "post", "put", "patch", "head", "options"]

    def create(self, request, *args, **kwargs):
        serializer = ServicoCreateSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        data = serializer.validated_data
        try:
            servico = criar_servico_com_documentos(
                pessoa=Pessoa.objects.get(pk=data["pessoa"]),
                empresa=Empresa.objects.get(pk=data["empresa"]),
                motivo_contratacao=data["motivo_contratacao"],
                funcao=data["funcao"],
                data_servico=data["data_servico"],
                assinatura=data.get("assinatura") or Servico.Assinatura.WALQUIRIA,
                dias_trabalhados=_parse_dias(data["dias_trabalhados"]),
                textos_por_tipo=_textos_selecionados(data),
                usuario=request.user,
            )
        except (
            ServicoDocumentoError,
            Pessoa.DoesNotExist,
            Empresa.DoesNotExist,
            TextoPadrao.DoesNotExist,
        ) as exc:
            return Response({"detail": str(exc)}, status=status.HTTP_400_BAD_REQUEST)
        return Response(ServicoSerializer(servico).data, status=status.HTTP_201_CREATED)

    def update(self, request, *args, **kwargs):
        servico = self.get_object()
        serializer = ServicoCreateSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        data = serializer.validated_data
        try:
            servico = editar_servico_com_documentos(
                servico=servico,
                motivo_contratacao=data["motivo_contratacao"],
                funcao=data["funcao"],
                data_servico=data["data_servico"],
                assinatura=data.get("assinatura") or servico.assinatura,
                dias_trabalhados=_parse_dias(data["dias_trabalhados"]),
                textos_por_tipo=_textos_selecionados(data),
                usuario=request.user,
                motivo_alteracao=data.get("motivo_alteracao") or "Edição via API",
            )
        except (ServicoDocumentoError, TextoPadrao.DoesNotExist) as exc:
            return Response({"detail": str(exc)}, status=status.HTTP_400_BAD_REQUEST)
        return Response(ServicoSerializer(servico).data)
