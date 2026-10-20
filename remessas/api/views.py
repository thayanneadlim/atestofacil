from rest_framework import status, viewsets
from rest_framework.decorators import action
from rest_framework.response import Response

from core.services.remessa import RemessaError, criar_remessa_completa, editar_remessa
from empresas.models import Empresa
from pessoas.models import Pessoa
from remessas.api.serializers import RemessaCreateSerializer, RemessaSerializer
from remessas.models import Remessa
from textos_padroes.models import TextoPadrao


def _parse_dias(raw):
    from datetime import datetime
    from decimal import Decimal

    dias = []
    for item in raw:
        data = item["data"]
        if isinstance(data, str):
            data = datetime.strptime(data, "%Y-%m-%d").date()
        dias.append(
            {
                "data": data,
                "quantidade_horas": Decimal(str(item["quantidade_horas"])),
                "observacao": item.get("observacao") or "",
            }
        )
    return dias


class RemessaViewSet(viewsets.ModelViewSet):
    queryset = Remessa.objects.select_related("pessoa", "servico__empresa")
    serializer_class = RemessaSerializer
    search_fields = ["numero_formatado", "pessoa__nome", "pessoa__cpf"]
    filterset_fields = ["pessoa", "status", "ano"]
    http_method_names = ["get", "post", "head", "options"]

    def create(self, request, *args, **kwargs):
        serializer = RemessaCreateSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        data = serializer.validated_data
        try:
            remessa = criar_remessa_completa(
                pessoa=Pessoa.objects.get(pk=data["pessoa"]),
                empresa=Empresa.objects.get(pk=data["empresa"]),
                motivo_contratacao=data["motivo_contratacao"],
                funcao=data["funcao"],
                data_servico=data["data_servico"],
                dias_trabalhados=_parse_dias(data["dias_trabalhados"]),
                textos_por_tipo={
                    "DOCUMENTO_1": TextoPadrao.objects.get(pk=data["texto_documento_1"]),
                    "DOCUMENTO_2": TextoPadrao.objects.get(pk=data["texto_documento_2"]),
                    "DOCUMENTO_3": TextoPadrao.objects.get(pk=data["texto_documento_3"]),
                },
                usuario=request.user,
            )
        except (
            RemessaError,
            Pessoa.DoesNotExist,
            Empresa.DoesNotExist,
            TextoPadrao.DoesNotExist,
        ) as exc:
            return Response({"detail": str(exc)}, status=status.HTTP_400_BAD_REQUEST)
        return Response(RemessaSerializer(remessa).data, status=status.HTTP_201_CREATED)

    @action(detail=True, methods=["post"])
    def editar(self, request, pk=None):
        remessa = self.get_object()
        serializer = RemessaCreateSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        data = serializer.validated_data
        try:
            remessa = editar_remessa(
                remessa=remessa,
                motivo_contratacao=data["motivo_contratacao"],
                funcao=data["funcao"],
                data_servico=data["data_servico"],
                dias_trabalhados=_parse_dias(data["dias_trabalhados"]),
                textos_por_tipo={
                    "DOCUMENTO_1": TextoPadrao.objects.get(pk=data["texto_documento_1"]),
                    "DOCUMENTO_2": TextoPadrao.objects.get(pk=data["texto_documento_2"]),
                    "DOCUMENTO_3": TextoPadrao.objects.get(pk=data["texto_documento_3"]),
                },
                usuario=request.user,
                motivo_alteracao=data.get("motivo_alteracao") or "Edição via API",
            )
        except (RemessaError, TextoPadrao.DoesNotExist) as exc:
            return Response({"detail": str(exc)}, status=status.HTTP_400_BAD_REQUEST)
        return Response(RemessaSerializer(remessa).data)
