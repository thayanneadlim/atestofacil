from rest_framework import viewsets
from rest_framework.exceptions import ValidationError

from textos_padroes.api.serializers import TextoPadraoSerializer
from textos_padroes.models import TextoPadrao


class TextoPadraoViewSet(viewsets.ModelViewSet):
    queryset = TextoPadrao.objects.all()
    serializer_class = TextoPadraoSerializer
    search_fields = ["nome", "descricao"]
    filterset_fields = ["tipo_documento", "ativo"]

    def perform_destroy(self, instance):
        if instance.em_uso():
            raise ValidationError(
                "Texto já utilizado em documentos. Desative em vez de excluir."
            )
        instance.delete()
