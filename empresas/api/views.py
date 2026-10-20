from rest_framework import viewsets
from rest_framework.exceptions import ValidationError

from empresas.api.serializers import EmpresaSerializer
from empresas.models import Empresa


class EmpresaViewSet(viewsets.ModelViewSet):
    queryset = Empresa.objects.all()
    serializer_class = EmpresaSerializer
    search_fields = ["nome", "cnpj"]
    filterset_fields = ["ativo"]

    def perform_destroy(self, instance):
        if not instance.pode_excluir():
            raise ValidationError(
                "Empresa vinculada a serviços. Desative em vez de excluir."
            )
        instance.delete()
