from rest_framework import viewsets

from pessoas.api.serializers import PessoaListSerializer, PessoaSerializer
from pessoas.models import Pessoa


class PessoaViewSet(viewsets.ModelViewSet):
    queryset = Pessoa.objects.all()
    search_fields = ["nome", "cpf"]
    filterset_fields = ["ativo"]
    ordering_fields = ["nome", "created_at"]

    def get_serializer_class(self):
        if self.action == "list":
            return PessoaListSerializer
        return PessoaSerializer
