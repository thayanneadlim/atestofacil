from rest_framework import serializers

from empresas.models import Empresa


class EmpresaSerializer(serializers.ModelSerializer):
    class Meta:
        model = Empresa
        fields = ["id", "nome", "cnpj", "ativo", "created_at", "updated_at"]
        read_only_fields = ["created_at", "updated_at"]
