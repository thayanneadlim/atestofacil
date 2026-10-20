from rest_framework import serializers

from pessoas.models import Pessoa


class PessoaSerializer(serializers.ModelSerializer):
    cpf_formatado = serializers.CharField(read_only=True)

    class Meta:
        model = Pessoa
        fields = [
            "id",
            "nome",
            "cpf",
            "cpf_formatado",
            "chave_pix",
            "ativo",
            "created_at",
            "updated_at",
        ]
        read_only_fields = ["created_at", "updated_at", "cpf_formatado"]


class PessoaListSerializer(serializers.ModelSerializer):
    cpf_mascarado = serializers.SerializerMethodField()

    class Meta:
        model = Pessoa
        fields = ["id", "nome", "cpf_mascarado", "ativo", "created_at"]

    def get_cpf_mascarado(self, obj):
        cpf = obj.cpf or ""
        if len(cpf) != 11:
            return "***"
        return f"***.{cpf[3:6]}.***-{cpf[9:]}"
