from rest_framework import serializers

from textos_padroes.models import TextoPadrao


class TextoPadraoSerializer(serializers.ModelSerializer):
    class Meta:
        model = TextoPadrao
        fields = [
            "id",
            "nome",
            "descricao",
            "conteudo",
            "tipo_documento",
            "ativo",
            "created_at",
            "updated_at",
        ]
        read_only_fields = ["created_at", "updated_at"]
