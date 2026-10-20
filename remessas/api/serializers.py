from rest_framework import serializers

from remessas.models import Remessa


class RemessaSerializer(serializers.ModelSerializer):
    class Meta:
        model = Remessa
        fields = [
            "id",
            "numero",
            "ano",
            "numero_formatado",
            "pessoa",
            "servico",
            "data_criacao",
            "status",
            "created_at",
            "updated_at",
        ]
        read_only_fields = fields


class RemessaCreateSerializer(serializers.Serializer):
    pessoa = serializers.IntegerField()
    empresa = serializers.IntegerField()
    data_servico = serializers.DateField()
    motivo_contratacao = serializers.CharField()
    funcao = serializers.CharField(max_length=255)
    texto_documento_1 = serializers.IntegerField()
    texto_documento_2 = serializers.IntegerField()
    texto_documento_3 = serializers.IntegerField()
    dias_trabalhados = serializers.ListField(child=serializers.DictField())
    motivo_alteracao = serializers.CharField(required=False, allow_blank=True)
