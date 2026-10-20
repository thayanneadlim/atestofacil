from datetime import datetime

from rest_framework import serializers

from servicos.models import DiaTrabalhado, Servico


class DiaTrabalhadoSerializer(serializers.ModelSerializer):
    class Meta:
        model = DiaTrabalhado
        fields = [
            "id",
            "data",
            "hora_inicio",
            "hora_fim",
            "quantidade_horas",
            "observacao",
            "created_at",
        ]
        read_only_fields = ["id", "quantidade_horas", "created_at"]


class ServicoSerializer(serializers.ModelSerializer):
    dias_trabalhados = DiaTrabalhadoSerializer(many=True, read_only=True)
    total_horas = serializers.DecimalField(
        max_digits=8, decimal_places=2, read_only=True
    )

    class Meta:
        model = Servico
        fields = [
            "id",
            "numero",
            "ano",
            "numero_documento",
            "numero_rpa",
            "ano_rpa",
            "numero_rpa_documento",
            "empresa",
            "pessoa",
            "data_servico",
            "motivo_contratacao",
            "funcao",
            "assinatura",
            "dias_trabalhados",
            "total_horas",
            "created_at",
            "updated_at",
        ]
        read_only_fields = [
            "numero",
            "ano",
            "numero_documento",
            "numero_rpa",
            "ano_rpa",
            "numero_rpa_documento",
            "total_horas",
            "created_at",
            "updated_at",
        ]


class ServicoCreateSerializer(serializers.Serializer):
    pessoa = serializers.IntegerField()
    empresa = serializers.IntegerField()
    data_servico = serializers.DateField()
    motivo_contratacao = serializers.CharField()
    funcao = serializers.CharField(max_length=255)
    assinatura = serializers.ChoiceField(
        choices=Servico.Assinatura.choices,
        required=False,
    )
    texto_documento_1 = serializers.IntegerField()
    texto_documento_2 = serializers.IntegerField(required=False, allow_null=True)
    texto_documento_3 = serializers.IntegerField(required=False, allow_null=True)
    dias_trabalhados = serializers.ListField(child=serializers.DictField())
    motivo_alteracao = serializers.CharField(required=False, allow_blank=True)
