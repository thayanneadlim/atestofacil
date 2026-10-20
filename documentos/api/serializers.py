from rest_framework import serializers

from documentos.models import Documento, DocumentoVersao


class DocumentoVersaoSerializer(serializers.ModelSerializer):
    class Meta:
        model = DocumentoVersao
        fields = [
            "id",
            "versao",
            "conteudo_renderizado",
            "arquivo_pdf",
            "snapshot_dados",
            "motivo_alteracao",
            "criado_por",
            "created_at",
        ]


class DocumentoSerializer(serializers.ModelSerializer):
    versoes = DocumentoVersaoSerializer(many=True, read_only=True)

    class Meta:
        model = Documento
        fields = [
            "id",
            "remessa",
            "tipo_documento",
            "numero_documento",
            "pessoa",
            "servico",
            "texto_padrao",
            "conteudo_renderizado",
            "arquivo_pdf",
            "data_documento",
            "versao_atual",
            "snapshot_dados",
            "versoes",
            "created_at",
            "updated_at",
        ]
        read_only_fields = fields


class DocumentoListSerializer(serializers.ModelSerializer):
    class Meta:
        model = Documento
        fields = [
            "id",
            "remessa",
            "tipo_documento",
            "numero_documento",
            "pessoa",
            "data_documento",
            "versao_atual",
        ]
