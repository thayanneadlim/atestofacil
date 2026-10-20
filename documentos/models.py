from django.conf import settings
from django.db import models


class Documento(models.Model):
    remessa = models.ForeignKey(
        "remessas.Remessa",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="documentos",
    )
    tipo_documento = models.CharField(max_length=20, choices=settings.DOCUMENTO_TIPOS)
    numero_documento = models.CharField(max_length=20, db_index=True)
    pessoa = models.ForeignKey(
        "pessoas.Pessoa",
        on_delete=models.PROTECT,
        related_name="documentos",
    )
    servico = models.ForeignKey(
        "servicos.Servico",
        on_delete=models.CASCADE,
        related_name="documentos",
    )
    texto_padrao = models.ForeignKey(
        "textos_padroes.TextoPadrao",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="documentos",
    )
    conteudo_renderizado = models.TextField()
    arquivo_pdf = models.FileField(upload_to="documentos/", blank=True)
    data_documento = models.DateField()
    versao_atual = models.PositiveIntegerField(default=1)
    snapshot_dados = models.JSONField(default=dict, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Documento"
        verbose_name_plural = "Documentos"
        ordering = ["tipo_documento"]
        constraints = [
            models.UniqueConstraint(
                fields=["servico", "tipo_documento"],
                name="uniq_documento_tipo_por_servico",
            )
        ]

    def __str__(self) -> str:
        return f"{self.numero_documento} - {self.tipo_documento}"


class DocumentoVersao(models.Model):
    documento = models.ForeignKey(
        Documento,
        on_delete=models.CASCADE,
        related_name="versoes",
    )
    versao = models.PositiveIntegerField()
    conteudo_renderizado = models.TextField()
    arquivo_pdf = models.CharField(max_length=500, blank=True, default="")
    snapshot_dados = models.JSONField(default=dict, blank=True)
    motivo_alteracao = models.CharField(max_length=255, blank=True, default="")
    criado_por = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Versão de documento"
        verbose_name_plural = "Versões de documentos"
        ordering = ["-versao"]
        constraints = [
            models.UniqueConstraint(
                fields=["documento", "versao"],
                name="uniq_documento_versao",
            )
        ]

    def __str__(self) -> str:
        return f"{self.documento_id} v{self.versao}"
