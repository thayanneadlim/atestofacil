from django.conf import settings
from django.db import models


class TextoPadrao(models.Model):
    nome = models.CharField(max_length=255)
    descricao = models.TextField(blank=True, default="")
    conteudo = models.TextField()
    tipo_documento = models.CharField(
        max_length=20,
        choices=settings.DOCUMENTO_TIPOS,
    )
    ativo = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Texto padrão"
        verbose_name_plural = "Textos padrões"
        ordering = ["tipo_documento", "nome"]

    def __str__(self) -> str:
        return f"{self.nome} ({self.get_tipo_documento_display()})"

    def em_uso(self) -> bool:
        return self.documentos.exists()
