from django.db import models
from django.core.exceptions import ValidationError
import re


class Empresa(models.Model):
    nome = models.CharField(max_length=255)
    cnpj = models.CharField(max_length=14, blank=True, default="")
    ativo = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Empresa"
        verbose_name_plural = "Empresas"
        ordering = ["nome"]

    def clean(self):
        if not self.nome:
            raise ValidationError({"nome": "Nome é obrigatório."})
        if self.cnpj:
            digitos = re.sub(r"\D", "", self.cnpj)
            if len(digitos) not in (0, 14):
                raise ValidationError({"cnpj": "CNPJ deve conter 14 dígitos."})
            self.cnpj = digitos

    def save(self, *args, **kwargs):
        if self.cnpj:
            self.cnpj = re.sub(r"\D", "", self.cnpj)
        self.full_clean()
        return super().save(*args, **kwargs)

    def __str__(self) -> str:
        return self.nome

    def pode_excluir(self) -> bool:
        return not self.servicos.exists()
