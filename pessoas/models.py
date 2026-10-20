from django.db import models
from django.core.exceptions import ValidationError

from core.services.templates_texto import normalizar_cpf, validar_cpf


class Pessoa(models.Model):
    nome = models.CharField(max_length=255)
    cpf = models.CharField(max_length=11, db_index=True)
    chave_pix = models.CharField(max_length=255)
    ativo = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Pessoa"
        verbose_name_plural = "Pessoas"
        ordering = ["nome"]
        constraints = [
            models.UniqueConstraint(
                fields=["cpf"],
                condition=models.Q(ativo=True),
                name="uniq_pessoa_cpf_ativo",
            )
        ]

    def clean(self):
        self.cpf = normalizar_cpf(self.cpf)
        if not self.nome:
            raise ValidationError({"nome": "Nome é obrigatório."})
        if not self.cpf:
            raise ValidationError({"cpf": "CPF é obrigatório."})
        if not validar_cpf(self.cpf):
            raise ValidationError({"cpf": "CPF inválido."})
        if not self.chave_pix:
            raise ValidationError({"chave_pix": "Chave Pix é obrigatória."})

    def save(self, *args, **kwargs):
        self.cpf = normalizar_cpf(self.cpf)
        self.full_clean()
        return super().save(*args, **kwargs)

    def __str__(self) -> str:
        return self.nome

    @property
    def cpf_formatado(self) -> str:
        from core.services.templates_texto import formatar_cpf

        return formatar_cpf(self.cpf)
