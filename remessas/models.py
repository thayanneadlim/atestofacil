from django.db import models


class Remessa(models.Model):
    class Status(models.TextChoices):
        RASCUNHO = "RASCUNHO", "Rascunho"
        GERADA = "GERADA", "Gerada"
        EDITADA = "EDITADA", "Editada"
        CANCELADA = "CANCELADA", "Cancelada"

    numero = models.PositiveIntegerField()
    ano = models.PositiveIntegerField()
    numero_formatado = models.CharField(max_length=20, unique=True)
    pessoa = models.ForeignKey(
        "pessoas.Pessoa",
        on_delete=models.PROTECT,
        related_name="remessas",
    )
    servico = models.OneToOneField(
        "servicos.Servico",
        on_delete=models.PROTECT,
        related_name="remessa_principal",
    )
    data_criacao = models.DateField()
    status = models.CharField(
        max_length=20,
        choices=Status.choices,
        default=Status.RASCUNHO,
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Remessa"
        verbose_name_plural = "Remessas"
        ordering = ["-ano", "-numero"]
        constraints = [
            models.UniqueConstraint(
                fields=["numero", "ano"],
                name="uniq_remessa_numero_ano",
            )
        ]

    def __str__(self) -> str:
        return self.numero_formatado
