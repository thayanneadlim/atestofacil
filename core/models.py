from django.db import models


class SequenciaAnual(models.Model):
    class Tipo(models.TextChoices):
        REMESSA = "REMESSA", "Remessa"
        SERVICO = "SERVICO", "CI e previsão de pagamento"
        RPA = "RPA", "RPA"

    tipo = models.CharField(max_length=20, choices=Tipo.choices)
    ano = models.PositiveIntegerField()
    ultimo_numero = models.PositiveIntegerField(default=0)

    class Meta:
        verbose_name = "Sequência anual"
        verbose_name_plural = "Sequências anuais"
        constraints = [
            models.UniqueConstraint(
                fields=["tipo", "ano"],
                name="uniq_sequencia_tipo_ano",
            )
        ]

    def __str__(self) -> str:
        return f"{self.tipo} {self.ano}: {self.ultimo_numero}"
