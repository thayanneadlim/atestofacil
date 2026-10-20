from dataclasses import dataclass
from datetime import datetime, time, timedelta
from decimal import Decimal

from django.core.exceptions import ValidationError
from django.core.validators import FileExtensionValidator
from django.db import models
from django.db.models import Sum

from servicos.atesto_assinado import caminho_atesto_assinado, validar_tamanho_atesto

LIMITE_HORARIO_NOTURNO = time(22, 0)
VALOR_DIA_6_A_8 = Decimal("110.00")
VALOR_DIA_6_A_8_APOS_22 = Decimal("125.00")
VALOR_DIA_4 = Decimal("55.00")
VALOR_DIA_4_APOS_22 = Decimal("62.50")


def calcular_quantidade_horas(hora_inicio: time, hora_fim: time) -> Decimal:
    if hora_inicio is None or hora_fim is None:
        raise ValidationError("Informe hora de início e hora de fim.")
    inicio = datetime.combine(datetime.today().date(), hora_inicio)
    fim = datetime.combine(datetime.today().date(), hora_fim)
    # Fim anterior ao início: o serviço atravessa a meia-noite e termina no dia seguinte.
    # A data do serviço não é alterada.
    if fim < inicio:
        fim += timedelta(days=1)
    elif fim == inicio:
        raise ValidationError(
            {"hora_fim": "A hora de fim deve ser diferente da hora de início."}
        )
    horas = Decimal(str((fim - inicio).total_seconds() / 3600)).quantize(Decimal("0.01"))
    if horas <= 0:
        raise ValidationError("O total de horas deve ser maior que zero.")
    return horas


def horario_apos_22h(hora_fim: time | None) -> bool:
    """Horário final depois das 22:00. Às 22:00 em ponto ainda é até 22:00."""
    return bool(hora_fim and hora_fim > LIMITE_HORARIO_NOTURNO)


def calcular_valor_dia(quantidade_horas, hora_fim: time | None) -> Decimal | None:
    """
    Valor de um dia, pela faixa de horas daquele dia.

    6 horas ou mais: R$ 110,00, ou R$ 125,00 se terminar depois das 22:00.
    4 ou 5 horas: R$ 55,00, ou R$ 62,50 se terminar depois das 22:00.
    Outras quantidades não entram na tabela.
    """
    if quantidade_horas is None:
        return None
    horas = Decimal(str(quantidade_horas)).quantize(Decimal("0.01"))
    noturno = horario_apos_22h(hora_fim)
    if horas >= Decimal("6.00"):
        return VALOR_DIA_6_A_8_APOS_22 if noturno else VALOR_DIA_6_A_8
    if horas in (Decimal("4.00"), Decimal("5.00")):
        return VALOR_DIA_4_APOS_22 if noturno else VALOR_DIA_4
    return None


def formatar_reais(valor: Decimal) -> str:
    quantized = Decimal(valor).quantize(Decimal("0.01"))
    sinal = "-" if quantized < 0 else ""
    inteiro, centavos = f"{abs(quantized):.2f}".split(".")
    grupos = []
    while inteiro:
        grupos.append(inteiro[-3:])
        inteiro = inteiro[:-3]
    return f"{sinal}R$ {'.'.join(reversed(grupos))},{centavos}"


def formatar_horas(quantidade_horas) -> str:
    horas = Decimal(str(quantidade_horas or 0)).quantize(Decimal("0.01"))
    if horas == horas.to_integral_value():
        return f"{int(horas)}h"
    texto = f"{horas:.2f}".rstrip("0").replace(".", ",")
    return f"{texto}h"


@dataclass
class ItemValorDia:
    data: object
    hora_inicio: time | None
    hora_fim: time | None
    quantidade_horas: Decimal
    valor: Decimal | None
    faixa: str
    apos_22: bool

    @property
    def valor_formatado(self) -> str:
        if self.valor is None:
            return "—"
        return formatar_reais(self.valor)

    @property
    def horas_formatadas(self) -> str:
        return formatar_horas(self.quantidade_horas)


@dataclass
class ResumoPagamento:
    total_dias: int
    dias_6_a_8: int
    dias_4: int
    dias_ate_22: int
    dias_apos_22: int
    dias_fora_faixa: int
    valor_total: Decimal
    itens: list

    @property
    def valor_total_formatado(self) -> str:
        return formatar_reais(self.valor_total)


def montar_resumo_pagamento(dias) -> ResumoPagamento:
    """Soma o pagamento dia a dia. Não usa o total de horas do serviço."""
    itens = []
    dias_6_a_8 = 0
    dias_4 = 0
    dias_ate_22 = 0
    dias_apos_22 = 0
    dias_fora_faixa = 0
    valor_total = Decimal("0.00")

    for dia in dias:
        horas = Decimal(str(dia.quantidade_horas or 0)).quantize(Decimal("0.01"))
        apos_22 = horario_apos_22h(dia.hora_fim)
        valor = calcular_valor_dia(horas, dia.hora_fim)
        if horas >= Decimal("6.00"):
            faixa = "6_8"
            dias_6_a_8 += 1
        elif horas in (Decimal("4.00"), Decimal("5.00")):
            faixa = "4"
            dias_4 += 1
        else:
            faixa = "outra"
            dias_fora_faixa += 1
        if apos_22:
            dias_apos_22 += 1
        else:
            dias_ate_22 += 1
        if valor is not None:
            valor_total += valor
        itens.append(
            ItemValorDia(
                data=dia.data,
                hora_inicio=getattr(dia, "hora_inicio", None),
                hora_fim=dia.hora_fim,
                quantidade_horas=horas,
                valor=valor,
                faixa=faixa,
                apos_22=apos_22,
            )
        )

    return ResumoPagamento(
        total_dias=len(itens),
        dias_6_a_8=dias_6_a_8,
        dias_4=dias_4,
        dias_ate_22=dias_ate_22,
        dias_apos_22=dias_apos_22,
        dias_fora_faixa=dias_fora_faixa,
        valor_total=valor_total.quantize(Decimal("0.01")),
        itens=itens,
    )


class Servico(models.Model):
    class Assinatura(models.TextChoices):
        ASAPHE = "ASAPHE", "Asaphe"
        WALQUIRIA = "WALQUIRIA", "Walquiria"

    numero = models.PositiveIntegerField()
    ano = models.PositiveIntegerField()
    numero_documento = models.CharField(max_length=20, unique=True)
    numero_rpa = models.PositiveIntegerField()
    ano_rpa = models.PositiveIntegerField()
    numero_rpa_documento = models.CharField(max_length=20, unique=True)
    empresa = models.ForeignKey(
        "empresas.Empresa",
        on_delete=models.PROTECT,
        related_name="servicos",
    )
    pessoa = models.ForeignKey(
        "pessoas.Pessoa",
        on_delete=models.PROTECT,
        related_name="servicos",
    )
    remessa = models.ForeignKey(
        "remessas.Remessa",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="servicos_vinculados",
    )
    data_servico = models.DateField()
    motivo_contratacao = models.TextField()
    funcao = models.CharField(max_length=255)
    assinatura = models.CharField(
        max_length=20,
        choices=Assinatura.choices,
        default=Assinatura.WALQUIRIA,
        verbose_name="Assinatura",
    )
    atesto_assinado = models.FileField(
        blank=True,
        upload_to=caminho_atesto_assinado,
        validators=[
            FileExtensionValidator(
                allowed_extensions=["pdf", "png", "jpg", "jpeg"],
                message="Anexe um PDF ou uma imagem (PNG ou JPG).",
            ),
            validar_tamanho_atesto,
        ],
        verbose_name="Atesto assinado",
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Serviço"
        verbose_name_plural = "Serviços"
        ordering = ["-ano", "-numero"]
        constraints = [
            models.UniqueConstraint(
                fields=["numero", "ano"],
                name="uniq_servico_numero_ano",
            ),
            models.UniqueConstraint(
                fields=["numero_rpa", "ano_rpa"],
                name="uniq_servico_numero_rpa_ano",
            ),
        ]

    def __str__(self) -> str:
        return self.numero_documento

    def dias_uteis_vencimento(self) -> int:
        """Walquiria mantém 1 dia útil. Asaphe usa 2 dias úteis."""
        if self.assinatura == self.Assinatura.ASAPHE:
            return 2
        return 1

    @property
    def total_horas(self):
        total = self.dias_trabalhados.aggregate(s=Sum("quantidade_horas"))["s"]
        return total or 0

    @property
    def resumo_pagamento(self) -> ResumoPagamento:
        return montar_resumo_pagamento(self.dias_trabalhados.all())


class DiaTrabalhado(models.Model):
    servico = models.ForeignKey(
        Servico,
        on_delete=models.CASCADE,
        related_name="dias_trabalhados",
    )
    data = models.DateField()
    hora_inicio = models.TimeField()
    hora_fim = models.TimeField()
    quantidade_horas = models.DecimalField(max_digits=5, decimal_places=2)
    observacao = models.CharField(max_length=255, blank=True, default="")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Dia trabalhado"
        verbose_name_plural = "Dias trabalhados"
        ordering = ["data"]
        constraints = [
            models.UniqueConstraint(
                fields=["servico", "data"],
                name="uniq_dia_por_servico",
            )
        ]

    def clean(self):
        if self.hora_inicio and self.hora_fim:
            self.quantidade_horas = calcular_quantidade_horas(self.hora_inicio, self.hora_fim)

    def save(self, *args, **kwargs):
        if self.hora_inicio and self.hora_fim:
            self.quantidade_horas = calcular_quantidade_horas(self.hora_inicio, self.hora_fim)
        self.full_clean()
        return super().save(*args, **kwargs)

    def __str__(self) -> str:
        return (
            f"{self.data} ({self.hora_inicio.strftime('%H:%M')}–"
            f"{self.hora_fim.strftime('%H:%M')} / {self.quantidade_horas}h)"
        )
