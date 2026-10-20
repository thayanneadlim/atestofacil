from django.db import connection, transaction

from core.models import SequenciaAnual


def gerar_proximo_numero(tipo: str, ano: int | None = None) -> tuple[int, int, str]:
    """
    Gera o próximo número sequencial por tipo/ano com bloqueio de linha.

    Retorna (numero, ano, numero_formatado) no formato 001/2026.
    Em PostgreSQL usa SELECT ... FOR UPDATE.
    """
    from django.utils import timezone

    if ano is None:
        ano = timezone.localdate().year

    with transaction.atomic():
        try:
            seq = (
                SequenciaAnual.objects.select_for_update()
                .get(tipo=tipo, ano=ano)
            )
        except SequenciaAnual.DoesNotExist:
            SequenciaAnual.objects.create(tipo=tipo, ano=ano, ultimo_numero=0)
            seq = (
                SequenciaAnual.objects.select_for_update()
                .get(tipo=tipo, ano=ano)
            )
        seq.ultimo_numero += 1
        seq.save(update_fields=["ultimo_numero"])
        numero = seq.ultimo_numero
        numero_formatado = f"{numero:03d}/{ano}"
        return numero, ano, numero_formatado


def obter_ultimo_numero(tipo: str, ano: int | None = None) -> int:
    """Último número gravado na sequência do ano. Zero quando ela ainda não existe."""
    from django.utils import timezone

    if ano is None:
        ano = timezone.localdate().year
    seq = SequenciaAnual.objects.filter(tipo=tipo, ano=ano).first()
    return seq.ultimo_numero if seq else 0


def definir_ultimo_numero(tipo: str, ultimo_numero: int, ano: int | None = None) -> SequenciaAnual:
    """
    Grava o último número da sequência.

    O próximo documento desse tipo será esse valor + 1.
    """
    from django.utils import timezone

    if ano is None:
        ano = timezone.localdate().year
    if ultimo_numero < 0:
        raise ValueError("O número não pode ser negativo.")

    with transaction.atomic():
        seq, _criada = SequenciaAnual.objects.get_or_create(
            tipo=tipo,
            ano=ano,
            defaults={"ultimo_numero": ultimo_numero},
        )
        seq = SequenciaAnual.objects.select_for_update().get(pk=seq.pk)
        seq.ultimo_numero = ultimo_numero
        seq.save(update_fields=["ultimo_numero"])
        return seq


def backend_suporta_concorrencia() -> bool:
    return connection.vendor == "postgresql"