"""Caminho e validação do arquivo de atesto assinado anexado ao serviço."""

from pathlib import Path

from django.core.exceptions import ValidationError

TAMANHO_MAXIMO_ATESTO = 10 * 1024 * 1024
EXTENSOES_ATESTO = {".pdf", ".png", ".jpg", ".jpeg"}


def caminho_atesto_assinado(instance, filename: str) -> str:
    extensao = Path(filename).suffix.lower()
    if extensao not in EXTENSOES_ATESTO:
        extensao = ""
    identificador = getattr(instance, "numero_documento", None) or instance.pk or "novo"
    return f"atestos_assinados/{identificador}{extensao}"


def validar_tamanho_atesto(arquivo) -> None:
    tamanho = getattr(arquivo, "size", None)
    if tamanho is not None and tamanho > TAMANHO_MAXIMO_ATESTO:
        raise ValidationError("O arquivo deve ter no máximo 10 MB.")
