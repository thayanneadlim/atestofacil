from core.services.templates_texto import extrair_n_doc_sem_ano

_CARACTERES_INVALIDOS = '\\:*?"<>|'


def nome_arquivo_download(
    prefixo: str, numero_documento: str, nome: str, extensao: str
) -> str:
    """
    Nome do arquivo ao baixar.

    CI/PP/RPA/ATESTO + número do documento sem o ano + nome da pessoa.
    Ex.: CI 001 - Maria Silva.docx
    """
    numero = extrair_n_doc_sem_ano(numero_documento or "")
    numero = "".join(c for c in numero if c not in _CARACTERES_INVALIDOS).strip()
    pessoa = " ".join((nome or "").split())
    pessoa = "".join(c for c in pessoa if c not in _CARACTERES_INVALIDOS).strip()
    extensao = (extensao or "").lstrip(".")
    return f"{prefixo} {numero} - {pessoa}.{extensao}"
