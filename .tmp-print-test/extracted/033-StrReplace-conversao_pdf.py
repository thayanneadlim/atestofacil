OLD:
    destino.write_bytes(buffer.getvalue())


def pdf_tem_assinatura(conteudo: bytes) -> bool:
    """Confere o cabeçalho sem carregar o arquivo inteiro como planilha."""
    return zipfile.ZipFile is not None and conteudo.startswith(b"%PDF")


NEW:
    destino.write_bytes(buffer.getvalue())
