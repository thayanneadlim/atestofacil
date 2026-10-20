OLD:
            linhas = [linha for linha in detalhe.splitlines() if linha.strip()]
            raise ConversaoPdfError(_mensagem(" | ".join(linhas[-8:])))

NEW:
            linha = detalhe.splitlines()[-1] if detalhe else ""
            raise ConversaoPdfError(_mensagem(linha))