OLD:
            detalhe = (resultado.stdout or "").strip()
            if log.is_file():
                detalhe = (detalhe + "\n" + log.read_text(encoding="utf-8", errors="replace")).strip()

NEW:
            detalhe = (resultado.stdout or "").replace("\ufeff", "").strip()
            if log.is_file():
                texto_log = log.read_text(encoding="utf-8", errors="replace").replace("\ufeff", "").strip()
                detalhe = (detalhe + "\n" + texto_log).strip()