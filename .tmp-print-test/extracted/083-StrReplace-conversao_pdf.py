OLD:
            detalhe = log.read_text(encoding="utf-8", errors="replace").strip() if log.is_file() else ""
            if not detalhe:
                detalhe = (resultado.stderr or resultado.stdout or "").strip()

NEW:
            detalhe = (resultado.stdout or "").strip()
            if log.is_file():
                detalhe = (detalhe + "\n" + log.read_text(encoding="utf-8", errors="replace")).strip()
            if not detalhe:
                detalhe = (resultado.stderr or "").strip()