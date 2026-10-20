OLD:
_SSML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


class ImpressaoError(Exception):
    pass


def _ocultar_aba_ir(conteudo: bytes) -> bytes:
    """A aba IR-2024 é tabela de apoio; a impressão fica só com o formulário RPA.

    A edição fica no XML para não descartar desenhos nem a configuração de página.
    """
    saida = BytesIO()
    with zipfile.ZipFile(BytesIO(conteudo)) as origem, zipfile.ZipFile(
        saida, "w", compression=zipfile.ZIP_DEFLATED
    ) as destino:
        for info in origem.infolist():
            dados = origem.read(info.filename)
            if info.filename == "xl/workbook.xml":
                raiz = etree.fromstring(dados)
                for planilha in raiz.findall(f"{{{_SSML}}}sheets/{{{_SSML}}}sheet"):
                    if planilha.get("name") == "IR-2024":
                        planilha.set("state", "hidden")
                dados = etree.tostring(
                    raiz, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            destino.writestr(info, dados)
    return saida.getvalue()

NEW:
class ImpressaoError(Exception):
    pass


_ABA_IR = re.compile(br"<sheet\b[^>]*name=\"IR-2024\"[^>]*/?>")


def _ocultar_aba_ir(conteudo: bytes) -> bytes:
    """A aba IR-2024 é tabela de apoio; a impressão fica só com o formulário RPA.

    Se ela já está oculta, o arquivo não é regravado. Reescrever o XML inteiro
    faz o Excel recusar a exportação para PDF.
    """
    with zipfile.ZipFile(BytesIO(conteudo)) as origem:
        if "xl/workbook.xml" not in origem.namelist():
            return conteudo
        xml = origem.read("xl/workbook.xml")
    marcador = _ABA_IR.search(xml)
    if marcador is None or b'state="hidden"' in marcador.group(0) or b"state='hidden'" in marcador.group(0):
        return conteudo
    tag = marcador.group(0)
    if b"state=" in tag:
        nova = re.sub(br'state="[^"]*"', b'state="hidden"', tag, count=1)
    else:
        nova = tag[:-2] + b' state="hidden"/>' if tag.endswith(b"/>") else tag[:-1] + b' state="hidden">'
    xml = xml[: marcador.start()] + nova + xml[marcador.end() :]
    saida = BytesIO()
    with zipfile.ZipFile(BytesIO(conteudo)) as origem, zipfile.ZipFile(
        saida, "w", compression=zipfile.ZIP_DEFLATED
    ) as destino:
        for info in origem.infolist():
            dados = xml if info.filename == "xl/workbook.xml" else origem.read(info.filename)
            destino.writestr(info, dados)
    return saida.getvalue()