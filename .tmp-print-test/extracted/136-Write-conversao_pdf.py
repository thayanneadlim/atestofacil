"""Converte planilhas e documentos do Word em PDF e junta os arquivos.

Usa o Microsoft Office quando ele está instalado, porque a exportação
respeita área de impressão, escala, ajuste à página e margens do arquivo.
Cada arquivo abre numa sessão nova: o Excel recusa a exportação seguinte
se ela reutiliza o mesmo processo. Se o Office não estiver disponível,
usa o LibreOffice. A união dos PDFs fica no próprio Python.
"""

import os
import shutil
import subprocess
import sys
import tempfile
import time
from io import BytesIO
from pathlib import Path

from pypdf import PdfReader, PdfWriter


class ConversaoPdfError(Exception):
    pass


_MENSAGEM = "Não foi possível preparar os documentos para impressão."

_SCRIPT_OFFICE = r"""
param(
    [Parameter(Mandatory = $true)][string]$Origem,
    [Parameter(Mandatory = $true)][string]$Destino
)
$ErrorActionPreference = "Stop"
$extensao = [IO.Path]::GetExtension($Origem).ToLower()
if ($extensao -eq ".xlsx") {
    $processo = "EXCEL"
} else {
    $processo = "WINWORD"
}
$antes = @(Get-Process -Name $processo -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
$app = $null
try {
    if ($processo -eq "EXCEL") {
        $app = New-Object -ComObject Excel.Application
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $app.AskToUpdateLinks = $false
        $doc = $app.Workbooks.Open($Origem, 0, $true)
        $doc.ExportAsFixedFormat(0, $Destino)
        $doc.Close($false)
    } else {
        $app = New-Object -ComObject Word.Application
        $app.Visible = $false
        $app.DisplayAlerts = 0
        $doc = $app.Documents.Open($Origem, $false, $true, $false)
        $doc.ExportAsFixedFormat($Destino, 17)
        $doc.Close($false)
    }
    try { $app.Quit() } catch {}
} catch {
    Write-Output $_.Exception.Message
    exit 1
} finally {
    Start-Sleep -Milliseconds 500
    $novos = @(Get-Process -Name $processo -ErrorAction SilentlyContinue | Where-Object { $antes -notcontains [int]$_.Id })
    foreach ($item in $novos) {
        Stop-Process -Id $item.Id -Force -ErrorAction SilentlyContinue
    }
}
if (-not (Test-Path -LiteralPath $Destino) -or (Get-Item -LiteralPath $Destino).Length -eq 0) {
    Write-Output "PDF nao gerado"
    exit 1
}
exit 0
"""


def localizar_soffice() -> str | None:
    """Procura o LibreOffice no PATH e nos diretórios habituais."""
    encontrado = shutil.which("soffice") or shutil.which("soffice.exe")
    if encontrado:
        return encontrado
    candidatos: list[Path] = []
    if sys.platform == "win32":
        raizes = [
            os.environ.get("ProgramFiles", r"C:\Program Files"),
            os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
            os.environ.get("LOCALAPPDATA", ""),
        ]
        for raiz in raizes:
            if not raiz:
                continue
            base = Path(raiz)
            candidatos.append(base / "LibreOffice" / "program" / "soffice.exe")
            candidatos.append(base / "Programs" / "LibreOffice" / "program" / "soffice.exe")
    elif sys.platform == "darwin":
        candidatos.append(Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"))
    else:
        candidatos.extend(
            [
                Path("/usr/bin/soffice"),
                Path("/usr/lib/libreoffice/program/soffice"),
                Path("/snap/bin/libreoffice"),
            ]
        )
    for caminho in candidatos:
        if caminho.is_file():
            return str(caminho)
    return None


def _office_disponivel() -> bool:
    if sys.platform != "win32":
        return False
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"):
            pass
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"):
            pass
    except OSError:
        return False
    return True


def _flags_processo() -> int:
    if sys.platform == "win32":
        return subprocess.CREATE_NO_WINDOW
    return 0


def _mensagem(detalhe: str = "") -> str:
    texto = " ".join(detalhe.replace("\ufeff", "").split())
    if not texto:
        return _MENSAGEM
    return f"{_MENSAGEM} {texto[:400]}"


def _executar(comando: list[str], *, timeout: int) -> None:
    try:
        resultado = subprocess.run(
            comando,
            capture_output=True,
            text=True,
            encoding="cp1252" if sys.platform == "win32" else "utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
            stdin=subprocess.DEVNULL,
            creationflags=_flags_processo(),
        )
    except subprocess.TimeoutExpired as exc:
        raise ConversaoPdfError(_MENSAGEM) from exc
    except OSError as exc:
        raise ConversaoPdfError(_MENSAGEM) from exc
    if resultado.returncode != 0:
        detalhe = (resultado.stderr or resultado.stdout or "").strip()
        linha = detalhe.splitlines()[-1] if detalhe else ""
        raise ConversaoPdfError(_mensagem(linha))


def _exigir_pdfs(pares: list[tuple[Path, Path]]) -> None:
    ausentes = [
        origem.name
        for origem, destino in pares
        if not destino.is_file() or destino.stat().st_size == 0
    ]
    if ausentes:
        raise ConversaoPdfError(
            "Não foi possível converter todos os documentos: " + ", ".join(ausentes)
        )


def _limpar_parciais(pares: list[tuple[Path, Path]]) -> None:
    for _, destino in pares:
        if destino.exists():
            destino.unlink()


def _converter_um_office(script: Path, origem: Path, destino: Path) -> None:
    ultima = ""
    for _ in range(3):
        if destino.exists():
            destino.unlink()
        try:
            resultado = subprocess.run(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script),
                    "-Origem",
                    str(origem),
                    "-Destino",
                    str(destino),
                ],
                capture_output=True,
                text=True,
                encoding="cp1252",
                errors="replace",
                timeout=90,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise ConversaoPdfError(_MENSAGEM) from exc
        except OSError as exc:
            raise ConversaoPdfError(_MENSAGEM) from exc
        if resultado.returncode == 0 and destino.is_file() and destino.stat().st_size > 0:
            return
        texto = (resultado.stdout or resultado.stderr or "").strip()
        ultima = texto.splitlines()[-1] if texto else ""
        time.sleep(2)
    raise ConversaoPdfError(_mensagem(ultima))


def _converter_office(pares: list[tuple[Path, Path]]) -> None:
    with tempfile.TemporaryDirectory(prefix="office-impressao-") as pasta:
        script = Path(pasta) / "converter.ps1"
        script.write_bytes(b"\xef\xbb\xbf" + _SCRIPT_OFFICE.encode("ascii"))
        for origem, destino in pares:
            _converter_um_office(script, origem.resolve(), destino.resolve())


def _converter_libreoffice(pares: list[tuple[Path, Path]]) -> None:
    soffice = localizar_soffice()
    if not soffice:
        raise ConversaoPdfError(_MENSAGEM)
    with tempfile.TemporaryDirectory(prefix="lo-impressao-") as pasta:
        perfil = Path(pasta) / "profile"
        perfil.mkdir()
        saida = pares[0][1].parent
        _executar(
            [
                soffice,
                "--headless",
                "--norestore",
                "--nofirststartwizard",
                "--nologo",
                f"-env:UserInstallation={perfil.resolve().as_uri()}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(saida),
                *[str(origem) for origem, _ in pares],
            ],
            timeout=180,
        )


def converter_para_pdf(pares: list[tuple[Path, Path]]) -> None:
    """Gera um PDF para cada par (origem, destino), na mesma ordem."""
    if not pares:
        return
    erros: list[str] = []
    if _office_disponivel():
        try:
            _converter_office(pares)
            _exigir_pdfs(pares)
            return
        except ConversaoPdfError as exc:
            erros.append(str(exc))
            _limpar_parciais(pares)
    if localizar_soffice():
        try:
            _converter_libreoffice(pares)
            _exigir_pdfs(pares)
            return
        except ConversaoPdfError as exc:
            erros.append(str(exc))
    if erros:
        raise ConversaoPdfError(erros[-1])
    raise ConversaoPdfError(
        f"{_MENSAGEM} Este computador precisa do Microsoft Office ou do LibreOffice."
    )


def unir_pdfs(pdfs: list[Path], destino: Path) -> None:
    """Junta os PDFs na ordem recebida, preservando o tamanho de cada página."""
    if not pdfs:
        raise ConversaoPdfError("Nenhuma página foi gerada para impressão.")
    escritor = PdfWriter()
    try:
        for caminho in pdfs:
            leitor = PdfReader(str(caminho))
            if leitor.is_encrypted:
                leitor.decrypt("")
            if not leitor.pages:
                raise ConversaoPdfError(f"O arquivo {caminho.name} não tem páginas.")
            for pagina in leitor.pages:
                escritor.add_page(pagina)
        buffer = BytesIO()
        escritor.write(buffer)
    except ConversaoPdfError:
        raise
    except Exception as exc:
        raise ConversaoPdfError("Não foi possível juntar os documentos para impressão.") from exc
    destino.write_bytes(buffer.getvalue())
