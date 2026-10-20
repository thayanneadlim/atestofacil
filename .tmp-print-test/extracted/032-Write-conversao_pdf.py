"""Converte planilhas e documentos do Word em PDF e junta os arquivos.

Usa o Microsoft Office quando ele está instalado, porque a exportação
respeita área de impressão, escala, ajuste à página e margens do arquivo.
Se o Office não estiver disponível, usa o LibreOffice. A união dos PDFs
fica no próprio Python, sem depender de programas extras.
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path

from pypdf import PdfReader, PdfWriter


class ConversaoPdfError(Exception):
    pass


_MENSAGEM = "Não foi possível preparar os documentos para impressão."

_SCRIPT_OFFICE = r"""
param(
    [Parameter(Mandatory = $true)][string]$Mapa,
    [Parameter(Mandatory = $true)][string]$Pids,
    [Parameter(Mandatory = $true)][string]$Log
)

$ErrorActionPreference = "Stop"

function Write-Log([string]$Message) {
    Set-Content -LiteralPath $Log -Value $Message -Encoding utf8
}

function Add-OwnedPid([int]$ProcessId) {
    Add-Content -LiteralPath $Pids -Value $ProcessId -Encoding ascii
}

function Find-Owned([string]$Name, [int[]]$Before) {
    $script:ownedFound = @()
    for ($i = 0; $i -lt 40; $i++) {
        $now = @(Get-Process -Name $Name -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
        $script:ownedFound = @($now | Where-Object { $Before -notcontains $_ })
        if ($script:ownedFound.Count -gt 0) { return }
        Start-Sleep -Milliseconds 250
    }
}

function Test-ComBusy($ErrorRecord) {
    $ex = $ErrorRecord.Exception
    while ($null -ne $ex) {
        if ($ex.HResult -eq -2147418111 -or $ex.HResult -eq -2146827862) { return $true }
        $ex = $ex.InnerException
    }
    return $false
}

function Invoke-ComRetry {
    param([scriptblock]$Action)
    for ($try = 1; $try -le 12; $try++) {
        try {
            return (& $Action)
        } catch {
            if (-not (Test-ComBusy $_) -or $try -eq 12) { throw }
            Start-Sleep -Milliseconds (250 * $try)
        }
    }
}

function Stop-Owned {
    param([int[]]$Owned)
    foreach ($procId in $Owned) {
        Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
    }
}

function Convert-Planilhas($Itens) {
    if (@($Itens).Count -eq 0) { return }
    $excel = $null
    $owned = @()
    $attached = $false
    try {
        $before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
        $excel = New-Object -ComObject Excel.Application
        Find-Owned "EXCEL" $before
        $owned = @($script:ownedFound)
        if ($owned.Count -eq 0) {
            $attached = $true
            throw "Nao foi possivel abrir o Excel em uma sessao separada."
        }
        foreach ($procId in $owned) { Add-OwnedPid $procId }
        Invoke-ComRetry { $excel.Visible = $false } | Out-Null
        Invoke-ComRetry { $excel.DisplayAlerts = $false } | Out-Null
        Invoke-ComRetry { $excel.AskToUpdateLinks = $false } | Out-Null
        Invoke-ComRetry { $excel.ScreenUpdating = $false } | Out-Null
        try { $excel.AutomationSecurity = 3 } catch {}
        foreach ($item in @($Itens)) {
            $wb = $null
            try {
                $origem = $item.origem
                $destino = $item.destino
                $wb = Invoke-ComRetry { $excel.Workbooks.Open($origem, 0, $true) }
                Invoke-ComRetry {
                    foreach ($ws in @($wb.Worksheets)) {
                        if ($ws.Name -eq "IR-2024" -and $ws.Visible -ne 0 -and $ws.Visible -ne 2) {
                            $ws.Visible = 0
                        }
                    }
                } | Out-Null
                try { Invoke-ComRetry { $excel.Calculate() } | Out-Null } catch {}
                Invoke-ComRetry { $wb.ExportAsFixedFormat(0, $destino, 0, $true, $false) } | Out-Null
            } finally {
                if ($null -ne $wb) {
                    try { Invoke-ComRetry { $wb.Close($false) } | Out-Null } catch {}
                    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {}
                }
            }
            if (-not (Test-Path -LiteralPath $item.destino) -or (Get-Item -LiteralPath $item.destino).Length -eq 0) {
                throw "PDF nao gerado para $($item.origem)"
            }
        }
    } finally {
        if ($null -ne $excel -and -not $attached) {
            try { $excel.Quit() | Out-Null } catch {}
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
            Start-Sleep -Milliseconds 400
            Stop-Owned $owned
        } elseif ($null -ne $excel) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Convert-Documentos($Itens) {
    if (@($Itens).Count -eq 0) { return }
    $word = $null
    $owned = @()
    $attached = $false
    try {
        $before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
        $word = New-Object -ComObject Word.Application
        Find-Owned "WINWORD" $before
        $owned = @($script:ownedFound)
        if ($owned.Count -eq 0) {
            $attached = $true
            throw "Nao foi possivel abrir o Word em uma sessao separada."
        }
        foreach ($procId in $owned) { Add-OwnedPid $procId }
        Invoke-ComRetry { $word.Visible = $false } | Out-Null
        Invoke-ComRetry { $word.DisplayAlerts = 0 } | Out-Null
        try { $word.AutomationSecurity = 3 } catch {}
        foreach ($item in @($Itens)) {
            $doc = $null
            try {
                $origem = $item.origem
                $destino = $item.destino
                $doc = Invoke-ComRetry { $word.Documents.Open($origem, $false, $true, $false) }
                Invoke-ComRetry { $doc.ExportAsFixedFormat($destino, 17, $false, 0) } | Out-Null
            } finally {
                if ($null -ne $doc) {
                    try { Invoke-ComRetry { $doc.Close($false) } | Out-Null } catch {}
                    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch {}
                }
            }
            if (-not (Test-Path -LiteralPath $item.destino) -or (Get-Item -LiteralPath $item.destino).Length -eq 0) {
                throw "PDF nao gerado para $($item.origem)"
            }
        }
    } finally {
        if ($null -ne $word -and -not $attached) {
            try { $word.Quit() | Out-Null } catch {}
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
            Start-Sleep -Milliseconds 400
            Stop-Owned $owned
        } elseif ($null -ne $word) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

try {
    $itens = @(Get-Content -LiteralPath $Mapa -Encoding UTF8 -Raw | ConvertFrom-Json)
    $planilhas = @($itens | Where-Object { $_.origem.ToLower().EndsWith(".xlsx") })
    $documentos = @($itens | Where-Object {
        $extensao = [IO.Path]::GetExtension($_.origem).ToLower()
        $extensao -eq ".docx" -or $extensao -eq ".doc"
    })
    Convert-Planilhas $planilhas
    Convert-Documentos $documentos
} catch {
    $ex = $_.Exception
    while ($null -ne $ex.InnerException) { $ex = $ex.InnerException }
    Write-Log $ex.Message
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
    texto = " ".join(detalhe.split())
    if not texto:
        return _MENSAGEM
    return f"{_MENSAGEM} {texto[:400]}"


def _encerrar_pids(arquivo: Path) -> None:
    if not arquivo.is_file():
        return
    vistos: set[str] = set()
    for linha in arquivo.read_text(encoding="ascii", errors="ignore").splitlines():
        pid = linha.strip()
        if not pid.isdigit() or pid in vistos or pid == "0":
            continue
        vistos.add(pid)
        subprocess.run(
            ["taskkill", "/PID", pid, "/F"],
            capture_output=True,
            check=False,
            stdin=subprocess.DEVNULL,
            creationflags=_flags_processo(),
        )


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


def _converter_office(pares: list[tuple[Path, Path]]) -> None:
    with tempfile.TemporaryDirectory(prefix="office-impressao-") as pasta:
        base = Path(pasta)
        script = base / "converter.ps1"
        mapa = base / "mapa.json"
        pids = base / "pids.txt"
        log = base / "erro.txt"
        script.write_bytes(b"\xef\xbb\xbf" + _SCRIPT_OFFICE.encode("ascii"))
        mapa.write_bytes(
            b"\xef\xbb\xbf"
            + json.dumps(
                [
                    {"origem": str(origem.resolve()), "destino": str(destino.resolve())}
                    for origem, destino in pares
                ]
            ).encode("utf-8")
        )
        pids.write_text("", encoding="ascii")
        comando = [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script),
            "-Mapa",
            str(mapa),
            "-Pids",
            str(pids),
            "-Log",
            str(log),
        ]
        try:
            resultado = subprocess.run(
                comando,
                capture_output=True,
                text=True,
                encoding="cp1252",
                errors="replace",
                timeout=180,
                check=False,
                stdin=subprocess.DEVNULL,
                creationflags=_flags_processo(),
            )
        except subprocess.TimeoutExpired as exc:
            raise ConversaoPdfError(_MENSAGEM) from exc
        except OSError as exc:
            raise ConversaoPdfError(_MENSAGEM) from exc
        finally:
            _encerrar_pids(pids)
        if resultado.returncode != 0:
            detalhe = log.read_text(encoding="utf-8", errors="replace").strip() if log.is_file() else ""
            if not detalhe:
                detalhe = (resultado.stderr or resultado.stdout or "").strip()
            linha = detalhe.splitlines()[-1] if detalhe else ""
            raise ConversaoPdfError(_mensagem(linha))


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


def pdf_tem_assinatura(conteudo: bytes) -> bool:
    """Confere o cabeçalho sem carregar o arquivo inteiro como planilha."""
    return zipfile.ZipFile is not None and conteudo.startswith(b"%PDF")
