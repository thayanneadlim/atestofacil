OLD:
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

NEW:
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

function Invoke-Com([scriptblock]$Acao) {
    $ultimo = $null
    for ($i = 0; $i -lt 20; $i++) {
        try {
            return (& $Acao)
        } catch {
            $ultimo = $_
            $msg = [string]$_.Exception.Message
            $ocupado = $msg.Contains("0x80010001") -or $msg.Contains("0x8001010A") -or $msg.Contains("0x800AC472")
            if (-not $ocupado) { throw }
            Start-Sleep -Milliseconds 600
        }
    }
    throw $ultimo
}

function Get-NovosProcessos([string]$Nome, [int[]]$Anteriores) {
    return @(Get-Process -Name $Nome -ErrorAction SilentlyContinue | Where-Object { $Anteriores -notcontains [int]$_.Id })
}

$antes = @(Get-Process -Name $processo -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
$app = $null
$doc = $null
try {
    if ($processo -eq "EXCEL") {
        $app = New-Object -ComObject Excel.Application
        Invoke-Com { $app.Visible = $false }
        Invoke-Com { $app.DisplayAlerts = $false }
        Invoke-Com { $app.AskToUpdateLinks = $false }
        $doc = Invoke-Com { $app.Workbooks.Open($Origem, 0, $true) }
        Invoke-Com { $doc.ExportAsFixedFormat(0, $Destino) }
    } else {
        $app = New-Object -ComObject Word.Application
        Invoke-Com { $app.Visible = $false }
        Invoke-Com { $app.DisplayAlerts = 0 }
        $doc = Invoke-Com { $app.Documents.Open($Origem, $false, $true, $false) }
        Invoke-Com { $doc.ExportAsFixedFormat($Destino, 17) }
    }
    Invoke-Com { $doc.Close($false) }
    $doc = $null
} catch {
    Write-Output $_.Exception.Message
    exit 1
} finally {
    if ($null -ne $doc) {
        try { $doc.Close($false) } catch {}
    }
    $novos = Get-NovosProcessos $processo $antes
    if ($novos.Count -gt 0 -and $null -ne $app) {
        try { $app.Quit() } catch {}
    }
    $app = $null
    $limite = (Get-Date).AddSeconds(15)
    do {
        $novos = Get-NovosProcessos $processo $antes
        if ($novos.Count -eq 0) { break }
        Start-Sleep -Milliseconds 400
    } while ((Get-Date) -lt $limite)
    foreach ($item in $novos) {
        Stop-Process -Id $item.Id -Force -ErrorAction SilentlyContinue
    }
    if ($novos.Count -gt 0) {
        Start-Sleep -Seconds 2
    }
    Start-Sleep -Milliseconds 800
}
if (-not (Test-Path -LiteralPath $Destino) -or (Get-Item -LiteralPath $Destino).Length -eq 0) {
    Write-Output "PDF nao gerado"
    exit 1
}
exit 0
"""