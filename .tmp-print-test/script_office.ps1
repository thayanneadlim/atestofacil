
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
    $lista = @(Get-Process -Name $Nome -ErrorAction SilentlyContinue | Where-Object { $Anteriores -notcontains [int]$_.Id })
    return ,$lista
}

$antes = @(Get-Process -Name $processo -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
$script:app = $null
$script:doc = $null
try {
    if ($processo -eq "EXCEL") {
        $script:app = New-Object -ComObject Excel.Application
        Invoke-Com { $script:app.Visible = $false }
        Invoke-Com { $script:app.DisplayAlerts = $false }
        Invoke-Com { $script:app.AskToUpdateLinks = $false }
        $script:doc = Invoke-Com { $script:app.Workbooks.Open($Origem, 0, $true) }
        Invoke-Com { $script:doc.ExportAsFixedFormat(0, $Destino) }
    } else {
        $script:app = New-Object -ComObject Word.Application
        Invoke-Com { $script:app.Visible = $false }
        Invoke-Com { $script:app.DisplayAlerts = 0 }
        $script:doc = Invoke-Com { $script:app.Documents.Open($Origem, $false, $true, $false) }
        Invoke-Com { $script:doc.ExportAsFixedFormat($Destino, 17) }
    }
    Invoke-Com { $script:doc.Close($false) }
    $script:doc = $null
} catch {
    Write-Output $_.Exception.Message
    exit 1
} finally {
    if ($null -ne $script:doc) {
        try { $script:doc.Close($false) } catch {}
    }
    $novos = Get-NovosProcessos $processo $antes
    if ($novos.Count -gt 0 -and $null -ne $script:app) {
        try { $script:app.Quit() } catch {}
    }
    $script:app = $null
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
