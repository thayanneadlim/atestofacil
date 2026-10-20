
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

function Get-NovosProcessos([string]$Nome, [int[]]$Anteriores) {
    $lista = @(Get-Process -Name $Nome -ErrorAction SilentlyContinue | Where-Object { $Anteriores -notcontains [int]$_.Id })
    return ,$lista
}

$antes = @(Get-Process -Name $processo -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
$app = $null
$doc = $null
try {
    if ($processo -eq "EXCEL") {
        $app = New-Object -ComObject Excel.Application
        Start-Sleep -Seconds 2
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $app.AskToUpdateLinks = $false
        $doc = $app.Workbooks.Open($Origem, 0, $true)
        $doc.ExportAsFixedFormat(0, $Destino)
    } else {
        $app = New-Object -ComObject Word.Application
        Start-Sleep -Seconds 2
        $app.Visible = $false
        $app.DisplayAlerts = 0
        $doc = $app.Documents.Open($Origem, $false, $true, $false)
        $doc.ExportAsFixedFormat($Destino, 17)
    }
    $doc.Close($false)
    $doc = $null
} catch {
    if (-not ((Test-Path -LiteralPath $Destino) -and (Get-Item -LiteralPath $Destino).Length -gt 0)) {
        Write-Output $_.Exception.Message
        exit 1
    }
} finally {
    if ($null -ne $doc) {
        try { $doc.Close($false) } catch {}
    }
    if ($null -ne $app) {
        try { $app.Quit() } catch {}
    }
    $limite = (Get-Date).AddSeconds(20)
    do {
        $novos = Get-NovosProcessos $processo $antes
        if ($novos.Count -eq 0) { break }
        Start-Sleep -Milliseconds 400
    } while ((Get-Date) -lt $limite)
    $novos = Get-NovosProcessos $processo $antes
    foreach ($item in $novos) {
        Stop-Process -Id $item.Id -Force -ErrorAction SilentlyContinue
    }
    if ($novos.Count -gt 0) {
        Start-Sleep -Seconds 8
    }
}
if (-not (Test-Path -LiteralPath $Destino) -or (Get-Item -LiteralPath $Destino).Length -eq 0) {
    Write-Output "PDF nao gerado"
    exit 1
}
exit 0
