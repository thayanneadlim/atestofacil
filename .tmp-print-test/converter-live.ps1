
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

function Test-ComOcupado($Erro) {
    $msg = [string]$Erro.Exception.Message
    return $msg.Contains("0x80010001") -or $msg.Contains("0x8001010A") -or $msg.Contains("0x800AC472") -or $msg.Contains("rejeitada pelo chamado")
}

function Get-ProcessosNovos([string]$Nome, [int[]]$Anteriores) {
    return @(Get-Process -Name $Nome -ErrorAction SilentlyContinue | Where-Object { $Anteriores -notcontains [int]$_.Id })
}

$antes = @(Get-Process -Name $processo -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
$app = $null
$doc = $null
try {
    $ultimo = $null
    for ($tentativa = 0; $tentativa -lt 6 -and $null -eq $app; $tentativa++) {
        try {
            if ($processo -eq "EXCEL") {
                $app = New-Object -ComObject Excel.Application
                try { $app.IgnoreRemoteRequests = $true } catch {}
            } else {
                $app = New-Object -ComObject Word.Application
            }
            Write-Output "STEP new"
        } catch {
            $ultimo = $_
            if (-not (Test-ComOcupado $_)) { throw }
            Start-Sleep -Seconds 3
            $app = $null
        }
    }
    if ($null -eq $app) { throw $ultimo }

    Start-Sleep -Seconds 2
    if ($processo -eq "EXCEL") {
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $app.AskToUpdateLinks = $false
        try { $app.ScreenUpdating = $false } catch {}
        try { $app.EnableEvents = $false } catch {}
    } else {
        $app.Visible = $false
        $app.DisplayAlerts = 0
    }
    Write-Output "STEP props"

    $exportou = $false
    for ($tentativa = 0; $tentativa -lt 8 -and -not $exportou; $tentativa++) {
        try {
            if ($null -eq $doc) {
            if ($processo -eq "EXCEL") {
                Write-Output "STEP open"
                $doc = $app.Workbooks.Open($Origem, 0, $true)
            } else {
                Write-Output "STEP open"
                $doc = $app.Documents.Open($Origem, $false, $true, $false)
            }
            Write-Output "STEP opened"
            }
            if (Test-Path -LiteralPath $Destino) { Remove-Item -LiteralPath $Destino -Force }
            if ($processo -eq "EXCEL") {
                Write-Output "STEP export"
                $doc.ExportAsFixedFormat(0, $Destino)
            } else {
                Write-Output "STEP export"
                $doc.ExportAsFixedFormat($Destino, 17)
            }
            Write-Output "STEP exported"
            $exportou = $true
        } catch {
            $ultimo = $_
            if (-not (Test-ComOcupado $_)) { throw }
            Start-Sleep -Seconds 2
        }
    }
    if (-not $exportou) { throw $ultimo }
    try { $doc.Close($false) } catch {}
} catch {
    if (-not ((Test-Path -LiteralPath $Destino) -and (Get-Item -LiteralPath $Destino).Length -gt 0)) {
        Write-Output $_.Exception.Message
        exit 1
    }
} finally {
    if ($null -ne $doc) {
        try { $doc.Close($false) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch {}
        $doc = $null
    }
    if ($null -ne $app) {
        try { $app.DisplayAlerts = $false } catch {}
        try { $app.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
        $app = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $limite = (Get-Date).AddSeconds(15)
    do {
        $novos = Get-ProcessosNovos $processo $antes
        if ($novos.Count -eq 0) { break }
        Start-Sleep -Milliseconds 400
    } while ((Get-Date) -lt $limite)
    $novos = @(Get-ProcessosNovos $processo $antes)
    foreach ($item in $novos) {
        Stop-Process -Id $item.Id -Force -ErrorAction SilentlyContinue
    }
    if ($novos.Count -gt 0) {
        Start-Sleep -Seconds 12
    }
}
if (-not (Test-Path -LiteralPath $Destino) -or (Get-Item -LiteralPath $Destino).Length -eq 0) {
    Write-Output "PDF nao gerado"
    exit 1
}
exit 0
