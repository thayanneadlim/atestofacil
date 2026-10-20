OLD:
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

NEW:
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