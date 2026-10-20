OLD:
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

NEW:
try {
    $brutos = Get-Content -LiteralPath $Mapa -Encoding UTF8 -Raw | ConvertFrom-Json
    $planilhas = New-Object System.Collections.ArrayList
    $documentos = New-Object System.Collections.ArrayList
    foreach ($item in $brutos) {
        $extensao = [IO.Path]::GetExtension([string]$item.origem).ToLower()
        if ($extensao -eq ".xlsx") {
            [void]$planilhas.Add($item)
        } elseif ($extensao -eq ".docx" -or $extensao -eq ".doc") {
            [void]$documentos.Add($item)
        }
    }
    if ($planilhas.Count -gt 0) { Convert-Planilhas $planilhas }
    if ($documentos.Count -gt 0) { Convert-Documentos $documentos }
} catch {