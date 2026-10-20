$ErrorActionPreference = "Stop"
$origem = "D:\AtestoFacil\.tmp-print-test\01-previsao.xlsx"
$destino = "D:\AtestoFacil\.tmp-print-test\pdf\diag-previsao.pdf"
function Passo($nome, $bloco) {
    try {
        & $bloco
        Write-Output "OK $nome"
    } catch {
        Write-Output "FAIL $nome :: $($_.Exception.Message)"
        throw
    }
}
$app = $null
try {
    Passo "new" { $script:app = New-Object -ComObject Excel.Application }
    Start-Sleep -Seconds 2
    Passo "visible" { $script:app.Visible = $false }
    Passo "alerts" { $script:app.DisplayAlerts = $false }
    Passo "links" { $script:app.AskToUpdateLinks = $false }
    Passo "open" { $script:doc = $script:app.Workbooks.Open($origem, 0, $true) }
    Passo "export" { $script:doc.ExportAsFixedFormat(0, $destino) }
    Passo "close" { $script:doc.Close($false) }
    Passo "quit" { $script:app.Quit() }
} catch {
    Write-Output "STOPPED"
} finally {
    Start-Sleep -Milliseconds 400
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
if (Test-Path -LiteralPath $destino) { Write-Output ("SIZE " + (Get-Item -LiteralPath $destino).Length) } else { Write-Output "NOFILE" }
