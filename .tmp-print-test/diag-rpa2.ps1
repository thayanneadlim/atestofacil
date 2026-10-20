$ErrorActionPreference = "Stop"
$origem = "D:\AtestoFacil\.tmp-print-test\03-rpa.xlsx"
$destino = "D:\AtestoFacil\.tmp-print-test\pdf\diag-rpa2.pdf"
function Passo($nome, $bloco) {
    try { & $bloco; Write-Output "OK $nome" } catch { Write-Output "FAIL $nome :: $($_.Exception.Message)"; throw }
}
try {
    Passo "new" { $script:app = New-Object -ComObject Excel.Application }
    Passo "visible" { $script:app.Visible = $false }
    Passo "alerts" { $script:app.DisplayAlerts = $false }
    Passo "open" { $script:doc = $script:app.Workbooks.Open($origem, 0, $true) }
    Passo "export" { $script:doc.ExportAsFixedFormat(0, $destino) }
    Passo "close" { $script:doc.Close($false) }
    try { $script:app.Quit(); Write-Output "OK quit" } catch { Write-Output "FAIL quit :: $($_.Exception.Message)" }
} catch { Write-Output "STOPPED" }
finally {
    Start-Sleep -Seconds 1
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
if (Test-Path $destino) { Write-Output ("SIZE " + (Get-Item $destino).Length) } else { Write-Output "NOFILE" }
