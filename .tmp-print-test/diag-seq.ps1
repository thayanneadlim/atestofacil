$ErrorActionPreference = "Stop"
function Exportar($origem, $destino, $rotulo) {
    Write-Output "---- $rotulo ----"
    $app = $null
    try {
        $app = New-Object -ComObject Excel.Application
        Start-Sleep -Seconds 1
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $app.AskToUpdateLinks = $false
        $doc = $app.Workbooks.Open($origem, 0, $true)
        Start-Sleep -Seconds 1
        try { Write-Output ("STATE " + $app.CalculationState + " READY " + $app.Ready) } catch { Write-Output ("STATEFAIL " + $_.Exception.Message) }
        try {
            $doc.ExportAsFixedFormat(0, $destino)
            Write-Output "OK export"
        } catch {
            Write-Output ("FAIL export :: " + $_.Exception.Message)
            $titulos = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.MainWindowTitle })
            Write-Output ("TITULOS " + ($titulos -join " | "))
            try { Write-Output ("CAPTION " + $app.Caption) } catch {}
        }
        try { $doc.Close($false) } catch {}
        try { $app.Quit() } catch { Write-Output ("FAIL quit :: " + $_.Exception.Message) }
    } catch {
        Write-Output ("FAIL geral :: " + $_.Exception.Message)
    } finally {
        Start-Sleep -Seconds 2
        Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 2
    }
    if (Test-Path $destino) { Write-Output ("SIZE " + (Get-Item $destino).Length) } else { Write-Output "NOFILE" }
}
Exportar "D:\AtestoFacil\.tmp-print-test\01-previsao.xlsx" "D:\AtestoFacil\.tmp-print-test\pdf\a-previsao.pdf" "previsao-1"
Exportar "D:\AtestoFacil\.tmp-print-test\01-previsao.xlsx" "D:\AtestoFacil\.tmp-print-test\pdf\b-previsao.pdf" "previsao-2"
Exportar "D:\AtestoFacil\.tmp-print-test\03-rpa.xlsx" "D:\AtestoFacil\.tmp-print-test\pdf\c-rpa.pdf" "rpa-3"
