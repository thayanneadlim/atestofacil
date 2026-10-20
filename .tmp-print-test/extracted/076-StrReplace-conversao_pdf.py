OLD:
            $doc = $word.Documents.Open($origem, $false, $true, $false)
            $doc.ExportAsFixedFormat($destino, 17, $false, 0)
            $doc.Close($false)

NEW:
            $doc = $word.Documents.Open($origem, $false, $true, $false)
            $exportou = $false
            for ($i = 1; $i -le 8; $i++) {
                try {
                    $doc.ExportAsFixedFormat($destino, 17)
                    $exportou = $true
                    break
                } catch {
                    $hr = $_.Exception.HResult
                    if (($hr -ne -2147418111 -and $hr -ne -2146827862) -or $i -eq 8) { throw }
                    Start-Sleep -Milliseconds 500
                }
            }
            if (-not $exportou) { throw "PDF nao gerado para $origem" }
            $doc.Close($false)