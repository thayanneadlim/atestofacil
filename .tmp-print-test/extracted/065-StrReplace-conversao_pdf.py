OLD:
function Convert-Documentos($Itens) {
    $word = $null
    $owned = @()
    $attached = $false
    $recuperacao = Suspend-Recovery "Word"
    try {
        $before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
        $word = New-Object -ComObject Word.Application
        Find-Owned "WINWORD" $before
        $owned = @($script:ownedFound)
        if ($owned.Count -eq 0) {
            $attached = $true
            throw "Nao foi possivel abrir o Word em uma sessao separada."
        }
        foreach ($procId in $owned) { Add-OwnedPid $procId }
        Invoke-ComRetry { $word.Visible = $false } | Out-Null
        Invoke-ComRetry { $word.DisplayAlerts = 0 } | Out-Null
        try { $word.AutomationSecurity = 3 } catch {}
        foreach ($item in $Itens) {
            $doc = $null
            try {
                $origem = [string]$item.origem
                $destino = [string]$item.destino
                $doc = Invoke-ComRetry { $word.Documents.Open($origem, $false, $true, $false) }
                Invoke-ComRetry { $doc.ExportAsFixedFormat($destino, 17, $false, 0) } | Out-Null
            } finally {
                if ($null -ne $doc) {
                    try { Invoke-ComRetry { $doc.Close($false) } | Out-Null } catch {}
                    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch {}
                }
            }
            if (-not (Test-Path -LiteralPath $destino) -or (Get-Item -LiteralPath $destino).Length -eq 0) {
                throw "PDF nao gerado para $origem"
            }
        }
    } finally {
        if ($null -ne $word -and -not $attached) {
            try { $word.Quit() | Out-Null } catch {}
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
            Start-Sleep -Milliseconds 400
            Stop-Owned $owned
        } elseif ($null -ne $word) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Resume-Recovery $recuperacao
    }
}

NEW:
function Convert-Documentos($Itens) {
    $word = $null
    $owned = @()
    $recuperacao = Suspend-Recovery "Word"
    try {
        $before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        Start-Sleep -Milliseconds 300
        $owned = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id } | Where-Object { $before -notcontains $_ })
        foreach ($procId in $owned) { Add-OwnedPid $procId }
        foreach ($item in $Itens) {
            $origem = [string]$item.origem
            $destino = [string]$item.destino
            $doc = $word.Documents.Open($origem, $false, $true, $false)
            $doc.ExportAsFixedFormat($destino, 17, $false, 0)
            $doc.Close($false)
            if (-not (Test-Path -LiteralPath $destino) -or (Get-Item -LiteralPath $destino).Length -eq 0) {
                throw "PDF nao gerado para $origem"
            }
        }
    } finally {
        if ($null -ne $word) {
            try { $word.Quit() } catch {}
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {}
        }
        Start-Sleep -Milliseconds 400
        Stop-Owned $owned
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Resume-Recovery $recuperacao
    }
}