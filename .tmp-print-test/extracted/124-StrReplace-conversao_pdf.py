OLD:
            Start-Sleep -Milliseconds 600
            Stop-Owned $owned
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            Resume-Recovery $recuperacao
        }
        Start-Sleep -Seconds 1
    }

NEW:
            Start-Sleep -Milliseconds 600
            Stop-Owned $owned
            for ($espera = 0; $espera -lt 20; $espera++) {
                $vivo = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $owned -contains [int]$_.Id })
                if ($vivo.Count -eq 0) { break }
                Stop-Owned $owned
                Start-Sleep -Milliseconds 300
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            Resume-Recovery $recuperacao
        }
        Start-Sleep -Seconds 1
    }