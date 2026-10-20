OLD:
function Find-Owned([string]$Name, [int[]]$Before) {
    $script:ownedFound = @()
    for ($i = 0; $i -lt 40; $i++) {
        $now = @(Get-Process -Name $Name -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
        $script:ownedFound = @($now | Where-Object { $Before -notcontains $_ })
        if ($script:ownedFound.Count -gt 0) { return }
        Start-Sleep -Milliseconds 250
    }
}

function Test-ComBusy($ErrorRecord) {
    $ex = $ErrorRecord.Exception
    while ($null -ne $ex) {
        if ($ex.HResult -eq -2147418111 -or $ex.HResult -eq -2146827862) { return $true }
        $ex = $ex.InnerException
    }
    return $false
}

function Invoke-ComRetry {
    param([scriptblock]$Action)
    for ($try = 1; $try -le 12; $try++) {
        try {
            return (& $Action)
        } catch {
            if (-not (Test-ComBusy $_) -or $try -eq 12) { throw }
            Start-Sleep -Milliseconds (250 * $try)
        }
    }
}

function Stop-Owned {

NEW:
function Stop-Owned {