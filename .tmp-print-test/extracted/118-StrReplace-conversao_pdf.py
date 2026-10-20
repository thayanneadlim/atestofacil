OLD:
function Suspend-Recovery([string]$App) {
    $base = "HKCU:\Software\Microsoft\Office\16.0\$App\Resiliency"
    $recovery = Join-Path $base "DocumentRecovery"
    $backup = Join-Path $base "DocumentRecoveryAtivo"
    if ((Test-Path -LiteralPath $backup) -and -not (Test-Path -LiteralPath $recovery)) {
        Rename-Item -LiteralPath $backup -NewName "DocumentRecovery"
    }
    $moved = $false
    if ((Test-Path -LiteralPath $recovery) -and -not (Test-Path -LiteralPath $backup)) {
        Rename-Item -LiteralPath $recovery -NewName "DocumentRecoveryAtivo"
        $moved = $true
    }
    return New-Object psobject -Property @{
        Moved = $moved
        Recovery = $recovery
        Backup = $backup
    }
}

NEW:
function Suspend-Recovery([string]$App) {
    return New-Object psobject -Property @{ Moved = $false }
}